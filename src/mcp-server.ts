// The MCP JSON-RPC surface and the HTTP route it is served on.
//
// Transport is stateless Streamable HTTP. Every POST builds its own `Server` and its own
// transport, answers the one request, and tears both down; nothing survives to the next
// request. That is what `--max-instances=1` on Cloud Run needs, because an instance is
// replaced without warning and a client holding a session id against a dead instance has
// no way to recover.
//
// SSE is refused, in two places. `enableJsonResponse` makes a POST answer with a JSON
// body instead of opening a stream, and GET is answered 405 here rather than reaching the
// transport — the SDK's stateless mode still opens a standalone SSE stream on GET, and
// that stream is exactly what holds a Cloud Run instance alive and bills for idle time.
//
// This builds on the SDK's low-level `Server`, not on `McpServer`. `McpServer` installs
// its tools/list and tools/call handlers as a side effect of the first `registerTool`
// call, so a server with no tools registered answers tools/list with "method not found"
// instead of an empty list. Issue #14 ships with no tools and the empty list has to work.
// The low-level class also puts the tools/call error mapping in one place — see
// `toolErrorResult` in ./mcp-tools.ts — rather than behind the SDK's own wrapper.

import express, { Router, type NextFunction, type Request, type Response } from 'express';
import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StreamableHTTPServerTransport } from '@modelcontextprotocol/sdk/server/streamableHttp.js';
import type { Transport } from '@modelcontextprotocol/sdk/shared/transport.js';
import {
  CallToolRequestSchema,
  ErrorCode,
  ListToolsRequestSchema,
  McpError,
} from '@modelcontextprotocol/sdk/types.js';

import { mcpLogFields, setMcpLogFields } from './logging.ts';
import { indexTools, toolDescriptors, toolErrorResult, type ToolDefinition } from './mcp-tools.ts';
import { SERVICE_NAME, VERSION } from './version.ts';

/** The path the MCP endpoint is mounted at; #21's metadata documents this URL. */
export const MCP_PATH = '/mcp';

/**
 * Largest JSON-RPC request accepted. A tool call carries ids and, for the write tools in
 * #18, an HTML fragment; nothing legitimate approaches this. The cap is here so a body
 * cannot be used to exhaust the single instance's memory.
 */
const MAX_REQUEST_BODY = '1mb';

/**
 * JSON-RPC's implementation-defined server-error code. The SDK answers an unsupported
 * HTTP method with this same code, so a client sees one shape whether the refusal came
 * from here or from the transport.
 */
const JSONRPC_SERVER_ERROR = -32000;

/**
 * Build one MCP server over the given tools.
 *
 * Called per request. Construction touches no I/O — the tools were built once at startup
 * and are shared — so this costs an object and a Map.
 */
export function createMcpServer(tools: readonly ToolDefinition[]): Server {
  const byName = indexTools(tools);

  const server = new Server(
    { name: SERVICE_NAME, version: VERSION },
    // Declared up front rather than inferred from what is registered: a handler cannot
    // be installed for a capability that was not declared, and the tool list is allowed
    // to be empty. `listChanged: false` is the truth here — the list is fixed at startup.
    { capabilities: { tools: { listChanged: false } } },
  );

  server.setRequestHandler(ListToolsRequestSchema, () => ({ tools: toolDescriptors(tools) }));

  server.setRequestHandler(CallToolRequestSchema, async (request) => {
    const { name, arguments: args } = request.params;
    const tool = byName.get(name);
    // A name that is not in the list is a protocol fault, not a tool failure, so it is a
    // JSON-RPC error rather than an isError result. The client sent something tools/list
    // never offered.
    if (tool === undefined) {
      throw new McpError(ErrorCode.InvalidParams, `Unknown tool: ${name}`);
    }

    try {
      return await tool.handle(args ?? {});
    } catch (err) {
      // Every failure inside a tool — an expired refresh token, a page that is gone, a
      // document resvg rejects — comes back as a readable result the model can act on.
      // Nothing propagates as an unhandled rejection.
      return toolErrorResult(name, err);
    }
  });

  return server;
}

/**
 * The MCP route. Mounted at MCP_PATH by ./server.ts; everything under it is closed by
 * #23's bearer-token middleware, which goes in front of this router.
 */
export function mcpRouter(tools: readonly ToolDefinition[]): Router {
  const router = Router();

  router.post(
    '/',
    express.json({ limit: MAX_REQUEST_BODY }),
    async (req: Request, res: Response) => {
      // Read the loggable names before handing the body to the transport, so a request
      // that fails still says which method failed.
      setMcpLogFields(res, mcpLogFields(req.body));

      const server = createMcpServer(tools);
      const transport = new StreamableHTTPServerTransport({
        // Omitting sessionIdGenerator is what selects stateless mode. It is omitted
        // rather than set to undefined because exactOptionalPropertyTypes rejects the
        // explicit undefined that the SDK's own example passes.
        enableJsonResponse: true,
      });

      // The pair is per-request and owns no shared state, so it is closed as soon as the
      // response ends — including when the client disconnects mid-request.
      res.on('close', () => {
        void transport.close();
        void server.close();
      });

      try {
        // The cast is over a mismatch inside the SDK's own declarations: its `Transport`
        // interface declares `onclose` optional, while the class exposes it as
        // `(() => void) | undefined`, and exactOptionalPropertyTypes treats those as
        // different types. The object is the transport the SDK intends here.
        await server.connect(transport as unknown as Transport);
        // The body is already parsed by express.json above; the transport must be told,
        // or it would try to read a stream that has been consumed.
        await transport.handleRequest(req, res, req.body);
      } catch (err) {
        writeJsonRpcError(res, 500, ErrorCode.InternalError, 'Internal server error');
        console.error(`mcp request failed: ${errorLabel(err)}`);
      }
    },
  );

  // GET is where an SSE stream would be opened and DELETE is where a session would be
  // closed. Neither exists here. They are answered before the transport sees them.
  router.all('/', (req: Request, res: Response) => {
    res.setHeader('Allow', 'POST');
    writeJsonRpcError(res, 405, JSONRPC_SERVER_ERROR, `${req.method} is not supported`);
  });

  // A body express.json could not parse, or one over the size cap, reaches here as an
  // error. Without this it would leave Express's default HTML error page, which an MCP
  // client cannot read.
  router.use((err: unknown, _req: Request, res: Response, next: NextFunction) => {
    if (res.headersSent) return next(err);

    const type = typeof err === 'object' && err !== null && 'type' in err ? err.type : undefined;
    if (type === 'entity.too.large') {
      writeJsonRpcError(res, 413, ErrorCode.InvalidRequest, 'Request body is too large');
      return;
    }
    if (type === 'entity.parse.failed') {
      writeJsonRpcError(res, 400, ErrorCode.ParseError, 'Request body is not valid JSON');
      return;
    }
    next(err);
  });

  return router;
}

/**
 * A JSON-RPC error envelope with `id: null`. Used for the failures that happen before or
 * outside a request's own handling, where the request id is unknown or unparsed.
 */
function writeJsonRpcError(res: Response, status: number, code: number, message: string): void {
  res.status(status).json({ jsonrpc: '2.0', error: { code, message }, id: null });
}

/** The name of a thrown value, never its message: a message can carry a request body. */
function errorLabel(err: unknown): string {
  return err instanceof Error ? err.name : typeof err;
}
