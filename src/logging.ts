// Request logging.
//
// One JSON line per HTTP request, written on response finish. What is in it is fixed by
// this file: the HTTP verb, the path, the status, the duration, and — for an MCP POST —
// the JSON-RPC method name and, on a tools/call, the tool name.
//
// What is deliberately absent: headers, request bodies, response bodies, query strings,
// and tool arguments. The Authorization header carries a bearer token (#23), tool
// arguments carry page and section ids, and a tool result carries the user's page
// content. A query string is left out because the MCP auth spec forbids a token there
// and #23 rejects one that arrives anyway — logging the query would put the rejected
// token in the log, which is the outcome the rule exists to prevent.
//
// The JSON-RPC method and the tool name are protocol-level names from a fixed set. They
// are the two fields that make a log line useful for telling a failing tool from a
// failing transport, and neither can hold user content.

import type { NextFunction, Request, RequestHandler, Response } from 'express';

/** Where a log line goes. Injected so a test can read the lines instead of stdout. */
export type LogSink = (line: string) => void;

const defaultSink: LogSink = (line) => {
  console.log(line);
};

/** The per-request fields the MCP route contributes, stashed on `res.locals`. */
export interface McpLogFields {
  readonly rpc?: string;
  readonly tool?: string;
}

const MCP_LOCALS_KEY = 'mcpLog';

/** Record the JSON-RPC method and tool name for the line the logger will write. */
export function setMcpLogFields(res: Response, fields: McpLogFields): void {
  res.locals[MCP_LOCALS_KEY] = fields;
}

function readMcpLogFields(res: Response): McpLogFields {
  const fields: unknown = res.locals[MCP_LOCALS_KEY];
  return typeof fields === 'object' && fields !== null ? (fields as McpLogFields) : {};
}

/** Express middleware writing one line per request once the response has finished. */
export function requestLogger(sink: LogSink = defaultSink): RequestHandler {
  return (req: Request, res: Response, next: NextFunction) => {
    const startedAt = process.hrtime.bigint();
    // Captured now, not on finish: Express rewrites req.url when a request enters a
    // mounted router, so by the time the response ends req.path reads '/'. req.path
    // rather than req.originalUrl because originalUrl carries the query string.
    const path = req.path;

    res.on('finish', () => {
      const durationMs = Number((process.hrtime.bigint() - startedAt) / 1_000_000n);
      const { rpc, tool } = readMcpLogFields(res);
      sink(
        JSON.stringify({
          event: 'request',
          method: req.method,
          path,
          status: res.statusCode,
          durationMs,
          ...(rpc === undefined ? {} : { rpc }),
          ...(tool === undefined ? {} : { tool }),
        }),
      );
    });

    next();
  };
}

/** How many methods of a JSON-RPC batch are named before the rest are counted. */
const MAX_LOGGED_BATCH_METHODS = 5;

/**
 * Read the loggable names out of a parsed JSON-RPC body.
 *
 * Only `method` and, for `tools/call`, `params.name` are touched. A batch is an array,
 * which is why this joins several method names; a long batch is truncated so one request
 * cannot write an unbounded log line.
 */
export function mcpLogFields(body: unknown): McpLogFields {
  const messages = Array.isArray(body) ? body : [body];
  const methods: string[] = [];
  let tool: string | undefined;

  for (const message of messages.slice(0, MAX_LOGGED_BATCH_METHODS)) {
    if (typeof message !== 'object' || message === null) continue;

    const method = 'method' in message ? message.method : undefined;
    if (typeof method !== 'string') continue;
    methods.push(method);

    if (method !== 'tools/call' || tool !== undefined) continue;
    const params: unknown = 'params' in message ? message.params : undefined;
    if (typeof params !== 'object' || params === null || !('name' in params)) continue;
    if (typeof params.name === 'string') tool = params.name;
  }

  if (messages.length > MAX_LOGGED_BATCH_METHODS) {
    methods.push(`+${messages.length - MAX_LOGGED_BATCH_METHODS} more`);
  }

  return {
    ...(methods.length === 0 ? {} : { rpc: methods.join(',') }),
    ...(tool === undefined ? {} : { tool }),
  };
}
