import test, { after } from 'node:test';
import assert from 'node:assert/strict';
import type { AddressInfo } from 'node:net';
import type { Server as HttpServer } from 'node:http';

import express from 'express';

import type { Config } from '../src/config.ts';
import { GraphRequestError } from '../src/graph-structure.ts';
import { MCP_PATH, mcpRouter } from '../src/mcp-server.ts';
import { requiredString, type ToolDefinition } from '../src/mcp-tools.ts';
import { createApp } from '../src/server.ts';

const STUB_CONFIG: Config = {
  graph: { clientId: 'client-id', authority: 'https://login.microsoftonline.com/common' },
  firestore: { cacheDocumentPath: 'tokencache/msal', projectId: 'proj' },
  oauth: { clientId: 'mcp-client', clientSecret: 'mcp-secret', tokenSigningKey: 'x'.repeat(32) },
  server: { port: 0 },
};

/** Both media types, which the Streamable HTTP transport requires of every POST. */
const MCP_HEADERS = {
  'content-type': 'application/json',
  accept: 'application/json, text/event-stream',
};

const INITIALIZE = {
  jsonrpc: '2.0',
  id: 1,
  method: 'initialize',
  params: {
    protocolVersion: '2025-06-18',
    capabilities: {},
    clientInfo: { name: 'test-client', version: '0.0.0' },
  },
};

// The service under test, wired to the real tool registry.
const serviceServer = createApp(STUB_CONFIG).listen(0);
const serviceReady = new Promise<void>((resolve) =>
  serviceServer.once('listening', () => resolve()),
);

after(() => {
  serviceServer.close();
  injectedServer.close();
});

async function mcp(body: unknown, init: RequestInit = {}): Promise<Response> {
  await serviceReady;
  const { port } = serviceServer.address() as AddressInfo;
  return fetch(`http://127.0.0.1:${port}${MCP_PATH}`, {
    method: 'POST',
    headers: MCP_HEADERS,
    body: typeof body === 'string' ? body : JSON.stringify(body),
    ...init,
  });
}

test('POST initialize is answered over Streamable HTTP', async () => {
  const res = await mcp(INITIALIZE);
  assert.equal(res.status, 200);

  // JSON, not text/event-stream: no stream is opened, so no Cloud Run instance is held.
  assert.match(res.headers.get('content-type') ?? '', /application\/json/);

  const body = (await res.json()) as { result: Record<string, unknown> };
  const info = body.result['serverInfo'] as Record<string, unknown>;
  assert.equal(info['name'], 'onenote-mcp');
  assert.ok((body.result['capabilities'] as Record<string, unknown>)['tools']);
});

test('the initialize response carries no session id', async () => {
  const res = await mcp(INITIALIZE);
  await res.arrayBuffer();

  // Stateless mode. A session id would be a promise this service cannot keep: the
  // instance answering the next request may be a different one.
  assert.equal(res.headers.get('mcp-session-id'), null);
});

test('tools/list answers on a request that never initialised', async () => {
  // The acceptance criterion for issue #14: a second request holds no state from the
  // first, so tools/list has to work on a server that saw no initialize of its own.
  const res = await mcp({ jsonrpc: '2.0', id: 2, method: 'tools/list', params: {} });
  assert.equal(res.status, 200);

  const body = (await res.json()) as { result: { tools: { name: string }[] } };
  assert.deepEqual(
    body.result.tools.map((tool) => tool.name),
    [
      'list_notebooks',
      'list_sections',
      'list_pages',
      'search_pages',
      'find_page_by_name',
      'list_pages_by_name',
      'get_page_content',
    ],
  );
});

test('GET is refused with 405 rather than opening an SSE stream', async () => {
  await serviceReady;
  const { port } = serviceServer.address() as AddressInfo;
  const res = await fetch(`http://127.0.0.1:${port}${MCP_PATH}`, {
    method: 'GET',
    headers: { accept: 'text/event-stream' },
  });

  assert.equal(res.status, 405);
  assert.equal(res.headers.get('allow'), 'POST');

  const body = (await res.json()) as { error: { code: number } };
  assert.equal(body.error.code, -32000);
});

test('DELETE is refused with 405', async () => {
  await serviceReady;
  const { port } = serviceServer.address() as AddressInfo;
  const res = await fetch(`http://127.0.0.1:${port}${MCP_PATH}`, { method: 'DELETE' });

  assert.equal(res.status, 405);
  assert.equal(res.headers.get('allow'), 'POST');
});

test('a body that is not JSON is refused as a JSON-RPC parse error', async () => {
  const res = await mcp('{ not json');
  assert.equal(res.status, 400);
  assert.match(res.headers.get('content-type') ?? '', /application\/json/);

  const body = (await res.json()) as { jsonrpc: string; error: { code: number } };
  assert.equal(body.jsonrpc, '2.0');
  assert.equal(body.error.code, -32700);
});

test('calling a tool that was never registered is a JSON-RPC error', async () => {
  const res = await mcp({
    jsonrpc: '2.0',
    id: 3,
    method: 'tools/call',
    params: { name: 'no_such_tool', arguments: {} },
  });
  assert.equal(res.status, 200);

  const body = (await res.json()) as { error: { code: number; message: string } };
  assert.equal(body.error.code, -32602);
  assert.match(body.error.message, /no_such_tool/);
});

// ---------------------------------------------------------------------------
// The same router with tools injected: what a tool failure looks like from outside.
// ---------------------------------------------------------------------------

const GRAPH_20266 = JSON.stringify({
  error: { code: '20266', message: 'maximum sections exceeded' },
});

const TOOLS: ToolDefinition[] = [
  {
    name: 'echo',
    title: 'Echo',
    description: 'Returns the text it was given.',
    inputSchema: {
      type: 'object',
      properties: { text: { type: 'string' } },
      required: ['text'],
    },
    handle: async (args) => ({ content: [{ type: 'text', text: requiredString(args, 'text') }] }),
  },
  {
    name: 'graph_failure',
    title: 'Graph failure',
    description: 'Always throws the Graph error a tool would.',
    inputSchema: { type: 'object', properties: {} },
    handle: () => {
      throw new GraphRequestError('https://graph.example/x', 400, 'Bad Request', GRAPH_20266);
    },
  },
  {
    name: 'surprise',
    title: 'Surprise',
    description: 'Always throws something this server does not model.',
    inputSchema: { type: 'object', properties: {} },
    handle: () => {
      throw new Error('page body: the quick brown fox');
    },
  },
];

const injected = express();
injected.use(MCP_PATH, mcpRouter(TOOLS));
const injectedServer = injected.listen(0);
const injectedReady = new Promise<void>((resolve) =>
  injectedServer.once('listening', () => resolve()),
);

async function callTool(name: string, args: Record<string, unknown>): Promise<CallResult> {
  await injectedReady;
  const { port } = injectedServer.address() as AddressInfo;
  const res = await fetch(`http://127.0.0.1:${port}${MCP_PATH}`, {
    method: 'POST',
    headers: MCP_HEADERS,
    body: JSON.stringify({ jsonrpc: '2.0', id: 9, method: 'tools/call', params: { name, arguments: args } }),
  });
  return (await res.json()) as CallResult;
}

interface CallResult {
  result?: { isError?: boolean; content: { type: string; text: string }[] };
  error?: { code: number; message: string };
}

test('tools/list advertises the registered tools and their input schemas', async () => {
  await injectedReady;
  const { port } = injectedServer.address() as AddressInfo;
  const res = await fetch(`http://127.0.0.1:${port}${MCP_PATH}`, {
    method: 'POST',
    headers: MCP_HEADERS,
    body: JSON.stringify({ jsonrpc: '2.0', id: 4, method: 'tools/list', params: {} }),
  });

  const body = (await res.json()) as {
    result: { tools: { name: string; inputSchema: { required?: string[] } }[] };
  };
  assert.deepEqual(
    body.result.tools.map((tool) => tool.name),
    ['echo', 'graph_failure', 'surprise'],
  );
  assert.deepEqual(body.result.tools[0]?.inputSchema.required, ['text']);
});

test('a tool that succeeds returns its content', async () => {
  const body = await callTool('echo', { text: 'hello' });
  assert.equal(body.result?.isError, undefined);
  assert.equal(body.result?.content[0]?.text, 'hello');
});

test('a bad argument is an isError result naming the argument', async () => {
  const body = await callTool('echo', {});
  assert.equal(body.result?.isError, true);
  assert.match(body.result?.content[0]?.text ?? '', /'text'/);
});

test('a Graph failure surfaces as a readable tool error, not a rejection', async () => {
  const body = await callTool('graph_failure', {});
  assert.equal(body.error, undefined, 'a Graph failure is a tool result, not a protocol error');
  assert.equal(body.result?.isError, true);

  const text = body.result?.content[0]?.text ?? '';
  assert.match(text, /400 Bad Request/);
  assert.match(text, /20266/);
});

test('an unmodelled error is reported without its message', async () => {
  const body = await callTool('surprise', {});
  assert.equal(body.result?.isError, true);

  // The message of an arbitrary error may quote a request body, and this text reaches a
  // client that may log it.
  const text = body.result?.content[0]?.text ?? '';
  assert.ok(!text.includes('quick brown fox'), `leaked message: ${text}`);
  assert.match(text, /unexpected error/);
});
