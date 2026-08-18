import test from 'node:test';
import assert from 'node:assert/strict';
import type { AddressInfo } from 'node:net';

import express from 'express';

import { mcpLogFields, requestLogger } from '../src/logging.ts';
import { MCP_PATH, mcpRouter } from '../src/mcp-server.ts';
import type { ToolDefinition } from '../src/mcp-tools.ts';

test('mcpLogFields reports the JSON-RPC method', () => {
  assert.deepEqual(mcpLogFields({ jsonrpc: '2.0', id: 1, method: 'tools/list' }), {
    rpc: 'tools/list',
  });
});

test('mcpLogFields reports the tool name on a tools/call', () => {
  const fields = mcpLogFields({
    jsonrpc: '2.0',
    id: 1,
    method: 'tools/call',
    params: { name: 'get_page_content', arguments: { pageId: 'page-1' } },
  });
  assert.deepEqual(fields, { rpc: 'tools/call', tool: 'get_page_content' });
});

test('mcpLogFields names the methods of a batch and counts the rest', () => {
  const batch = Array.from({ length: 8 }, (_, i) => ({ jsonrpc: '2.0', id: i, method: 'ping' }));
  assert.deepEqual(mcpLogFields(batch), { rpc: 'ping,ping,ping,ping,ping,+3 more' });
});

test('mcpLogFields survives a body that is not a JSON-RPC message', () => {
  assert.deepEqual(mcpLogFields(null), {});
  assert.deepEqual(mcpLogFields('nonsense'), {});
  assert.deepEqual(mcpLogFields({ method: 42 }), {});
});

const SECRET_ARGUMENT = 'the-page-body-nobody-should-log';

const TOOLS: ToolDefinition[] = [
  {
    name: 'echo',
    title: 'Echo',
    description: 'Returns a fixed string.',
    inputSchema: { type: 'object', properties: { text: { type: 'string' } } },
    handle: async () => ({ content: [{ type: 'text', text: SECRET_ARGUMENT }] }),
  },
];

test('a request log line carries no argument, no result, and no header', async (t) => {
  const lines: string[] = [];
  const app = express();
  app.use(requestLogger((line) => lines.push(line)));
  app.use(MCP_PATH, mcpRouter(TOOLS));

  const server = app.listen(0);
  t.after(() => server.close());
  await new Promise<void>((resolve) => server.once('listening', () => resolve()));
  const { port } = server.address() as AddressInfo;

  const res = await fetch(`http://127.0.0.1:${port}${MCP_PATH}?access_token=leaked-token`, {
    method: 'POST',
    headers: {
      'content-type': 'application/json',
      accept: 'application/json, text/event-stream',
      authorization: 'Bearer leaked-token',
    },
    body: JSON.stringify({
      jsonrpc: '2.0',
      id: 1,
      method: 'tools/call',
      params: { name: 'echo', arguments: { text: SECRET_ARGUMENT } },
    }),
  });
  await res.arrayBuffer();

  assert.equal(lines.length, 1);
  const line = lines[0] ?? '';

  // The two protocol names are the point of the line.
  const parsed = JSON.parse(line) as Record<string, unknown>;
  assert.equal(parsed['rpc'], 'tools/call');
  assert.equal(parsed['tool'], 'echo');
  assert.equal(parsed['path'], MCP_PATH);
  assert.equal(parsed['status'], 200);

  // Nothing else. The bearer token is rejected from a query string by #23, and logging
  // the query would put the rejected token in the log anyway.
  assert.ok(!line.includes(SECRET_ARGUMENT), `leaked tool argument or result: ${line}`);
  assert.ok(!line.includes('leaked-token'), `leaked bearer token: ${line}`);
  assert.ok(!line.includes('access_token'), `leaked query string: ${line}`);
});
