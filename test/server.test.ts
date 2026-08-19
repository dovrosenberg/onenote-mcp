import test, { after } from 'node:test';
import assert from 'node:assert/strict';
import type { AddressInfo } from 'node:net';

import type { Config } from '../src/config.ts';
import { createApp } from '../src/server.ts';

const STUB_CONFIG: Config = {
  graph: { clientId: 'client-id', authority: 'https://login.microsoftonline.com/common' },
  firestore: { cacheDocumentPath: 'tokencache/msal', projectId: 'proj' },
  oauth: {
    clientId: 'mcp-client',
    clientSecret: 'mcp-secret',
    tokenSigningKey: 'x'.repeat(32),
    publicUrl: 'https://onenote-mcp.example.run.app',
  },
  server: { port: 0 },
};

// Port 0 asks the OS for an ephemeral port, so the test cannot collide with a dev
// server already bound to 8080.
const server = createApp(STUB_CONFIG).listen(0);
const ready = new Promise<void>((resolve) => server.once('listening', () => resolve()));

after(() => {
  server.close();
});

async function get(path: string): Promise<Response> {
  await ready;
  const { port } = server.address() as AddressInfo;
  return fetch(`http://127.0.0.1:${port}${path}`);
}

test('GET /healthz returns 200 with the service identity', async () => {
  const res = await get('/healthz');
  assert.equal(res.status, 200);

  const body = (await res.json()) as Record<string, unknown>;
  assert.equal(body['status'], 'ok');
  assert.equal(body['service'], 'onenote-mcp');
  assert.match(String(body['version']), /^\d+\.\d+\.\d+$/);
});

test('GET / returns 404 — no catch-all is exposed', async () => {
  const res = await get('/');
  assert.equal(res.status, 404);
});

test('/healthz leaks no configuration', async () => {
  const body = (await (await get('/healthz')).json()) as Record<string, unknown>;

  // The service deploys with --allow-unauthenticated, so this response is public.
  // Guards against a later phase adding config echo to the health endpoint.
  for (const key of Object.keys(body)) {
    assert.doesNotMatch(key, /secret|key|client_?id|token|authority/i, `leaked key: ${key}`);
  }

  const serialised = JSON.stringify(body);
  for (const secret of ['mcp-secret', 'client-id', 'mcp-client', 'x'.repeat(32), 'proj']) {
    assert.ok(!serialised.includes(secret), `leaked value: ${secret}`);
  }
});
