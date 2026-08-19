// Drives the OAuth mount over a real HTTP server through createApp, the way
// test/mcp-server.test.ts does. The assertions are about two public documents and the
// routes they advertise, so they are only meaningful over the wire: the paths come from
// the issuer URL rather than from a mount point, and getting that wrong produces a
// document that reads correctly and a 404 at every URL it names.
//
// Nothing here can confirm that Claude accepts the documents. That waits for a real
// connect against the deployed URL.

import test, { after } from 'node:test';
import assert from 'node:assert/strict';
import type { AddressInfo } from 'node:net';

import type { Config } from '../src/config.ts';
import { createApp } from '../src/server.ts';
import { createClientsStore, protectedResourceMetadataUrl } from '../src/oauth-router.ts';

const PUBLIC_URL = 'https://onenote-mcp.example.run.app';
const CLIENT_ID = 'mcp-client';
const CLIENT_SECRET = 'mcp-secret';
const SIGNING_KEY = 'x'.repeat(32);

const STUB_CONFIG: Config = {
  graph: { clientId: 'client-id', authority: 'https://login.microsoftonline.com/common' },
  firestore: { cacheDocumentPath: 'tokencache/msal', projectId: 'proj' },
  oauth: {
    clientId: CLIENT_ID,
    clientSecret: CLIENT_SECRET,
    tokenSigningKey: SIGNING_KEY,
    publicUrl: PUBLIC_URL,
  },
  server: { port: 0 },
};

const server = createApp(STUB_CONFIG).listen(0);
const ready = new Promise<void>((resolve) => server.once('listening', () => resolve()));

after(() => {
  server.close();
});

async function request(path: string, init: RequestInit = {}): Promise<Response> {
  await ready;
  const { port } = server.address() as AddressInfo;
  return fetch(`http://127.0.0.1:${port}${path}`, init);
}

async function metadata(path: string): Promise<Record<string, unknown>> {
  const res = await request(path);
  assert.equal(res.status, 200, `${path} answered ${res.status}`);
  return (await res.json()) as Record<string, unknown>;
}

test('the authorization-server metadata names this service as the issuer', async () => {
  const doc = await metadata('/.well-known/oauth-authorization-server');

  // The href, not the configured string: URL normalisation adds the trailing slash that
  // MCP_PUBLIC_URL is required not to carry. Every other URL in the document is built by
  // concatenating a path onto the same base, so they agree with each other by
  // construction.
  assert.equal(doc['issuer'], new URL(PUBLIC_URL).href);
  assert.equal(doc['authorization_endpoint'], `${PUBLIC_URL}/authorize`);
  assert.equal(doc['token_endpoint'], `${PUBLIC_URL}/token`);
});

test('the metadata advertises PKCE S256, both grant types, and offline_access', async () => {
  const doc = await metadata('/.well-known/oauth-authorization-server');

  // Claude sends PKCE on every authorization request and reads this to confirm it is
  // supported.
  assert.deepEqual(doc['code_challenge_methods_supported'], ['S256']);
  assert.deepEqual(doc['response_types_supported'], ['code']);

  const grants = doc['grant_types_supported'] as string[];
  assert.ok(grants.includes('authorization_code'), `grant types: ${grants.join(', ')}`);
  assert.ok(grants.includes('refresh_token'), `grant types: ${grants.join(', ')}`);

  // The switch that makes Claude ask for a refresh token. Without it the operator
  // re-consents every time an access token expires.
  const scopes = doc['scopes_supported'] as string[];
  assert.ok(scopes.includes('offline_access'), `scopes: ${scopes.join(', ')}`);
});

test('the metadata offers no registration endpoint', async () => {
  const doc = await metadata('/.well-known/oauth-authorization-server');

  // The clients store has no registerClient, which is what keeps this key absent, which
  // is what keeps Dynamic Client Registration out of the picture. The client id and
  // secret are configured instead.
  assert.equal('registration_endpoint' in doc, false);
});

test('the metadata does not claim to accept clients without a secret', async () => {
  const doc = await metadata('/.well-known/oauth-authorization-server');

  // The SDK's own document says ['client_secret_post', 'none'] unconditionally, which is
  // untrue of a server whose one client record carries a secret. This route is
  // registered ahead of the SDK's to rewrite that one field; the assertion is what
  // notices if the override stops being reached.
  assert.deepEqual(doc['token_endpoint_auth_methods_supported'], ['client_secret_post']);
});

test('the protected-resource metadata names the MCP endpoint, path included', async () => {
  const doc = await metadata('/.well-known/oauth-protected-resource/mcp');

  // The /mcp path is the point: this is the canonical resource identifier Claude sends
  // on every authorization and token request, and the audience issue #23 binds a token
  // to. An origin here would make a token minted for any path acceptable.
  assert.equal(doc['resource'], `${PUBLIC_URL}/mcp`);
  assert.deepEqual(doc['authorization_servers'], [new URL(PUBLIC_URL).href]);
});

test('protected-resource metadata is served only at the path-suffixed URL', async () => {
  // Measured against SDK 1.30.0 by the spike in issue #20 and asserted here so a later
  // SDK version that starts serving the bare path does not change behaviour unnoticed.
  // Claude probes the suffixed path first and issue #23's 401 names it explicitly.
  const res = await request('/.well-known/oauth-protected-resource');
  assert.equal(res.status, 404);
});

test('the metadata URL handed to issue #23 is the one that answers', async () => {
  const url = protectedResourceMetadataUrl(STUB_CONFIG.oauth!);
  assert.equal(url, `${PUBLIC_URL}/.well-known/oauth-protected-resource/mcp`);

  // The constant and the route cannot disagree: the 401's resource_metadata parameter is
  // built from the first and fetched from the second.
  const res = await request(new URL(url).pathname);
  assert.equal(res.status, 200);
});

test('every endpoint the metadata advertises exists', async () => {
  const doc = await metadata('/.well-known/oauth-authorization-server');

  for (const key of ['authorization_endpoint', 'token_endpoint']) {
    const path = new URL(doc[key] as string).pathname;
    const res = await request(path, { method: 'POST' });
    await res.arrayBuffer();

    // Not a 404 and not a 405: the route is mounted and answered. What it answers is
    // issue #22's business — the provider behind it refuses everything until then.
    assert.notEqual(res.status, 404, `${path} is not mounted`);
    assert.notEqual(res.status, 405, `${path} refuses POST`);
  }
});

test('neither document carries secret material', async () => {
  const documents = [
    JSON.stringify(await metadata('/.well-known/oauth-authorization-server')),
    JSON.stringify(await metadata('/.well-known/oauth-protected-resource/mcp')),
  ];

  // Both are public: the routes are unauthenticated by necessity, since discovery
  // happens before a client holds any token. Same shape as the /healthz test.
  for (const doc of documents) {
    for (const secret of [CLIENT_SECRET, SIGNING_KEY, 'client-id']) {
      assert.ok(!doc.includes(secret), `leaked value in metadata: ${secret}`);
    }
  }
});

test('the clients store holds one client, with the three Claude redirect URIs', async () => {
  const store = createClientsStore(STUB_CONFIG.oauth!);

  const client = await store.getClient(CLIENT_ID);
  assert.equal(client?.client_id, CLIENT_ID);
  assert.equal(client?.client_secret, CLIENT_SECRET);

  // claude.ai covers every hosted surface; the two loopback entries are Claude Code,
  // which binds an ephemeral port the SDK's matcher ignores. No port is registered here
  // for that reason.
  assert.deepEqual(client?.redirect_uris, [
    'https://claude.ai/api/mcp/auth_callback',
    'http://localhost/callback',
    'http://127.0.0.1/callback',
  ]);

  assert.equal(await store.getClient('someone-else'), undefined);

  // Its absence is what keeps registration_endpoint out of the metadata.
  assert.equal(store.registerClient, undefined);
});
