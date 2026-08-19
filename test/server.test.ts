import test, { after } from 'node:test';
import assert from 'node:assert/strict';
import { createHmac } from 'node:crypto';
import type { AddressInfo } from 'node:net';

import express from 'express';

import type { Config } from '../src/config.ts';
import { MCP_PATH } from '../src/mcp-server.ts';
import { CONSENT_PATH } from '../src/oauth-provider.ts';
import { AUTHORIZATION_SERVER_METADATA_PATH } from '../src/oauth-router.ts';
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

// ---------------------------------------------------------------------------
// The bearer gate (issue #23).
//
// Every assertion here is over the wire, because the thing being tested is what an
// unauthenticated caller gets: the status, the WWW-Authenticate header, and whether a
// route answered at all. A direct call to the middleware shows none of that.
//
// The tokens are signed here rather than obtained by driving the OAuth flow, and the
// format is re-implemented rather than imported, for the reason test/oauth-provider.test.ts
// gives: a test that signs with the implementation's own function proves only that it
// agrees with itself.
// ---------------------------------------------------------------------------

const SIGNING_KEY = STUB_CONFIG.oauth?.tokenSigningKey ?? '';
const RESOURCE = `${STUB_CONFIG.oauth?.publicUrl ?? ''}${MCP_PATH}`;

function sign(payload: Record<string, string | number>, key: string = SIGNING_KEY): string {
  const body = Buffer.from(JSON.stringify(payload)).toString('base64url');
  return `${body}.${createHmac('sha256', key).update(body).digest('base64url')}`;
}

function accessToken(overrides: Record<string, string | number> = {}, key?: string): string {
  return sign(
    {
      k: 'a',
      c: STUB_CONFIG.oauth?.clientId ?? '',
      a: RESOURCE,
      p: 'offline_access',
      x: Math.floor(Date.now() / 1000) + 3600,
      n: 'nonce',
      ...overrides,
    },
    key,
  );
}

/** A JSON-RPC POST to the MCP endpoint, with whatever headers the caller supplies. */
async function mcpPost(headers: Record<string, string> = {}, path: string = MCP_PATH): Promise<Response> {
  await ready;
  const { port } = server.address() as AddressInfo;
  return fetch(`http://127.0.0.1:${port}${path}`, {
    method: 'POST',
    headers: {
      'content-type': 'application/json',
      accept: 'application/json, text/event-stream',
      ...headers,
    },
    body: JSON.stringify({ jsonrpc: '2.0', id: 1, method: 'tools/list', params: {} }),
  });
}

test('an unauthenticated MCP request is 401 and says where the metadata is', async () => {
  const res = await mcpPost();
  assert.equal(res.status, 401);

  // The three parameters Claude reads. Without resource_metadata the 401 is a dead end:
  // it is how the client finds the authorization server and starts the flow.
  const challenge = res.headers.get('www-authenticate') ?? '';
  assert.match(challenge, /^Bearer /);
  assert.match(challenge, /error="invalid_token"/);
  assert.match(challenge, /error_description="[^"]+"/);
  assert.match(challenge, /resource_metadata="[^"]+"/);
});

test('the resource_metadata URL in the 401 serves the protected-resource document', async () => {
  await ready;
  const { port } = server.address() as AddressInfo;
  const challenge = (await mcpPost()).headers.get('www-authenticate') ?? '';
  const advertised = /resource_metadata="([^"]+)"/.exec(challenge)?.[1];
  assert.ok(advertised !== undefined, `no resource_metadata in: ${challenge}`);

  // The host in that URL is MCP_PUBLIC_URL, which is not where this test server is
  // listening; the path is the part that has to be right.
  const res = await fetch(`http://127.0.0.1:${port}${new URL(advertised).pathname}`);
  assert.equal(res.status, 200);

  const doc = (await res.json()) as Record<string, unknown>;
  assert.equal(doc['resource'], RESOURCE);
});

test('a valid token reaches the MCP router', async () => {
  const res = await mcpPost({ authorization: `Bearer ${accessToken()}` });
  assert.equal(res.status, 200);

  const body = (await res.json()) as { result: { tools: unknown[] } };
  assert.ok(body.result.tools.length > 0);
});

test('a token in the query string is not honoured', async () => {
  // The middleware reads the Authorization header and nothing else, so this is 401
  // already. The test is here so that nothing later adds query parsing: the MCP auth
  // spec forbids a token in a query string, and src/logging.ts leaves the query out of
  // the log line on the assumption that one arriving is refused rather than used.
  const res = await mcpPost({}, `${MCP_PATH}?access_token=${accessToken()}`);
  assert.equal(res.status, 401);
});

test('a token minted for another resource server is refused', async () => {
  const foreign = accessToken({ a: 'https://someone-else.example.run.app/mcp' });
  const res = await mcpPost({ authorization: `Bearer ${foreign}` });

  // Right key, right shape, right expiry, wrong audience. The SDK's middleware compares
  // nothing here — this is the check src/oauth-provider.ts adds.
  assert.equal(res.status, 401);
});

test('an expired token is refused', async () => {
  const expired = accessToken({ x: Math.floor(Date.now() / 1000) - 60 });
  const res = await mcpPost({ authorization: `Bearer ${expired}` });
  assert.equal(res.status, 401);
});

test('a token signed with another key is refused', async () => {
  const forged = accessToken({}, 'y'.repeat(32));
  const res = await mcpPost({ authorization: `Bearer ${forged}` });
  assert.equal(res.status, 401);
});

test('a refresh token is not accepted as an access token', async () => {
  const refresh = accessToken({ k: 'r' });
  const res = await mcpPost({ authorization: `Bearer ${refresh}` });
  assert.equal(res.status, 401);
});

test('the 401 repeats nothing of the token it refused', async () => {
  const token = accessToken({ x: Math.floor(Date.now() / 1000) - 60 });
  const res = await mcpPost({ authorization: `Bearer ${token}` });

  // A 401 body reaches a client that may log it, and this one is generated by the SDK
  // rather than by this repository. The header is checked for the same reason.
  const body = await res.text();
  assert.ok(!body.includes(token), `the body repeats the token: ${body}`);
  assert.ok(
    !(res.headers.get('www-authenticate') ?? '').includes(token),
    'the WWW-Authenticate header repeats the token',
  );
});

// ---------------------------------------------------------------------------
// Fail closed: every registered route either is on the exempt list or needs a token.
// ---------------------------------------------------------------------------

/**
 * The routes that answer without a bearer token, and why each one has to.
 *
 * The list is longer than "everything except /mcp" because `mcpAuthRouter` mounts the
 * whole authorization flow at the application root, and none of it can be closed: a
 * client reaches those URLs precisely because it does not hold a token yet.
 */
const EXEMPT_PATHS = new Set([
  // Cloud Run's health check. Reports no configuration — see the test above.
  '/healthz',
  // Discovery. Read before a client holds anything.
  AUTHORIZATION_SERVER_METADATA_PATH,
  `/.well-known/oauth-protected-resource${MCP_PATH}`,
  // The authorization flow itself, and the consent form's POST target.
  '/authorize',
  '/token',
  CONSENT_PATH,
]);

interface RegisteredRoute {
  readonly method: string;
  readonly path: string;
}

/**
 * Every route `createApp` registers, with the path it is reachable at.
 *
 * Express 5 keeps no path string on a `use` layer — the mount is compiled into a matcher
 * function and the original string is dropped — so a mounted router's prefix cannot be
 * read back off the stack. The prefixes are recorded as they are registered instead, by
 * wrapping `Router.prototype.use` for the duration of the call. Leaf routes do keep their
 * own path, on `layer.route`, so only the prefixes need recording.
 *
 * This reaches into Express's internals, which is the cost of the property being tested:
 * a route added in a later issue has to show up here without anyone remembering to add
 * it to a list.
 */
function registeredRoutes(config: Config): RegisteredRoute[] {
  const routerClass = express.Router as unknown as {
    prototype: { use: (this: unknown, ...args: unknown[]) => unknown };
  };
  const original = routerClass.prototype.use;
  const mounts = new Map<unknown, { path: string; handler: unknown }[]>();

  routerClass.prototype.use = function record(this: unknown, ...args: unknown[]): unknown {
    const path = typeof args[0] === 'string' ? args[0] : '/';
    for (const handler of args) {
      if (typeof handler === 'string') continue;
      const siblings = mounts.get(this) ?? [];
      siblings.push({ path, handler });
      mounts.set(this, siblings);
    }
    return original.apply(this, args);
  };

  let app;
  try {
    app = createApp(config);
  } finally {
    routerClass.prototype.use = original;
  }

  const routes: RegisteredRoute[] = [];
  const seen = new Set<unknown>();

  const walk = (router: unknown, prefix: string): void => {
    // A router is a callable function carrying a `stack`, not a plain object.
    if (router === null || (typeof router !== 'object' && typeof router !== 'function')) return;
    if (seen.has(router)) return;
    seen.add(router);

    const stack = (router as { stack?: unknown[] }).stack ?? [];
    for (const entry of stack) {
      const route = (entry as { route?: { path?: string; methods?: Record<string, boolean> } }).route;
      if (route === undefined) continue;
      for (const method of Object.keys(route.methods ?? {})) {
        routes.push({ method, path: joinPath(prefix, route.path ?? '/') });
      }
    }

    for (const { path, handler } of mounts.get(router) ?? []) {
      // A middleware function has no stack of its own and registers nothing.
      const child = handler as { stack?: unknown[]; router?: unknown };
      if (child.stack !== undefined) walk(child, joinPath(prefix, path));
      else if (child.router !== undefined) walk(child.router, joinPath(prefix, path));
    }
  };

  walk((app as unknown as { router: unknown }).router, '');
  return routes;
}

/**
 * Join a mount prefix to a route path.
 *
 * A router mounted at the application root contributes a '/' prefix and every leaf route
 * inside one carries its own leading slash, so both the repeated separators and the
 * trailing one are removed — '/consent' has to compare equal to the path a request is
 * sent to.
 */
function joinPath(prefix: string, path: string): string {
  const joined = `${prefix}${path}`.replace(/\/{2,}/g, '/').replace(/\/+$/, '');
  return joined === '' ? '/' : joined;
}

test('every route the exempt list does not name requires a bearer token', async () => {
  await ready;
  const { port } = server.address() as AddressInfo;
  const routes = registeredRoutes(STUB_CONFIG);

  // The enumeration itself has to work, or this test passes by finding nothing.
  assert.ok(routes.length >= EXEMPT_PATHS.size, `only ${routes.length} routes found`);
  for (const exempt of EXEMPT_PATHS) {
    assert.ok(
      routes.some((route) => route.path === exempt),
      `the exempt list names ${exempt}, which is not a registered route`,
    );
  }

  const closed = routes.filter((route) => !EXEMPT_PATHS.has(route.path));
  assert.ok(closed.length > 0, 'no closed route was found — the gate is not mounted');

  for (const route of closed) {
    // `_all` is a router.all() registration; any verb reaches it.
    const method = route.method === '_all' ? 'GET' : route.method.toUpperCase();
    const res = await fetch(`http://127.0.0.1:${port}${route.path}`, { method });
    await res.arrayBuffer();

    assert.equal(res.status, 401, `${method} ${route.path} answered ${res.status} with no token`);
    assert.match(
      res.headers.get('www-authenticate') ?? '',
      /^Bearer /,
      `${method} ${route.path} answered 401 without a WWW-Authenticate challenge`,
    );
  }
});
