// The keepalive route, driven over a real HTTP server.
//
// It is exercised through its own router with a fake target rather than through
// `createApp`, because the success path has to call `refresh()` and the real one reaches
// Firestore and Entra. What `createApp` is used for here is the two facts a fake router
// cannot show: that the route is absent when no secret is configured, and that it sits
// outside the bearer gate — a scheduler holds a shared secret and cannot run an OAuth
// flow, so a keepalive request carrying no Authorization header must not be answered
// with the 401 that `/mcp` gives.
//
// What no test here covers is whether a forced refresh actually slides Microsoft's
// refresh-token window. That is a property of Entra, not of this code, and nothing
// confirms it until an operator watches the scheduler job run against the real tenant
// for longer than the window.

import test, { after } from 'node:test';
import assert from 'node:assert/strict';
import type { AddressInfo } from 'node:net';

import express from 'express';

import type { Config } from '../src/config.ts';
import { GraphAuthError } from '../src/graph-auth.ts';
import { KEEPALIVE_HEADER, KEEPALIVE_PATH, keepaliveRouter } from '../src/keepalive.ts';
import { setEventSink } from '../src/logging.ts';
import { createApp } from '../src/server.ts';

const SECRET = 's'.repeat(32);
const DOC_PATH = 'tokencache/msal';

// Every failure path below logs an operational event on purpose. Silenced so the run
// shows the assertions rather than the lines they expect.
setEventSink(() => {});

interface Recorder {
  calls: number;
}

/** A target that succeeds, counting the calls. */
function succeeding(recorder: Recorder): { refresh: () => Promise<string> } {
  return {
    refresh: () => {
      recorder.calls += 1;
      return Promise.resolve('fake-access-token');
    },
  };
}

/** A target that fails the way the real one does. */
function failing(reason: 'cache-unavailable' | 'silent-failed'): { refresh: () => Promise<string> } {
  return {
    refresh: () => Promise.reject(new GraphAuthError(reason, DOC_PATH)),
  };
}

function serve(router: express.Router): {
  post: (init?: RequestInit) => Promise<Response>;
  request: (method: string, init?: RequestInit) => Promise<Response>;
  close: () => void;
} {
  const app = express();
  app.use(KEEPALIVE_PATH, router);
  const server = app.listen(0);
  const ready = new Promise<void>((resolve) => server.once('listening', () => resolve()));

  const send = async (method: string, init: RequestInit = {}): Promise<Response> => {
    await ready;
    const { port } = server.address() as AddressInfo;
    return fetch(`http://127.0.0.1:${port}${KEEPALIVE_PATH}`, { method, ...init });
  };

  return {
    post: (init = {}) => send('POST', init),
    request: (method, init = {}) => send(method, init),
    close: () => server.close(),
  };
}

test('the right secret refreshes the token and answers 200', async () => {
  const recorder: Recorder = { calls: 0 };
  const target = serve(keepaliveRouter(SECRET, succeeding(recorder)));
  after(() => target.close());

  const res = await target.post({ headers: { [KEEPALIVE_HEADER]: SECRET } });

  assert.equal(res.status, 200);
  assert.deepEqual(await res.json(), { status: 'ok' });
  assert.equal(recorder.calls, 1);
  // The response carries no credential and must not be cached by anything in front.
  assert.equal(res.headers.get('cache-control'), 'no-store');
});

test('a wrong or missing secret is 401 and costs no refresh', async () => {
  const recorder: Recorder = { calls: 0 };
  const target = serve(keepaliveRouter(SECRET, succeeding(recorder)));
  after(() => target.close());

  const cases: Array<[string, RequestInit]> = [
    ['no header', {}],
    ['wrong value, same length', { headers: { [KEEPALIVE_HEADER]: 'w'.repeat(32) } }],
    ['a prefix of the secret', { headers: { [KEEPALIVE_HEADER]: SECRET.slice(0, 16) } }],
    ['the secret with a suffix', { headers: { [KEEPALIVE_HEADER]: `${SECRET}x` } }],
    ['empty', { headers: { [KEEPALIVE_HEADER]: '' } }],
  ];

  for (const [label, init] of cases) {
    const res = await target.post(init);
    assert.equal(res.status, 401, label);
    assert.deepEqual(await res.json(), { status: 'unauthorized' }, label);
  }

  // The point of the check running first: an unauthenticated flood costs a string
  // comparison, not a token request to Entra and a write to Firestore.
  assert.equal(recorder.calls, 0);
});

test('a failed refresh is 503 and says whether a retry can fix it', async () => {
  const unavailable = serve(keepaliveRouter(SECRET, failing('cache-unavailable')));
  after(() => unavailable.close());

  const res = await unavailable.post({ headers: { [KEEPALIVE_HEADER]: SECRET } });
  assert.equal(res.status, 503);

  const body = (await res.json()) as Record<string, unknown>;
  assert.equal(body['status'], 'failed');
  assert.equal(body['reason'], 'cache-unavailable');
  // The distinction the scheduler acts on: keep retrying, or stop and tell somebody.
  assert.equal(body['retryable'], true);
  assert.doesNotMatch(String(body['detail']), /npm run bootstrap/);
});

test('a dead refresh token is 503, not retryable, and names the CLI', async () => {
  const dead = serve(keepaliveRouter(SECRET, failing('silent-failed')));
  after(() => dead.close());

  const res = await dead.post({ headers: { [KEEPALIVE_HEADER]: SECRET } });
  assert.equal(res.status, 503);

  const body = (await res.json()) as Record<string, unknown>;
  assert.equal(body['reason'], 'silent-failed');
  assert.equal(body['retryable'], false);
  // This is the one failure a human has to act on, so the response says what to run.
  assert.match(String(body['detail']), /npm run bootstrap/);
});

test('an unexpected error is 503 and quotes nothing', async () => {
  const target = serve(
    keepaliveRouter(SECRET, {
      refresh: () => Promise.reject(new Error('POST /token 400 {"refresh_token":"secret"}')),
    }),
  );
  after(() => target.close());

  const res = await target.post({ headers: { [KEEPALIVE_HEADER]: SECRET } });
  assert.equal(res.status, 503);

  const body = await res.text();
  assert.deepEqual(JSON.parse(body), { status: 'failed', reason: 'unexpected', retryable: true });
  // An arbitrary error's message can carry a request body. This response reaches a
  // scheduler's logs.
  assert.ok(!body.includes('refresh_token'), body);
});

test('only POST is answered', async () => {
  const recorder: Recorder = { calls: 0 };
  const target = serve(keepaliveRouter(SECRET, succeeding(recorder)));
  after(() => target.close());

  for (const method of ['GET', 'PUT', 'DELETE']) {
    const res = await target.request(method, { headers: { [KEEPALIVE_HEADER]: SECRET } });
    assert.equal(res.status, 405, method);
    assert.equal(res.headers.get('allow'), 'POST');
  }

  // A GET that refreshed would let a link, a crawler or a browser preview spend a token
  // exchange, secret or no secret.
  assert.equal(recorder.calls, 0);
});

// ---------------------------------------------------------------------------
// Through createApp: mounting, and which gate the route is behind.
// ---------------------------------------------------------------------------

const BASE_CONFIG: Config = {
  graph: { clientId: 'client-id', authority: 'https://login.microsoftonline.com/common' },
  firestore: { cacheDocumentPath: DOC_PATH, projectId: 'proj' },
  oauth: {
    clientId: 'mcp-client',
    clientSecret: 'mcp-secret',
    tokenSigningKey: 'x'.repeat(32),
    publicUrl: 'https://onenote-mcp.example.run.app',
  },
  server: { port: 0 },
};

function serveApp(config: Config): { post: (init?: RequestInit) => Promise<Response>; close: () => void } {
  const server = createApp(config).listen(0);
  const ready = new Promise<void>((resolve) => server.once('listening', () => resolve()));

  return {
    post: async (init: RequestInit = {}) => {
      await ready;
      const { port } = server.address() as AddressInfo;
      return fetch(`http://127.0.0.1:${port}${KEEPALIVE_PATH}`, { method: 'POST', ...init });
    },
    close: () => server.close(),
  };
}

test('with no secret configured the route does not exist', async () => {
  const app = serveApp(BASE_CONFIG);
  after(() => app.close());

  // 404 rather than 401. An operator who has not configured this gets a scheduler job
  // that fails with "no such path", which says what is wrong; a 401 reads as a typo.
  assert.equal((await app.post()).status, 404);
  assert.equal((await app.post({ headers: { [KEEPALIVE_HEADER]: SECRET } })).status, 404);
});

test('with a secret configured the route is mounted outside the bearer gate', async () => {
  const app = serveApp({ ...BASE_CONFIG, server: { port: 0, keepaliveSecret: SECRET } });
  after(() => app.close());

  const res = await app.post();

  // 401 from the route's own check, not from requireBearerAuth: a scheduler has no
  // access token and no way to get one, so a WWW-Authenticate challenge would be a dead
  // end rather than a prompt.
  assert.equal(res.status, 401);
  assert.equal(res.headers.get('www-authenticate'), null);
  assert.deepEqual(await res.json(), { status: 'unauthorized' });
});
