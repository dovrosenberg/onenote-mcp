// POST /sync, driven over a real HTTP server.
//
// Its own router with a fake target for everything the route decides, because the real
// target reaches Graph, Firestore and Cloud Storage. `createApp` is used for the two
// facts a fake router cannot show: that the route is absent when no secret is
// configured, and that it sits outside the bearer gate — a scheduler holds a shared
// secret and cannot run an OAuth flow, so a request with no Authorization header must
// get this route's own 401 rather than the challenge `/mcp` gives.
//
// The three-paths-not-one-mode decision is asserted here rather than described, because
// it is the kind of thing a later change would quietly collapse into a query parameter.
// The reason it must not be is in src/logging.ts: the request log records the path and
// deliberately records no query string and no body, so a mode carried in either would
// appear in no log line.

import test, { after } from 'node:test';
import assert from 'node:assert/strict';
import type { AddressInfo } from 'node:net';

import express from 'express';

import type { Config } from '../src/config.ts';
import { GraphAuthError } from '../src/graph-auth.ts';
import { setEventSink } from '../src/logging.ts';
import { MirrorLeaseHeldError } from '../src/mirror-store.ts';
import type { SyncReport } from '../src/mirror-sync.ts';
import { createApp } from '../src/server.ts';
import { SYNC_HEADER, SYNC_PATH, syncRouter, type SyncTarget } from '../src/sync-route.ts';

const SECRET = 's'.repeat(32);

// Every failure path below logs an operational event on purpose. Silenced so the run
// shows the assertions rather than the lines they expect.
setEventSink(() => {});

function report(overrides: Partial<SyncReport> = {}): SyncReport {
  return {
    mode: 'incremental',
    outcome: 'complete',
    done: true,
    graphRequests: 3,
    sectionsVisited: 1,
    pagesUpdated: 2,
    pagesDeleted: 0,
    pagesFailed: 0,
    unknownNotebookIds: 0,
    unknownActiveNotebookIds: 0,
    sectionsSkippedInactive: 0,
    treeRead: true,
    durationMs: 1234,
    ...overrides,
  };
}

interface Recorder {
  readonly calls: string[];
}

function recording(recorder: Recorder, fail?: () => Error): SyncTarget {
  const make = (mode: string) => (): Promise<SyncReport> => {
    recorder.calls.push(mode);
    if (fail !== undefined) return Promise.reject(fail());
    return Promise.resolve(report({ mode: mode === 'runIncremental' ? 'incremental' : 'sweep' }));
  };

  return {
    runIncremental: make('runIncremental'),
    runSweep: make('runSweep'),
    runFullSweep: make('runFullSweep'),
    runSweepAll: make('runSweepAll'),
  };
}

function serve(router: express.Router): {
  send: (method: string, path?: string, init?: RequestInit) => Promise<Response>;
  close: () => void;
} {
  const app = express();
  app.use(SYNC_PATH, router);
  const server = app.listen(0);
  const ready = new Promise<void>((resolve) => server.once('listening', () => resolve()));

  return {
    send: async (method, path = '', init = {}) => {
      await ready;
      const { port } = server.address() as AddressInfo;
      return fetch(`http://127.0.0.1:${port}${SYNC_PATH}${path}`, { method, ...init });
    },
    close: () => server.close(),
  };
}

const auth = { headers: { [SYNC_HEADER]: SECRET } };

test('each path reaches its own mode and no other', async () => {
  // The whole reason there are three paths rather than one with a mode argument: the
  // request log records the path, so this is the only place the distinction survives.
  const recorder: Recorder = { calls: [] };
  const target = serve(syncRouter(SECRET, recording(recorder)));
  after(() => target.close());

  for (const [path, expected] of [
    ['', 'runIncremental'],
    ['/sweep', 'runSweep'],
    ['/sweep/full', 'runFullSweep'],
    ['/sweep/all', 'runSweepAll'],
  ] as const) {
    recorder.calls.length = 0;
    const res = await target.send('POST', path, auth);
    assert.equal(res.status, 200, path);
    assert.deepEqual(recorder.calls, [expected], path);
  }
});

test('a successful run answers 200 with the whole report', async () => {
  const target = serve(syncRouter(SECRET, recording({ calls: [] })));
  after(() => target.close());

  const res = await target.send('POST', '', auth);
  const body = (await res.json()) as Record<string, unknown>;

  assert.equal(res.status, 200);
  assert.equal(body['status'], 'ok');
  // Every count the report carries reaches the operator. A 200 that said only "ok" would
  // be indistinguishable from a run that did nothing.
  assert.equal(body['graphRequests'], 3);
  assert.equal(body['pagesUpdated'], 2);
  assert.equal(body['done'], true);
  assert.equal(body['outcome'], 'complete');
  assert.equal(res.headers.get('cache-control'), 'no-store');
});

test('a budget-exhausted run is 200, not 503', async () => {
  // It is a normal outcome with committed work behind it. A 503 would make the scheduler
  // retry immediately and spend the next hour's Graph budget inside this one, which is
  // the failure the budget exists to prevent.
  const exhausted: SyncTarget = {
    runIncremental: () =>
      Promise.resolve(report({ outcome: 'budget-exhausted', done: false, graphRequests: 120 })),
    runSweep: () => Promise.resolve(report()),
    runFullSweep: () => Promise.resolve(report()),
    runSweepAll: () => Promise.resolve(report()),
  };
  const target = serve(syncRouter(SECRET, exhausted));
  after(() => target.close());

  const res = await target.send('POST', '', auth);
  const body = (await res.json()) as Record<string, unknown>;

  assert.equal(res.status, 200);
  assert.equal(body['outcome'], 'budget-exhausted');
  assert.equal(body['done'], false, 'the scheduler learns there is more to do from the body');
});

test('a wrong or missing secret is 401 on every path and costs no run', async () => {
  const recorder: Recorder = { calls: [] };
  const target = serve(syncRouter(SECRET, recording(recorder)));
  after(() => target.close());

  const cases: Array<[string, RequestInit]> = [
    ['no header', {}],
    ['wrong value, same length', { headers: { [SYNC_HEADER]: 'w'.repeat(32) } }],
    ['a prefix of the secret', { headers: { [SYNC_HEADER]: SECRET.slice(0, 16) } }],
    ['the secret with a suffix', { headers: { [SYNC_HEADER]: `${SECRET}x` } }],
    ['empty', { headers: { [SYNC_HEADER]: '' } }],
  ];

  for (const path of ['', '/sweep', '/sweep/full', '/sweep/all']) {
    for (const [label, init] of cases) {
      const res = await target.send('POST', path, init);
      assert.equal(res.status, 401, `${path} ${label}`);
      assert.deepEqual(await res.json(), { status: 'unauthorized' }, `${path} ${label}`);
    }
  }

  // The point of checking the secret first: an unauthenticated flood costs a string
  // comparison, not a walk of the account's structure.
  assert.deepEqual(recorder.calls, []);
});

test('a held lease is 409 and retryable, not a failure', async () => {
  // The nightly sweep landing on a still-running incremental is an ordinary overlap.
  const target = serve(
    syncRouter(SECRET, recording({ calls: [] }, () => new MirrorLeaseHeldError('incremental'))),
  );
  after(() => target.close());

  const res = await target.send('POST', '', auth);
  const body = (await res.json()) as Record<string, unknown>;

  assert.equal(res.status, 409);
  assert.equal(body['status'], 'conflict');
  assert.equal(body['heldBy'], 'incremental');
  assert.equal(body['retryable'], true);
});

test('a dead refresh token reaches the operator with its own reason', async () => {
  // The one failure a human has to act on. Flattening it into "unexpected" would send
  // someone to the service logs to rediscover that the bootstrap CLI needs re-running.
  const target = serve(
    syncRouter(SECRET, recording({ calls: [] }, () =>
      new GraphAuthError('silent-failed', 'tokencache/msal'),
    )),
  );
  after(() => target.close());

  const res = await target.send('POST', '', auth);
  const body = (await res.json()) as Record<string, unknown>;

  assert.equal(res.status, 503);
  assert.equal(body['status'], 'failed');
  assert.equal(body['reason'], 'silent-failed');
  assert.match(String(body['detail']), /bootstrap/);
});

test('any other failure is 503 and quotes nothing', async () => {
  const target = serve(
    syncRouter(SECRET, recording({ calls: [] }, () =>
      new Error('connect ECONNREFUSED 10.0.0.1:443 while writing /users/dov/secret'),
    )),
  );
  after(() => target.close());

  const res = await target.send('POST', '', auth);
  const body = (await res.json()) as Record<string, unknown>;

  assert.equal(res.status, 503);
  assert.deepEqual(body, { status: 'failed', reason: 'unexpected', retryable: true });
  // An arbitrary error message can carry a path, a request body or a host. None of it
  // reaches a response a scheduler may log.
  assert.equal(JSON.stringify(body).includes('ECONNREFUSED'), false);
});

test('only POST is answered; everything else is 405 with Allow', async () => {
  // A GET would let a link preview, or anything that crawls a URL, spend a slice of the
  // hourly Graph budget.
  const recorder: Recorder = { calls: [] };
  const target = serve(syncRouter(SECRET, recording(recorder)));
  after(() => target.close());

  for (const path of ['', '/sweep', '/sweep/full', '/sweep/all']) {
    for (const method of ['GET', 'PUT', 'DELETE', 'PATCH']) {
      const res = await target.send(method, path, auth);
      assert.equal(res.status, 405, `${method} ${path}`);
      assert.equal(res.headers.get('allow'), 'POST', `${method} ${path}`);
    }
  }

  assert.deepEqual(recorder.calls, []);
});

// ---------------------------------------------------------------------------
// Through createApp: the two facts a bare router cannot show.
// ---------------------------------------------------------------------------

const BASE_CONFIG: Config = {
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

function serveApp(config: Config): {
  send: (method: string, path: string, init?: RequestInit) => Promise<Response>;
  close: () => void;
} {
  const server = createApp(config).listen(0);
  const ready = new Promise<void>((resolve) => server.once('listening', () => resolve()));

  return {
    send: async (method, path, init = {}) => {
      await ready;
      const { port } = server.address() as AddressInfo;
      return fetch(`http://127.0.0.1:${port}${path}`, { method, ...init });
    },
    close: () => server.close(),
  };
}

test('with no sync secret configured, every sync path is a 404', async () => {
  // A 404 tells an operator the service is not set up. A 401 would read as a mistyped
  // secret and send them to check the scheduler job instead of the deploy.
  const app = serveApp(BASE_CONFIG);
  after(() => app.close());

  for (const path of [
    SYNC_PATH,
    `${SYNC_PATH}/sweep`,
    `${SYNC_PATH}/sweep/full`,
    `${SYNC_PATH}/sweep/all`,
  ]) {
    const res = await app.send('POST', path);
    assert.equal(res.status, 404, path);
  }
});

test('mounted, the route answers its own 401 rather than the bearer challenge', async () => {
  // This is what "outside the bearer gate" means in practice. A scheduler has no browser
  // and nowhere to keep a refresh token, so it can never satisfy `/mcp`'s challenge.
  const app = serveApp({
    ...BASE_CONFIG,
    mirror: {
      rootDocumentPath: 'onenoteMirror/default',
      syncSecret: SECRET,
      bucket: 'a-bucket',
      readEnabled: false,
      syncRequestBudget: 120,
    },
  });
  after(() => app.close());

  const res = await app.send('POST', SYNC_PATH);

  assert.equal(res.status, 401);
  assert.deepEqual(await res.json(), { status: 'unauthorized' });
  assert.equal(
    res.headers.get('www-authenticate'),
    null,
    'a WWW-Authenticate header here would be the bearer gate answering, not this route',
  );
});
