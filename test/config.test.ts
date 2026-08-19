import test from 'node:test';
import assert from 'node:assert/strict';

import { ConfigError, loadConfig, type ConfigGroup } from '../src/config.ts';

const ALL_GROUPS: ConfigGroup[] = ['graph', 'firestore', 'oauth', 'server'];

const SIGNING_KEY = 'x'.repeat(32);

function fullEnv(overrides: NodeJS.ProcessEnv = {}): NodeJS.ProcessEnv {
  return {
    ONENOTE_CLIENT_ID: 'client-id',
    ONENOTE_AUTHORITY: 'https://login.microsoftonline.com/common',
    MCP_OAUTH_CLIENT_ID: 'mcp-client',
    MCP_OAUTH_CLIENT_SECRET: 'mcp-secret',
    MCP_TOKEN_SIGNING_KEY: SIGNING_KEY,
    MCP_PUBLIC_URL: 'https://onenote-mcp.example.run.app',
    ...overrides,
  };
}

function expectConfigError(env: NodeJS.ProcessEnv, groups = ALL_GROUPS): ConfigError {
  try {
    loadConfig(groups, env);
  } catch (err) {
    assert.ok(err instanceof ConfigError, `expected ConfigError, got ${String(err)}`);
    return err;
  }
  throw new assert.AssertionError({ message: 'expected loadConfig to throw ConfigError' });
}

test('a complete environment loads every group', () => {
  const config = loadConfig(ALL_GROUPS, fullEnv({ PORT: '9090', GOOGLE_CLOUD_PROJECT: 'proj' }));

  assert.deepEqual(config.graph, {
    clientId: 'client-id',
    authority: 'https://login.microsoftonline.com/common',
  });
  assert.deepEqual(config.oauth, {
    clientId: 'mcp-client',
    clientSecret: 'mcp-secret',
    tokenSigningKey: SIGNING_KEY,
    publicUrl: 'https://onenote-mcp.example.run.app',
  });
  assert.deepEqual(config.firestore, {
    cacheDocumentPath: 'tokencache/msal',
    projectId: 'proj',
  });
  assert.deepEqual(config.server, { port: 9090 });
});

test('every missing variable is reported in one error, not just the first', () => {
  const env = fullEnv();
  delete env['MCP_OAUTH_CLIENT_ID'];
  delete env['MCP_TOKEN_SIGNING_KEY'];

  const err = expectConfigError(env);

  assert.deepEqual([...err.missing].sort(), ['MCP_OAUTH_CLIENT_ID', 'MCP_TOKEN_SIGNING_KEY']);
  assert.equal(err.invalid.length, 0);
});

test('the error is named ConfigError and names every missing variable in its message', () => {
  const err = expectConfigError({});

  assert.equal(err.name, 'ConfigError');
  for (const name of [
    'ONENOTE_CLIENT_ID',
    'ONENOTE_AUTHORITY',
    'MCP_OAUTH_CLIENT_ID',
    'MCP_OAUTH_CLIENT_SECRET',
    'MCP_TOKEN_SIGNING_KEY',
  ]) {
    assert.match(err.message, new RegExp(name));
  }
});

test('the bootstrap CLI groups load without any MCP_OAUTH_* variable set', () => {
  const config = loadConfig(['graph', 'firestore-explicit'], {
    ONENOTE_CLIENT_ID: 'client-id',
    ONENOTE_AUTHORITY: 'https://login.microsoftonline.com/common',
    FIRESTORE_CACHE_DOC: 'tokencache/msal',
    GOOGLE_CLOUD_PROJECT: 'proj',
  });

  assert.equal(config.graph?.clientId, 'client-id');
  assert.equal(config.firestore?.cacheDocumentPath, 'tokencache/msal');
  assert.equal(config.oauth, undefined);
  assert.equal(config.server, undefined);
});

test('firestore-explicit requires both names that firestore defaults or leaves unset', () => {
  // The bootstrap CLI writes with the operator's own credentials. Defaulting either name
  // would seed a real document in a project or at a path the deployed server never
  // reads, and the CLI would still report success.
  const err = expectConfigError({}, ['firestore-explicit']);

  assert.deepEqual([...err.missing].sort(), ['FIRESTORE_CACHE_DOC', 'GOOGLE_CLOUD_PROJECT']);
  assert.equal(err.invalid.length, 0);
});

test('firestore-explicit still rejects a collection path', () => {
  const err = expectConfigError({ FIRESTORE_CACHE_DOC: 'a/b/c', GOOGLE_CLOUD_PROJECT: 'proj' }, [
    'firestore-explicit',
  ]);

  assert.equal(err.invalid.length, 1);
  assert.match(err.invalid[0] ?? '', /^FIRESTORE_CACHE_DOC: /);
});

test('a missing variable is described by the group that was asked for', () => {
  // FIRESTORE_CACHE_DOC appears in two groups with different reasons for being needed.
  // Looking the description up by name alone would print whichever row of the table
  // comes first, which is the server's, in an error raised by the bootstrap CLI.
  const err = expectConfigError({ GOOGLE_CLOUD_PROJECT: 'proj' }, ['firestore-explicit']);

  assert.match(err.message, /FIRESTORE_CACHE_DOC — .*bootstrap CLI requires it/);
});

test('PORT defaults to 8080 and parses as a number', () => {
  assert.equal(loadConfig(['server'], {}).server?.port, 8080);
  assert.equal(loadConfig(['server'], { PORT: '3000' }).server?.port, 3000);
});

test('MCP_KEEPALIVE_SECRET is optional, and absent means no keepalive route', () => {
  // Absent rather than empty. src/server.ts mounts the route only when the property is
  // there, so the path 404s and a scheduler job fails with "no such path" rather than
  // with a 401 that reads as a mistyped secret.
  assert.equal('keepaliveSecret' in (loadConfig(['server'], {}).server ?? {}), false);
  assert.equal(loadConfig(['server'], { MCP_KEEPALIVE_SECRET: '  ' }).server?.keepaliveSecret, undefined);

  const secret = 's'.repeat(32);
  assert.equal(loadConfig(['server'], { MCP_KEEPALIVE_SECRET: secret }).server?.keepaliveSecret, secret);
});

test('a short MCP_KEEPALIVE_SECRET is invalid, held to the signing key length', () => {
  // It is the only credential on a route that spends a token exchange, so it gets the
  // same floor as MCP_TOKEN_SIGNING_KEY. Nothing here can check it was chosen randomly.
  const err = expectConfigError({ MCP_KEEPALIVE_SECRET: 'short' }, ['server']);

  assert.equal(err.missing.length, 0);
  assert.equal(err.invalid.length, 1);
  assert.match(err.invalid[0] ?? '', /^MCP_KEEPALIVE_SECRET: expected at least 32 characters, got 5$/);
});

test('a malformed PORT is invalid, not missing', () => {
  for (const value of ['abc', '70000', '0', '3000.5', '-1']) {
    const err = expectConfigError({ PORT: value }, ['server']);
    assert.equal(err.missing.length, 0, `PORT=${value} should not be reported as missing`);
    assert.equal(err.invalid.length, 1, `PORT=${value} should be reported as invalid`);
    assert.match(err.invalid[0] ?? '', /^PORT: /);
  }
});

test('FIRESTORE_CACHE_DOC defaults, and rejects a collection path', () => {
  assert.equal(loadConfig(['firestore'], {}).firestore?.cacheDocumentPath, 'tokencache/msal');
  assert.equal(
    loadConfig(['firestore'], { FIRESTORE_CACHE_DOC: 'a/b/c/d' }).firestore?.cacheDocumentPath,
    'a/b/c/d',
  );

  for (const value of ['tokencache', 'a/b/c', 'a//b', '/a/b']) {
    const err = expectConfigError({ FIRESTORE_CACHE_DOC: value }, ['firestore']);
    assert.equal(err.invalid.length, 1, `FIRESTORE_CACHE_DOC=${value} should be invalid`);
  }
});

test('ONENOTE_AUTHORITY must be an https URL', () => {
  for (const value of ['not-a-url', 'http://login.microsoftonline.com/common', 'common']) {
    const err = expectConfigError(fullEnv({ ONENOTE_AUTHORITY: value }), ['graph']);
    assert.equal(err.invalid.length, 1, `ONENOTE_AUTHORITY=${value} should be invalid`);
    assert.match(err.invalid[0] ?? '', /^ONENOTE_AUTHORITY: /);
  }
});

test('MCP_TOKEN_SIGNING_KEY shorter than 32 characters is rejected', () => {
  const err = expectConfigError(fullEnv({ MCP_TOKEN_SIGNING_KEY: 'short' }), ['oauth']);

  assert.equal(err.invalid.length, 1);
  assert.match(err.invalid[0] ?? '', /^MCP_TOKEN_SIGNING_KEY: expected at least 32 characters/);
  assert.equal(loadConfig(['oauth'], fullEnv()).oauth?.tokenSigningKey, SIGNING_KEY);
});

test('MCP_PUBLIC_URL must be an https origin with nothing after it', () => {
  // Every OAuth URL this service publishes is this value with a path concatenated onto
  // it, so a trailing slash, a query string or a fragment would appear in the middle of
  // the result. RFC 8414 forbids the last two on an issuer identifier outright.
  const bad = [
    'not-a-url',
    'http://onenote-mcp.example.run.app',
    'https://onenote-mcp.example.run.app/',
    'https://onenote-mcp.example.run.app?a=1',
    'https://onenote-mcp.example.run.app#f',
  ];

  for (const value of bad) {
    const err = expectConfigError(fullEnv({ MCP_PUBLIC_URL: value }), ['oauth']);
    assert.equal(err.invalid.length, 1, `MCP_PUBLIC_URL=${value} should be invalid`);
    assert.match(err.invalid[0] ?? '', /^MCP_PUBLIC_URL: /);
  }

  assert.equal(
    loadConfig(['oauth'], fullEnv()).oauth?.publicUrl,
    'https://onenote-mcp.example.run.app',
  );
});

test('a whitespace-only value counts as missing, and values are trimmed', () => {
  const err = expectConfigError(fullEnv({ ONENOTE_CLIENT_ID: '   ' }), ['graph']);
  assert.deepEqual([...err.missing], ['ONENOTE_CLIENT_ID']);

  const config = loadConfig(['graph'], fullEnv({ ONENOTE_CLIENT_ID: '  client-id  ' }));
  assert.equal(config.graph?.clientId, 'client-id');
});

test('a whitespace-only optional value falls back to its default', () => {
  assert.equal(loadConfig(['server'], { PORT: '  ' }).server?.port, 8080);
});

test('GOOGLE_CLOUD_PROJECT is optional and left undefined when unset', () => {
  assert.equal(loadConfig(['firestore'], {}).firestore?.projectId, undefined);
});

test('missing and invalid variables are reported together', () => {
  const env = fullEnv({ PORT: 'abc' });
  delete env['ONENOTE_CLIENT_ID'];

  const err = expectConfigError(env);

  assert.deepEqual([...err.missing], ['ONENOTE_CLIENT_ID']);
  assert.equal(err.invalid.length, 1);
  assert.match(err.message, /ONENOTE_CLIENT_ID/);
  assert.match(err.message, /PORT/);
});

test('loadConfig does not read the real process.env when an env is injected', () => {
  process.env['ONENOTE_CLIENT_ID'] = 'leaked-from-real-env';
  try {
    const err = expectConfigError({}, ['graph']);
    assert.ok(err.missing.includes('ONENOTE_CLIENT_ID'));
  } finally {
    delete process.env['ONENOTE_CLIENT_ID'];
  }
});

// ---------------------------------------------------------------------------
// The mirror group (issue #30).
//
// Every variable is optional, so an unset environment leaves the feature entirely off
// and the deployed service behaves exactly as it did before the group existed. That is
// the property the first test asserts, and it is what lets the mirror ship inert.
// ---------------------------------------------------------------------------

const MIRROR_GROUPS: ConfigGroup[] = [...ALL_GROUPS, 'mirror'];

test('the mirror group is entirely optional and defaults to off', () => {
  const config = loadConfig(MIRROR_GROUPS, fullEnv());

  assert.deepEqual(config.mirror, {
    rootDocumentPath: 'onenoteMirror/default',
    readEnabled: false,
    syncRequestBudget: 120,
  });
});

test('the mirror group reads every variable when they are all set', () => {
  const config = loadConfig(
    MIRROR_GROUPS,
    fullEnv({
      MIRROR_ROOT_DOC: 'mirrors/prod',
      MIRROR_SYNC_SECRET: 'y'.repeat(32),
      MIRROR_BUCKET: 'onenote-mcp-mirror-505918',
      MIRROR_READ_ENABLED: 'true',
      MIRROR_SYNC_REQUEST_BUDGET: '60',
    }),
  );

  assert.deepEqual(config.mirror, {
    rootDocumentPath: 'mirrors/prod',
    syncSecret: 'y'.repeat(32),
    bucket: 'onenote-mcp-mirror-505918',
    readEnabled: true,
    syncRequestBudget: 60,
  });
});

test('MIRROR_SYNC_SECRET without MIRROR_BUCKET is a ConfigError naming the bucket', () => {
  // A sync has nowhere to put a rendered ink PNG without a bucket. SPECS cannot express
  // "required when another variable is present", so loadConfig checks the pair itself —
  // the first cross-field rule in the file.
  const err = expectConfigError(
    fullEnv({ MIRROR_SYNC_SECRET: 'y'.repeat(32) }),
    MIRROR_GROUPS,
  );

  assert.deepEqual(err.missing, ['MIRROR_BUCKET']);
  assert.match(err.message, /MIRROR_BUCKET/);
});

test('MIRROR_READ_ENABLED without MIRROR_BUCKET is a ConfigError too', () => {
  // Reads serve the stored ink PNG out of the bucket, so enabling them without one is
  // the same mistake from the other direction.
  const err = expectConfigError(
    fullEnv({ MIRROR_READ_ENABLED: 'true' }),
    MIRROR_GROUPS,
  );

  assert.deepEqual(err.missing, ['MIRROR_BUCKET']);
});

test('a bucket with neither the secret nor reads set is allowed', () => {
  // Provisioning the bucket before turning anything on is the ordinary deploy sequence.
  const config = loadConfig(
    MIRROR_GROUPS,
    fullEnv({ MIRROR_BUCKET: 'onenote-mcp-mirror-505918' }),
  );

  assert.equal(config.mirror?.bucket, 'onenote-mcp-mirror-505918');
  assert.equal(config.mirror?.readEnabled, false);
  assert.equal(config.mirror?.syncSecret, undefined);
});

test('MIRROR_READ_ENABLED accepts only true and false', () => {
  for (const value of ['true', 'TRUE', 'True']) {
    const config = loadConfig(MIRROR_GROUPS, fullEnv({
      MIRROR_READ_ENABLED: value,
      MIRROR_BUCKET: 'b-ucket',
    }));
    assert.equal(config.mirror?.readEnabled, true, value);
  }

  for (const value of ['false', 'FALSE']) {
    const config = loadConfig(MIRROR_GROUPS, fullEnv({ MIRROR_READ_ENABLED: value }));
    assert.equal(config.mirror?.readEnabled, false, value);
  }

  // Nothing else is guessed at. "1", "yes" and "on" all read as true to a human and
  // would silently enable the mirror if this were lenient.
  for (const value of ['1', 'yes', 'on', '0', 'no']) {
    const err = expectConfigError(fullEnv({ MIRROR_READ_ENABLED: value }), MIRROR_GROUPS);
    assert.deepEqual(err.missing, []);
    assert.match(err.message, /MIRROR_READ_ENABLED/);
  }
});

test('MIRROR_SYNC_REQUEST_BUDGET is bounded on both sides', () => {
  const budget = (value: string): number | undefined =>
    loadConfig(MIRROR_GROUPS, fullEnv({ MIRROR_SYNC_REQUEST_BUDGET: value })).mirror
      ?.syncRequestBudget;

  assert.equal(budget('10'), 10);
  assert.equal(budget('350'), 350);

  // Below 10 a run cannot get past the structure read and the first section, so it would
  // never make progress. Above 350 one run could spend the whole hourly budget of 400
  // and leave nothing for the interactive tools.
  for (const value of ['9', '351', '0', '-5', 'lots', '12.5']) {
    const err = expectConfigError(
      fullEnv({ MIRROR_SYNC_REQUEST_BUDGET: value }),
      MIRROR_GROUPS,
    );
    assert.match(err.message, /MIRROR_SYNC_REQUEST_BUDGET/);
  }
});

test('MIRROR_ROOT_DOC is held to the same document-path rule as the token cache', () => {
  const err = expectConfigError(fullEnv({ MIRROR_ROOT_DOC: 'onenoteMirror' }), MIRROR_GROUPS);
  assert.match(err.message, /MIRROR_ROOT_DOC/);

  assert.equal(
    loadConfig(MIRROR_GROUPS, fullEnv({ MIRROR_ROOT_DOC: 'a/b/c/d' })).mirror?.rootDocumentPath,
    'a/b/c/d',
  );
});

test('MIRROR_BUCKET is checked against the GCS naming rules, not just non-empty', () => {
  const ok = (value: string): string | undefined =>
    loadConfig(MIRROR_GROUPS, fullEnv({ MIRROR_BUCKET: value })).mirror?.bucket;

  assert.equal(ok('onenote-mcp-mirror-505918'), 'onenote-mcp-mirror-505918');
  assert.equal(ok('a.b_c-1'), 'a.b_c-1');

  // A typo fails here, at startup, alongside every other configuration problem, rather
  // than at the first PUT hours into a backfill.
  for (const value of [
    'gs://onenote-mcp-mirror',
    'OneNoteMirror',
    'has spaces',
    'ab',
    'x'.repeat(64),
    '-leading-dash',
    'trailing-dash-',
  ]) {
    const err = expectConfigError(fullEnv({ MIRROR_BUCKET: value }), MIRROR_GROUPS);
    assert.match(err.message, /MIRROR_BUCKET/);
  }
});

test('a caller that does not ask for the mirror group gets no mirror config', () => {
  const config = loadConfig(ALL_GROUPS, fullEnv({ MIRROR_READ_ENABLED: 'true' }));
  assert.equal(config.mirror, undefined);
});
