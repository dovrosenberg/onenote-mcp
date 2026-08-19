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
