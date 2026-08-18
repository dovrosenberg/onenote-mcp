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
  const config = loadConfig(['graph', 'firestore'], {
    ONENOTE_CLIENT_ID: 'client-id',
    ONENOTE_AUTHORITY: 'https://login.microsoftonline.com/common',
  });

  assert.equal(config.graph?.clientId, 'client-id');
  assert.equal(config.firestore?.cacheDocumentPath, 'tokencache/msal');
  assert.equal(config.oauth, undefined);
  assert.equal(config.server, undefined);
});

test('PORT defaults to 8080 and parses as a number', () => {
  assert.equal(loadConfig(['server'], {}).server?.port, 8080);
  assert.equal(loadConfig(['server'], { PORT: '3000' }).server?.port, 3000);
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
