import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile, readdir } from 'node:fs/promises';
import path from 'node:path';

import type { AccountInfo, AuthenticationResult, SilentFlowRequest } from '@azure/msal-node';

import {
  GRAPH_SCOPES,
  GraphAuthError,
  acquireGraphToken,
  type SilentTokenSource,
} from '../src/graph-auth.ts';
import { TokenCacheError } from '../src/token-cache.ts';

const DOC_PATH = 'tokencache/msal';

/** Fabricated throughout: a fake tenant of all zeroes and a reserved-TLD username. */
function fakeAccount(): AccountInfo {
  return {
    homeAccountId: 'fake-object-id.00000000-0000-0000-0000-000000000000',
    environment: 'login.microsoftonline.com',
    tenantId: '00000000-0000-0000-0000-000000000000',
    username: 'nobody@example.invalid',
    localAccountId: 'fake-object-id',
  };
}

interface FakeOptions {
  accounts?: () => Promise<AccountInfo[]>;
  silent?: (request: SilentFlowRequest) => Promise<AuthenticationResult | null>;
}

function fakeSource(options: FakeOptions): SilentTokenSource {
  return {
    getTokenCache: () => ({
      getAllAccounts: options.accounts ?? (() => Promise.resolve([fakeAccount()])),
    }),
    acquireTokenSilent:
      options.silent ??
      (() => {
        throw new Error('acquireTokenSilent was not expected in this test');
      }),
  };
}

function fakeResult(accessToken: string): AuthenticationResult {
  return {
    authority: 'https://login.microsoftonline.com/common/',
    uniqueId: 'fake-object-id',
    tenantId: '00000000-0000-0000-0000-000000000000',
    scopes: [...GRAPH_SCOPES],
    account: fakeAccount(),
    idToken: '',
    idTokenClaims: {},
    accessToken,
    fromCache: false,
    expiresOn: new Date(0),
    tokenType: 'Bearer',
    correlationId: 'fake-correlation-id',
  };
}

async function caught(promise: Promise<unknown>): Promise<unknown> {
  try {
    await promise;
  } catch (err) {
    return err;
  }
  assert.fail('expected the call to reject');
}

test('GRAPH_SCOPES is exactly Notes.Read and Notes.ReadWrite, fully qualified', () => {
  assert.deepEqual(GRAPH_SCOPES, [
    'https://graph.microsoft.com/Notes.Read',
    'https://graph.microsoft.com/Notes.ReadWrite',
  ]);
});

test('a token is returned and the request carries the scopes and the cached account', async () => {
  const account = fakeAccount();
  const requests: SilentFlowRequest[] = [];
  const client = fakeSource({
    accounts: () => Promise.resolve([account]),
    silent: (request) => {
      requests.push(request);
      return Promise.resolve(fakeResult('fake-access-token'));
    },
  });

  assert.equal(await acquireGraphToken(client, DOC_PATH), 'fake-access-token');
  assert.equal(requests.length, 1);
  assert.equal(requests[0]?.account, account);
  assert.deepEqual(requests[0]?.scopes, [...GRAPH_SCOPES]);
});

test('an empty cache throws GraphAuthError with reason no-account', async () => {
  const client = fakeSource({ accounts: () => Promise.resolve([]) });

  const err = await caught(acquireGraphToken(client, DOC_PATH));

  assert.ok(err instanceof GraphAuthError, `expected GraphAuthError, got ${String(err)}`);
  assert.equal(err.name, 'GraphAuthError');
  assert.equal(err.reason, 'no-account');
  assert.equal(err.documentPath, DOC_PATH);
  assert.match(err.message, /tokencache\/msal/);
  assert.match(err.message, /npm run bootstrap/);
});

test('a cache that will not decode throws GraphAuthError with reason cache-unreadable', async () => {
  // Both failures happen inside getAllAccounts, because that is the call that triggers
  // beforeCacheAccess: TokenCacheError from the decoder, SyntaxError from MSAL's
  // deserialize when the stored blob is a string but not JSON it recognises.
  const causes: Error[] = [
    new TokenCacheError('cache field is not a string', DOC_PATH),
    new SyntaxError('Unexpected token o in JSON at position 1'),
  ];

  for (const cause of causes) {
    const client = fakeSource({ accounts: () => Promise.reject(cause) });

    const err = await caught(acquireGraphToken(client, DOC_PATH));

    assert.ok(err instanceof GraphAuthError, `expected GraphAuthError, got ${String(err)}`);
    assert.equal(err.reason, 'cache-unreadable');
    assert.equal(err.documentPath, DOC_PATH);
    assert.equal(err.cause, cause);
    assert.match(err.message, /npm run bootstrap/);
    assert.match(err.message, new RegExp(cause.name));
  }
});

test('a dead refresh token throws GraphAuthError with reason silent-failed and does not leak the account', async () => {
  // Shaped like MSAL's InteractionRequiredAuthError for a revoked or expired refresh
  // token. Constructing the real class would couple the test to an internal signature.
  const cause = Object.assign(new Error('AADSTS700082: The refresh token has expired.'), {
    name: 'InteractionRequiredAuthError',
    errorCode: 'invalid_grant',
  });
  const account = fakeAccount();
  const client = fakeSource({
    accounts: () => Promise.resolve([account]),
    silent: () => Promise.reject(cause),
  });

  const err = await caught(acquireGraphToken(client, DOC_PATH));

  assert.ok(err instanceof GraphAuthError, `expected GraphAuthError, got ${String(err)}`);
  assert.equal(err.reason, 'silent-failed');
  assert.equal(err.cause, cause);
  assert.match(err.message, /expired or revoked/);
  assert.match(err.message, /npm run bootstrap/);
  assert.match(err.message, /InteractionRequiredAuthError/);
  // The UPN and the tenant id must not reach a message that can end up in a log.
  assert.doesNotMatch(err.message, new RegExp(account.username));
  assert.doesNotMatch(err.message, new RegExp(account.homeAccountId));
  assert.doesNotMatch(err.message, new RegExp(account.tenantId));
});

test('a null or blank result throws rather than returning an unusable token', async () => {
  // Either would otherwise be sent to Graph as a Bearer header and come back a 401 —
  // the failure mode this module exists to replace with a message that says what to do.
  const results: Array<AuthenticationResult | null> = [null, fakeResult('')];

  for (const result of results) {
    const client = fakeSource({ silent: () => Promise.resolve(result) });

    const err = await caught(acquireGraphToken(client, DOC_PATH));

    assert.ok(err instanceof GraphAuthError, `expected GraphAuthError, got ${String(err)}`);
    assert.equal(err.reason, 'silent-failed');
    assert.match(err.message, /npm run bootstrap/);
  }
});

test('no module under src/ calls acquireTokenByDeviceCode except bootstrap.ts', async () => {
  // A source-text check, not a behavioural one. The requirement is about a call site
  // that must never be added to the server path, and no runtime assertion can observe a
  // call that is absent. Device-code sign-in blocks on a human who is not there; on
  // Cloud Run the request would simply time out.
  const srcDir = path.join(import.meta.dirname, '..', 'src');
  const files = (await readdir(srcDir)).filter((name) => name.endsWith('.ts'));
  assert.ok(files.includes('graph-auth.ts'), 'expected to have scanned graph-auth.ts');

  for (const name of files) {
    const source = await readFile(path.join(srcDir, name), 'utf8');
    const mentions = source.includes('acquireTokenByDeviceCode');
    if (name === 'bootstrap.ts') continue;
    assert.equal(mentions, false, `src/${name} must not call acquireTokenByDeviceCode`);
  }
});
