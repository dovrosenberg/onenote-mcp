import test from 'node:test';
import assert from 'node:assert/strict';

import {
  TokenCacheError,
  isEmptyCache,
  overwriteWouldEmptyCache,
  readCache,
  type CacheSnapshot,
} from '../src/token-cache.ts';

const DOC_PATH = 'tokencache/msal';

function snapshot(data: Record<string, unknown> | undefined): CacheSnapshot {
  return { exists: data !== undefined, data: () => data };
}

function absentSnapshot(): CacheSnapshot {
  return { exists: false, data: () => undefined };
}

test('an absent document reads as an empty cache', () => {
  assert.equal(readCache(absentSnapshot(), DOC_PATH), null);
  // exists true but data() undefined is the same pre-bootstrap state, not an error.
  assert.equal(readCache({ exists: true, data: () => undefined }, DOC_PATH), null);
});

test('a document with a cache field returns it verbatim', () => {
  const blob = '{"Account": {"fake-account-id": {"realm": "  spaced  "}}}\n';

  assert.equal(readCache(snapshot({ cache: blob }), DOC_PATH), blob);
});

test('an empty cache field is an empty cache, not an absent one', () => {
  // The distinction is what stops beforeCacheAccess treating a deliberately emptied
  // document as pre-bootstrap.
  assert.equal(readCache(snapshot({ cache: '' }), DOC_PATH), '');
  assert.notEqual(readCache(snapshot({ cache: '' }), DOC_PATH), null);
});

test('a non-string cache field throws TokenCacheError', () => {
  const cases: Array<[string, Record<string, unknown>]> = [
    ['number', { cache: 42 }],
    ['object', { cache: null }],
    ['undefined', { updatedAt: 'whenever' }],
  ];

  for (const [foundType, data] of cases) {
    assert.throws(
      () => readCache(snapshot(data), 'other/doc'),
      (err: unknown) => {
        assert.ok(err instanceof TokenCacheError, `expected TokenCacheError, got ${String(err)}`);
        assert.equal(err.name, 'TokenCacheError');
        assert.equal(err.documentPath, 'other/doc');
        assert.match(err.message, /other\/doc/);
        assert.match(err.message, new RegExp(`found ${foundType}`));
        return true;
      },
    );
  }
});

test('extra fields on the document are ignored', () => {
  const blob = '{"Account": {}}';
  const data = { cache: blob, updatedAt: new Date(0), somethingElse: ['unexpected'] };

  assert.equal(readCache(snapshot(data), DOC_PATH), blob);
});

// ---------------------------------------------------------------------------
// The write guard.
//
// `afterCacheAccess` itself still has no automated test — it needs a Firestore backend,
// and the emulator is not installed here. What is tested is the decision it makes, which
// is a pure function over two strings and is where the behaviour worth protecting lives:
// the blob in that document is the only copy of the refresh token, and a write that
// replaces it with an empty cache costs a device-code sign-in to undo.
//
// `isEmptyCache` reads no MSAL key name on purpose. The fixtures below name `Account`
// and `RefreshToken` because that is what MSAL writes today, but every assertion holds
// with those keys renamed — which is the property that keeps this guard from silently
// inverting when MSAL changes its format.
// ---------------------------------------------------------------------------

const POPULATED = JSON.stringify({
  Account: { 'fake-account-id': { realm: 'fake-tenant' } },
  RefreshToken: { 'fake-rt-id': { secret: 'fake-secret' } },
  AppMetadata: {},
});

/** What MSAL serializes once the credentials have been removed from its cache. */
const EMPTIED = JSON.stringify({ Account: {}, RefreshToken: {}, AppMetadata: {} });

test('a cache holding credentials is not empty', () => {
  assert.equal(isEmptyCache(POPULATED), false);
  // The key names are not read: the same shape under different names answers the same.
  assert.equal(isEmptyCache(JSON.stringify({ Whatever: { one: { two: 3 } } })), false);
});

test('a cache whose every container is empty is empty', () => {
  assert.equal(isEmptyCache(EMPTIED), true);
  assert.equal(isEmptyCache('{}'), true);
  assert.equal(isEmptyCache(''), true);
  assert.equal(isEmptyCache('   \n'), true);
  assert.equal(isEmptyCache(JSON.stringify({ Renamed: {}, AlsoRenamed: [] })), true);
});

test('anything unrecognised is not empty, so the guard fails open', () => {
  // The only thing this decides is whether to refuse a write. Answering "empty" for a
  // format nobody here understands would block every write and strand the refresh token
  // in memory, which is the worse of the two mistakes.
  for (const value of ['not json at all', '[]', '"a string"', '42', 'null']) {
    assert.equal(isEmptyCache(value), false, value);
  }
});

test('replacing credentials with an empty cache is refused', () => {
  // The case the guard exists for: MSAL removes the refresh token from its in-memory
  // cache on some failures, and afterCacheAccess runs in a finally block, so the
  // emptied serialization can reach the document while the stored blob is still good.
  assert.equal(overwriteWouldEmptyCache(POPULATED, EMPTIED), true);

  // A stored blob this code cannot parse is refused too, because "not provably empty"
  // and "known to hold something" are the same answer to the only question being asked.
  // It costs nothing: a stored blob MSAL cannot deserialize fails earlier, in
  // beforeCacheAccess, as an unreadable cache.
  assert.equal(overwriteWouldEmptyCache('not a cache at all', EMPTIED), true);
});

test('every other combination is allowed through', () => {
  const cases: Array<[string, string | null, string]> = [
    ['the first write after bootstrap', null, POPULATED],
    ['filling an empty document', '', POPULATED],
    ['an ordinary refresh', POPULATED, POPULATED],
    ['emptying a document that was already empty', EMPTIED, EMPTIED],
    ['emptying an empty-string document', '', EMPTIED],
    ['replacing an unparseable document with a real cache', 'garbage', POPULATED],
  ];

  for (const [label, stored, candidate] of cases) {
    assert.equal(overwriteWouldEmptyCache(stored, candidate), false, label);
  }
});
