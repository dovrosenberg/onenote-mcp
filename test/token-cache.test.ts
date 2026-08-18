import test from 'node:test';
import assert from 'node:assert/strict';

import { TokenCacheError, readCache, type CacheSnapshot } from '../src/token-cache.ts';

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
