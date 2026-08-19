// Decoding a Graph response body, with no network anywhere in it.
//
// These functions were inside src/graph-structure.ts until they were extracted, and they
// were reachable only through a fake `fetch` keyed by exact URL. That made every decode
// assertion pay for a routing table, and it made the tolerances — a null displayName, an
// absent expanded relationship — awkward to state directly. Here they are stated
// directly.
//
// The last test in this file is the one that matters most and is the easiest to lose:
// no message a decoder throws may contain the value that failed. These bodies carry
// notebook, section and page names, which are the user's own writing, and this
// repository's output can reach a public log.

import test from 'node:test';
import assert from 'node:assert/strict';

import { GraphResponseError } from '../src/graph-structure.ts';
import {
  asRecord,
  describeError,
  isRecord,
  mapWithLimit,
  optionalString,
  quoteOData,
  requireString,
  safeText,
  toExpandedNotebook,
  toNode,
  toNodeArray,
  toPageSummary,
  toSectionWithParents,
  truncate,
} from '../src/graph-decode.ts';

const URL_ = 'https://graph.microsoft.com/v1.0/me/onenote/notebooks';

test('toNode reads id and displayName, and a null name becomes an empty string', () => {
  assert.deepEqual(toNode({ id: 'nb-1', displayName: '2026' }, URL_), {
    id: 'nb-1',
    displayName: '2026',
  });

  // Graph returns a null displayName on some nodes. A caller that only formats the name
  // should not have to handle undefined, so it is normalised rather than passed through.
  assert.deepEqual(toNode({ id: 'nb-2', displayName: null }, URL_), {
    id: 'nb-2',
    displayName: '',
  });
  assert.deepEqual(toNode({ id: 'nb-3' }, URL_), { id: 'nb-3', displayName: '' });
});

test('a node with no usable id is a GraphResponseError naming the request', () => {
  for (const item of [{ displayName: 'x' }, { id: '', displayName: 'x' }, { id: 42 }]) {
    assert.throws(
      () => toNode(item, URL_),
      (err: unknown) => err instanceof GraphResponseError && err.url === URL_,
    );
  }
});

test('a non-object inside "value" is a GraphResponseError', () => {
  for (const item of ['a string', 42, null, ['an', 'array']]) {
    assert.throws(() => asRecord(item, URL_), GraphResponseError);
  }
  assert.deepEqual(asRecord({ id: 'x' }, URL_), { id: 'x' });
});

test('an expanded relationship is empty when absent and a fault when not an array', () => {
  // Graph omits a relationship that holds nothing, and a notebook with no section groups
  // is ordinary rather than a malformed response.
  assert.deepEqual(toNodeArray(undefined, URL_), []);
  assert.deepEqual(toNodeArray(null, URL_), []);
  assert.deepEqual(toNodeArray([{ id: 'a' }], URL_), [{ id: 'a' }]);

  // Present but the wrong shape is the service returning something nothing here can use.
  assert.throws(() => toNodeArray({ id: 'a' }, URL_), GraphResponseError);
  assert.throws(() => toNodeArray('sections', URL_), GraphResponseError);
});

test('toExpandedNotebook reaches the sections inside a section group', () => {
  const decoded = toExpandedNotebook(
    {
      id: 'nb-1',
      displayName: '2026',
      sections: [{ id: 's-1', displayName: 'Index' }],
      sectionGroups: [
        {
          id: 'g-1',
          displayName: '062 - February',
          sections: [
            { id: 's-2', displayName: 'Daily log' },
            { id: 's-3', displayName: null },
          ],
        },
      ],
    },
    URL_,
  );

  assert.deepEqual(decoded, {
    id: 'nb-1',
    displayName: '2026',
    sections: [{ id: 's-1', displayName: 'Index' }],
    sectionGroups: [
      {
        id: 'g-1',
        displayName: '062 - February',
        sections: [
          { id: 's-2', displayName: 'Daily log' },
          { id: 's-3', displayName: '' },
        ],
      },
    ],
  });
});

test('a notebook with neither relationship expanded decodes to two empty lists', () => {
  assert.deepEqual(toExpandedNotebook({ id: 'nb-1', displayName: 'Scratch' }, URL_), {
    id: 'nb-1',
    displayName: 'Scratch',
    sections: [],
    sectionGroups: [],
  });
});

test('toSectionWithParents reads a null parent as null, not as missing', () => {
  // A section directly under a notebook has no parent section group, and Graph sends
  // null rather than omitting the key. Both readings have to end in the same place.
  assert.deepEqual(
    toSectionWithParents(
      {
        id: 's-1',
        displayName: 'Index',
        parentNotebook: { id: 'nb-1', displayName: '2026' },
        parentSectionGroup: null,
      },
      URL_,
    ),
    {
      id: 's-1',
      displayName: 'Index',
      parentNotebook: { id: 'nb-1', displayName: '2026' },
      parentSectionGroup: null,
    },
  );

  const orphan = toSectionWithParents({ id: 's-2', displayName: 'Loose' }, URL_);
  assert.equal(orphan.parentNotebook, null);
  assert.equal(orphan.parentSectionGroup, null);
});

test('toPageSummary defaults the two optional strings rather than dropping the page', () => {
  assert.deepEqual(toPageSummary({ id: 'p-1', title: 'Monday' , lastModifiedDateTime: '2026-08-19T12:00:00Z' }, URL_), {
    id: 'p-1',
    title: 'Monday',
    lastModifiedDateTime: '2026-08-19T12:00:00Z',
  });

  // An untitled page is a real thing in OneNote; it must not become a decode failure.
  assert.deepEqual(toPageSummary({ id: 'p-2' }, URL_), {
    id: 'p-2',
    title: '',
    lastModifiedDateTime: '',
  });

  assert.throws(() => toPageSummary({ title: 'no id' }, URL_), GraphResponseError);
});

test('isRecord and the two field readers', () => {
  assert.equal(isRecord({}), true);
  assert.equal(isRecord([]), false);
  assert.equal(isRecord(null), false);
  assert.equal(isRecord('x'), false);

  assert.equal(requireString({ a: 'x' }, 'a', URL_), 'x');
  assert.throws(() => requireString({ a: '' }, 'a', URL_), GraphResponseError);
  assert.throws(() => requireString({ a: 1 }, 'a', URL_), GraphResponseError);

  assert.equal(optionalString({ a: 'x' }, 'a'), 'x');
  assert.equal(optionalString({ a: '' }, 'a'), '', 'empty is a value, not an absence');
  assert.equal(optionalString({ a: null }, 'a'), undefined);
  assert.equal(optionalString({}, 'a'), undefined);
});

test('quoteOData doubles an embedded quote', () => {
  assert.equal(quoteOData('Monthly Log'), "'Monthly Log'");
  assert.equal(quoteOData("Dov's log"), "'Dov''s log'");
  assert.equal(quoteOData("''"), "''''''");
  assert.equal(quoteOData(''), "''");
});

test('mapWithLimit preserves input order and never exceeds the limit', async () => {
  let running = 0;
  let peak = 0;
  const release: (() => void)[] = [];

  const results = mapWithLimit([1, 2, 3, 4, 5, 6], 2, async (n) => {
    running += 1;
    peak = Math.max(peak, running);
    await new Promise<void>((resolve) => release.push(resolve));
    running -= 1;
    return n * 10;
  });

  // Let every queued task through, a batch at a time.
  for (let i = 0; i < 6; i += 1) {
    await Promise.resolve();
    release.shift()?.();
  }
  while (release.length > 0) {
    await Promise.resolve();
    release.shift()?.();
  }

  assert.deepEqual(await results, [10, 20, 30, 40, 50, 60]);
  assert.equal(peak, 2, "two workers ran, and never a third");
});

test('mapWithLimit handles an empty list and a limit above the list length', async () => {
  assert.deepEqual(await mapWithLimit([], 4, async () => 1), []);
  assert.deepEqual(await mapWithLimit([1, 2], 10, async (n) => n + 1), [2, 3]);
});

test('safeText answers empty rather than masking the status that caused the throw', async () => {
  const readable = { text: () => Promise.resolve('a body') } as unknown as Response;
  assert.equal(await safeText(readable), 'a body');

  const broken = { text: () => Promise.reject(new Error('stream closed')) } as unknown as Response;
  assert.equal(await safeText(broken), '');
});

test('truncate says how much it cut, and describeError names the error type', () => {
  assert.equal(truncate('abc', 10), 'abc');
  assert.equal(truncate('abcdef', 3), 'abc… (6 chars)');

  assert.equal(describeError(new TypeError('bad')), 'TypeError: bad');
  assert.equal(describeError('a string'), 'a string');
  assert.equal(describeError(undefined), 'undefined');
});

test('no decode failure quotes the value that failed', () => {
  // Every one of these carries a plausible piece of the user's own writing in the place
  // the decoder rejects. The rule is that a message says which key was unusable and
  // never what was in it — these bodies are notebook, section and page names, and this
  // repository's output can reach a public log.
  const secret = 'Therapy notes 2026';
  const cases: (() => unknown)[] = [
    () => toNode({ displayName: secret }, URL_),
    () => toNode({ id: '', displayName: secret }, URL_),
    () => toPageSummary({ title: secret }, URL_),
    () => asRecord(secret, URL_),
    () => toNodeArray({ displayName: secret }, URL_),
    () => requireString({ a: 1, name: secret }, 'a', URL_),
    () => toSectionWithParents({ displayName: secret }, URL_),
    () => toExpandedNotebook({ displayName: secret }, URL_),
  ];

  for (const run of cases) {
    assert.throws(run, (err: unknown) => {
      assert.ok(err instanceof GraphResponseError);
      assert.ok(
        !err.message.includes(secret),
        `message leaked the value: ${err.message}`,
      );
      return true;
    });
  }
});
