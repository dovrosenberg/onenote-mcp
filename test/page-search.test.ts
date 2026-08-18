// The search half of the structure tools, driven through a fake structure client.
//
// The fake is a plain object rather than a GraphStructure over a fake fetch: what is
// under test here is the walk, the bounds, and the reporting, and routing URLs would
// only restate what test/graph-structure.test.ts already asserts.

import test from 'node:test';
import assert from 'node:assert/strict';

import type { ExpandedNotebook, PageSummary } from '../src/graph-structure.ts';
import {
  flattenSections,
  searchAllSections,
  searchOneSection,
  titleMatches,
  toContainerKind,
  type SearchableStructure,
} from '../src/page-search.ts';

function page(id: string, title: string, modified: string): PageSummary {
  return { id, title, lastModifiedDateTime: modified };
}

/**
 * The tree as `getExpandedTree` returns it: a notebook's own sections, and its section
 * groups with their sections. That is two levels, which is Graph's cap on `$expand`
 * nesting — `sec-week`, in a group below a group, is deliberately absent, because an
 * unscoped search does not reach it and the counts have to say so.
 */
const TREE: ExpandedNotebook[] = [
  {
    id: 'nb-2026',
    displayName: '2026',
    sections: [{ id: 'sec-inbox', displayName: 'Inbox' }],
    sectionGroups: [
      {
        id: 'grp-march',
        displayName: 'March',
        sections: [
          { id: 'sec-daily', displayName: 'Daily todo' },
          { id: 'sec-week', displayName: 'Weekly todo' },
        ],
      },
    ],
  },
  {
    id: 'nb-2025',
    displayName: '2025',
    sections: [{ id: 'sec-archive', displayName: 'Archive' }],
    sectionGroups: [],
  },
];

const PAGES: Record<string, PageSummary[]> = {
  'sec-inbox': [page('p-1', 'Budget review', '2026-03-04T10:00:00Z')],
  'sec-daily': [
    page('p-2', 'budget notes', '2026-03-05T10:00:00Z'),
    page('p-3', 'Standup', '2026-03-06T10:00:00Z'),
  ],
  'sec-week': [page('p-4', 'BUDGET planning', '2026-03-07T10:00:00Z')],
  'sec-archive': [page('p-5', 'Old budget', '2025-12-01T10:00:00Z')],
};

interface Fake extends SearchableStructure {
  readonly listed: string[];
  readonly maxInFlight: () => number;
}

function fakeStructure(
  pages: Record<string, PageSummary[]> = PAGES,
  hooks: { onList?: (sectionId: string) => void } = {},
): Fake {
  const listed: string[] = [];
  let inFlight = 0;
  let peak = 0;

  return {
    listed,
    maxInFlight: () => peak,
    getExpandedTree: () => Promise.resolve(TREE),
    // Graph applies the title filter, so the fake does too: returning every page here
    // would test a matching rule the service no longer runs.
    findPagesMatchingTitle: async (sectionId, query) => {
      listed.push(sectionId);
      inFlight += 1;
      peak = Math.max(peak, inFlight);
      try {
        hooks.onList?.(sectionId);
        // A real request is a round trip; yielding here lets the concurrency bound show.
        await Promise.resolve();
        return (pages[sectionId] ?? []).filter((item) => titleMatches(item.title, query));
      } finally {
        inFlight -= 1;
      }
    },
  };
}

test('toContainerKind maps the caller-facing names and rejects anything else', () => {
  assert.equal(toContainerKind('notebook'), 'notebooks');
  assert.equal(toContainerKind('sectionGroup'), 'sectionGroups');
  for (const value of ['section', 'notebooks', 'NOTEBOOK', '']) {
    assert.equal(toContainerKind(value), null);
  }
});

test('titleMatches is a case-insensitive substring, and ignores surrounding blanks', () => {
  assert.ok(titleMatches('Budget review', 'budget'));
  assert.ok(titleMatches('budget notes', '  BUDGET '));
  assert.ok(!titleMatches('Standup', 'budget'));
});

test('flattenSections covers both levels of the expanded tree and builds a path', () => {
  const refs = flattenSections(TREE);

  assert.deepEqual(
    refs.map((ref) => ref.path),
    [
      '2026 / Inbox',
      '2026 / March / Daily todo',
      '2026 / March / Weekly todo',
      '2025 / Archive',
    ],
  );
  assert.deepEqual(
    refs.map((ref) => ref.id),
    ['sec-inbox', 'sec-daily', 'sec-week', 'sec-archive'],
  );
});

test('a scoped search reads one section and reports it as complete', async () => {
  const fake = fakeStructure();
  const result = await searchOneSection(fake, 'sec-daily', 'budget');

  assert.deepEqual(fake.listed, ['sec-daily']);
  assert.deepEqual(
    result.matches.map((match) => match.pageId),
    ['p-2'],
  );
  // No walk happened, so there is no path to report.
  assert.equal(result.matches[0]?.sectionPath, undefined);
  assert.equal(result.matches[0]?.sectionId, 'sec-daily');
  assert.equal(result.stoppedEarly, false);
  assert.equal(result.stoppedBecause, null);
});

test('an unscoped search covers every section and tags each match with its path', async () => {
  const fake = fakeStructure();
  const result = await searchAllSections(fake, 'budget');

  assert.equal(result.sectionsSearched, 4);
  assert.equal(result.sectionsFound, 4);
  assert.equal(result.stoppedEarly, false);
  assert.equal(result.stoppedBecause, null);

  // Newest first, and each match carries the path of the section it came from.
  assert.deepEqual(
    result.matches.map((match) => match.pageId),
    ['p-4', 'p-2', 'p-1', 'p-5'],
  );
  assert.equal(
    result.matches.find((match) => match.pageId === 'p-4')?.sectionPath,
    '2026 / March / Weekly todo',
  );
});

test('the account-wide walk never lists more sections at once than the bound allows', async () => {
  const fake = fakeStructure();
  await searchAllSections(fake, 'budget', { concurrency: 2 });
  assert.ok(fake.maxInFlight() <= 2, `listed ${fake.maxInFlight()} sections at once`);
});

test('the section limit stops the walk and says so', async () => {
  const fake = fakeStructure();
  const result = await searchAllSections(fake, 'budget', { maxSections: 2, concurrency: 1 });

  assert.equal(result.sectionsSearched, 2);
  assert.equal(result.sectionsFound, 4);
  assert.equal(result.stoppedEarly, true);
  assert.equal(result.stoppedBecause, 'section-limit');
  assert.deepEqual(fake.listed, ['sec-inbox', 'sec-daily']);
});

test('the time budget stops the walk and outranks the section limit', async () => {
  let clock = 0;
  const fake = fakeStructure(PAGES, { onList: () => (clock += 10) });

  const result = await searchAllSections(fake, 'budget', {
    timeBudgetMs: 15,
    concurrency: 1,
    now: () => clock,
  });

  // Two sections fit inside the budget; the third is not started.
  assert.equal(result.sectionsSearched, 2);
  assert.equal(result.stoppedEarly, true);
  assert.equal(result.stoppedBecause, 'time-budget');
});

test('the match cap keeps the newest matches and reports the full count', async () => {
  const result = await searchAllSections(fakeStructure(), 'budget', { maxMatches: 2 });

  assert.deepEqual(
    result.matches.map((match) => match.pageId),
    ['p-4', 'p-2'],
  );
  assert.equal(result.totalMatches, 4);
  // Every section was still searched: the cap is on what is returned, not on the walk.
  assert.equal(result.stoppedEarly, false);
});

test('a search with no matches is still reported as complete', async () => {
  const result = await searchAllSections(fakeStructure(), 'nothing matches this');

  assert.deepEqual(result.matches, []);
  assert.equal(result.stoppedEarly, false);
  assert.equal(result.sectionsSearched, 4);
});

test('a failure listing one section fails the search rather than being swallowed', async () => {
  const fake = fakeStructure(PAGES, {
    onList: (sectionId) => {
      if (sectionId === 'sec-daily') throw new Error('graph said no');
    },
  });

  await assert.rejects(searchAllSections(fake, 'budget', { concurrency: 1 }), /graph said no/);
});
