// The name resolver, driven through a plain fake `LookupStructure`.
//
// The fake counts calls, because the whole point of this module is what it does not do:
// the common path is one `getExpandedTree` and no container walk at all. A regression
// that reintroduces the walk passes every assertion about ids and fails the call counts.
//
// The fixture nests a section group inside a section group, which the expanded tree
// cannot reach — that level is what the fallback exists for, and the real account has
// none of it, so nothing but a test exercises that path.

import test from 'node:test';
import assert from 'node:assert/strict';

import type { ExpandedNotebook, SectionWithParents } from '../src/graph-structure.ts';
import {
  NameLookupError,
  matchNodes,
  matchOne,
  namesMatch,
  resolveSection,
  withoutOrderingPrefix,
  type LookupStructure,
} from '../src/name-lookup.ts';

const EXPANDED: ExpandedNotebook[] = [
  {
    id: 'nb-2026',
    displayName: 'Bullet Journal - 2026',
    sections: [
      { id: 'sec-inbox', displayName: 'Inbox' },
      { id: 'sec-dup', displayName: 'Notes' },
      { id: 'sec-dup2', displayName: 'notes' },
    ],
    sectionGroups: [
      {
        id: 'grp-feb',
        displayName: '062 - February',
        sections: [{ id: 'sec-log', displayName: 'Monthly Log' }],
      },
      { id: 'grp-deep', displayName: 'Archive', sections: [] },
    ],
  },
  {
    id: 'nb-2025',
    displayName: 'Bullet Journal - 2025',
    sections: [{ id: 'sec-inbox-2025', displayName: 'Inbox' }],
    sectionGroups: [],
  },
];

/**
 * What the account-wide filtered lookup returns. `Buried` sits under a section group
 * nested inside `Archive`, which is past what the expanded tree reaches — the only way
 * to it is this call, and its `parentSectionGroup` is the nested group rather than the
 * one the caller named.
 */
const BY_NAME: Record<string, SectionWithParents[]> = {
  Buried: [
    {
      id: 'sec-buried',
      displayName: 'Buried',
      parentNotebook: { id: 'nb-2026', displayName: 'Bullet Journal - 2026' },
      parentSectionGroup: { id: 'grp-deeper', displayName: 'Archive' },
    },
  ],
  Elsewhere: [
    {
      id: 'sec-elsewhere',
      displayName: 'Elsewhere',
      parentNotebook: { id: 'nb-2025', displayName: 'Bullet Journal - 2025' },
      parentSectionGroup: { id: 'grp-other', displayName: 'Archive' },
    },
  ],
};

interface Fake extends LookupStructure {
  readonly calls: { tree: number; byName: string[] };
}

function fake(tree: ExpandedNotebook[] = EXPANDED): Fake {
  const calls = { tree: 0, byName: [] as string[] };
  return {
    calls,
    getExpandedTree: () => {
      calls.tree += 1;
      return Promise.resolve(tree);
    },
    findSectionsByName: (displayName: string) => {
      calls.byName.push(displayName);
      return Promise.resolve(BY_NAME[displayName] ?? []);
    },
  };
}

test('names match in full and ignore case and surrounding whitespace', () => {
  assert.equal(namesMatch('Monthly Log', 'monthly log'), true);
  assert.equal(namesMatch('Monthly Log', '  Monthly Log '), true);
  assert.equal(namesMatch('Monthly Log', 'Monthly'), false);
  assert.equal(namesMatch('Monthly Log', 'Log'), false);
  assert.equal(namesMatch('Monthly Log', ''), false);
});

test('matchOne returns only id and display name, whatever the node carried', () => {
  const matched = matchOne(EXPANDED, 'bullet journal - 2026', 'notebookName');
  assert.deepEqual(matched.node, { id: 'nb-2026', displayName: 'Bullet Journal - 2026' });
  assert.equal(matched.rule, 'exact');
});

test('a section inside a named group resolves in exactly one request', async () => {
  const structure = fake();
  const resolved = await resolveSection(structure, {
    notebookName: 'bullet journal - 2026',
    sectionGroupName: '062 - february',
    sectionName: 'monthly log',
  });

  assert.deepEqual(resolved.notebook, { id: 'nb-2026', displayName: 'Bullet Journal - 2026' });
  assert.deepEqual(resolved.sectionGroup, { id: 'grp-feb', displayName: '062 - February' });
  assert.deepEqual(resolved.section, { id: 'sec-log', displayName: 'Monthly Log' });
  assert.equal(resolved.deepSearchUsed, false);

  assert.equal(structure.calls.tree, 1);
  assert.deepEqual(structure.calls.byName, [], 'the common path costs nothing beyond the tree');
});

test('a notebook-level section resolves with no group named', async () => {
  const structure = fake();
  const resolved = await resolveSection(structure, {
    notebookName: 'Bullet Journal - 2026',
    sectionName: 'Inbox',
  });

  assert.deepEqual(resolved.section, { id: 'sec-inbox', displayName: 'Inbox' });
  assert.equal(resolved.sectionGroup, null);
  assert.deepEqual(structure.calls.byName, []);
});

test('omitting the group name does not reach a section inside a group', async () => {
  // 'Monthly Log' exists, one level down. Finding it anyway would answer a question the
  // caller did not ask, and would make the argument meaningless.
  await assert.rejects(
    () =>
      resolveSection(fake(), {
        notebookName: 'Bullet Journal - 2026',
        sectionName: 'Monthly Log',
      }),
    (err: unknown) => err instanceof NameLookupError && err.argument === 'sectionName',
  );
});

test('two sections with the same name are reported, not resolved', async () => {
  const error = await resolveSection(fake(), {
    notebookName: 'Bullet Journal - 2026',
    sectionName: 'Notes',
  }).catch((err: unknown) => err);

  assert.ok(error instanceof NameLookupError);
  assert.equal(error.kind, 'ambiguous');
  assert.equal(error.argument, 'sectionName');
  assert.deepEqual(
    error.candidates.map((node) => node.id),
    ['sec-dup', 'sec-dup2'],
  );
  assert.match(error.message, /matched 2 of them/);
});

test('a missing notebook name lists the notebooks that were there', async () => {
  const error = await resolveSection(fake(), {
    notebookName: 'Bullet Journal - 2024',
    sectionName: 'Inbox',
  }).catch((err: unknown) => err);

  assert.ok(error instanceof NameLookupError);
  assert.equal(error.kind, 'not-found');
  assert.equal(error.argument, 'notebookName');
  assert.match(error.message, /Bullet Journal - 2026/);
  assert.match(error.message, /matched in full ignoring case/);
});

test('a section nested below the named group is found by one filtered request', async () => {
  const structure = fake();
  const resolved = await resolveSection(structure, {
    notebookName: 'Bullet Journal - 2026',
    sectionGroupName: 'Archive',
    sectionName: 'Buried',
  });

  assert.deepEqual(resolved.section, { id: 'sec-buried', displayName: 'Buried' });
  assert.equal(resolved.deepSearchUsed, true, 'the caller is told this took the other path');
  assert.deepEqual(structure.calls.byName, ['Buried'], 'one request, not one per container');
});

test('the filtered fallback runs only when the expanded tree came back empty-handed', async () => {
  const structure = fake();
  await resolveSection(structure, {
    notebookName: 'Bullet Journal - 2026',
    sectionGroupName: '062 - February',
    sectionName: 'Monthly Log',
  });
  assert.deepEqual(structure.calls.byName, []);
});

test('a section of that name in another notebook is not accepted', async () => {
  // The filtered lookup is account-wide and this section group name exists twice, so the
  // notebook has to be checked as well as the group.
  const error = await resolveSection(fake(), {
    notebookName: 'Bullet Journal - 2026',
    sectionGroupName: 'Archive',
    sectionName: 'Elsewhere',
  }).catch((err: unknown) => err);

  assert.ok(error instanceof NameLookupError);
  assert.equal(error.kind, 'not-found');
  assert.equal(error.argument, 'sectionName');
});

test('a section missing everywhere below the group is a not-found', async () => {
  const error = await resolveSection(fake(), {
    notebookName: 'Bullet Journal - 2026',
    sectionGroupName: 'Archive',
    sectionName: 'Nowhere',
  }).catch((err: unknown) => err);

  assert.ok(error instanceof NameLookupError);
  assert.equal(error.argument, 'sectionName');
  assert.equal(error.kind, 'not-found');
});

test('an empty container yields an error that says there was nothing to match', async () => {
  const empty: ExpandedNotebook[] = [
    { id: 'nb', displayName: 'Empty', sections: [], sectionGroups: [] },
  ];
  const error = await resolveSection(fake(empty), {
    notebookName: 'Empty',
    sectionName: 'Inbox',
  }).catch((err: unknown) => err);

  assert.ok(error instanceof NameLookupError);
  assert.match(error.message, /nothing there to match/);
});

// ---------------------------------------------------------------------------
// The matching ladder. The case it exists for: this account names its section groups
// '062 - February', and a caller knows the month, not the number.
// ---------------------------------------------------------------------------

test('an ordering prefix comes off the stored name, in the forms people write one', () => {
  assert.equal(withoutOrderingPrefix('062 - February'), 'February');
  assert.equal(withoutOrderingPrefix('02. February'), 'February');
  assert.equal(withoutOrderingPrefix('2) February'), 'February');
  assert.equal(withoutOrderingPrefix('03 February'), 'February');
  assert.equal(withoutOrderingPrefix('2026: Review'), 'Review');
  assert.equal(withoutOrderingPrefix('February'), 'February');
  // A name that is only digits keeps them; stripping it would leave nothing to match on.
  assert.equal(withoutOrderingPrefix('2026'), '2026');
});

test('the ladder prefers an exact match over a looser one', () => {
  const nodes = [
    { id: 'a', displayName: 'February' },
    { id: 'b', displayName: '062 - February' },
    { id: 'c', displayName: 'February planning' },
  ];

  const matched = matchNodes(nodes, 'february');
  assert.equal(matched.rule, 'exact');
  assert.deepEqual(
    matched.matches.map((node) => node.id),
    ['a'],
    'the prefixed and the substring candidates must not compete with an exact match',
  );
});

test("a month name finds the section group that carries a number in front of it", async () => {
  const structure = fake();
  const resolved = await resolveSection(structure, {
    notebookName: 'Bullet Journal - 2026',
    sectionGroupName: 'February',
    sectionName: 'Monthly Log',
  });

  assert.deepEqual(resolved.sectionGroup, { id: 'grp-feb', displayName: '062 - February' });
  assert.equal(resolved.matchedBy.sectionGroup, 'without-prefix');
  assert.equal(resolved.matchedBy.notebook, 'exact');
  assert.equal(resolved.matchedBy.section, 'exact');
});

test('a partial name falls to the substring rung and says so', async () => {
  const resolved = await resolveSection(fake(), {
    notebookName: 'Bullet Journal - 2026',
    sectionGroupName: 'febr',
    sectionName: 'Monthly Log',
  });

  assert.deepEqual(resolved.sectionGroup, { id: 'grp-feb', displayName: '062 - February' });
  assert.equal(resolved.matchedBy.sectionGroup, 'substring');
});

test('two groups that tie on the same rung are refused rather than guessed', async () => {
  // Both strip to 'February', so the without-prefix rung matches both and neither is
  // more exact than the other. A group whose stripped name were 'February review'
  // would lose to this rung instead of competing with it.
  const tree: ExpandedNotebook[] = [
    {
      id: 'nb',
      displayName: 'Journal',
      sections: [],
      sectionGroups: [
        { id: 'g1', displayName: '062 - February', sections: [] },
        { id: 'g2', displayName: '07. February', sections: [] },
      ],
    },
  ];

  const error = await resolveSection(fake(tree), {
    notebookName: 'Journal',
    sectionGroupName: 'february',
    sectionName: 'anything',
  }).catch((err: unknown) => err);

  assert.ok(error instanceof NameLookupError);
  assert.equal(error.kind, 'ambiguous');
  assert.equal(error.argument, 'sectionGroupName');
});
