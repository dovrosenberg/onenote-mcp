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

import type { ExpandedNotebook, Section, SectionGroup } from '../src/graph-structure.ts';
import {
  NameLookupError,
  matchOne,
  namesMatch,
  resolveSection,
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

/** Only reachable by walking: the expanded tree stops at a group's own sections. */
const NESTED: Record<string, { sections: Section[]; sectionGroups: SectionGroup[] }> = {
  'grp-deep': {
    sections: [],
    sectionGroups: [{ id: 'grp-deeper', displayName: 'Old' }],
  },
  'grp-deeper': {
    sections: [{ id: 'sec-buried', displayName: 'Buried' }],
    sectionGroups: [],
  },
};

interface Fake extends LookupStructure {
  readonly calls: { tree: number; children: string[] };
}

function fake(tree: ExpandedNotebook[] = EXPANDED): Fake {
  const calls = { tree: 0, children: [] as string[] };
  return {
    calls,
    getExpandedTree: () => {
      calls.tree += 1;
      return Promise.resolve(tree);
    },
    listContainerChildren: (_kind, containerId) => {
      calls.children.push(containerId);
      return Promise.resolve(NESTED[containerId] ?? { sections: [], sectionGroups: [] });
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
  assert.deepEqual(matched, { id: 'nb-2026', displayName: 'Bullet Journal - 2026' });
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
  assert.deepEqual(structure.calls.children, [], 'the common path walks no containers');
});

test('a notebook-level section resolves with no group named', async () => {
  const structure = fake();
  const resolved = await resolveSection(structure, {
    notebookName: 'Bullet Journal - 2026',
    sectionName: 'Inbox',
  });

  assert.deepEqual(resolved.section, { id: 'sec-inbox', displayName: 'Inbox' });
  assert.equal(resolved.sectionGroup, null);
  assert.deepEqual(structure.calls.children, []);
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
  assert.match(error.message, /matched in full, ignoring case/);
});

test('a section nested below the named group is found by the fallback walk', async () => {
  const structure = fake();
  const resolved = await resolveSection(structure, {
    notebookName: 'Bullet Journal - 2026',
    sectionGroupName: 'Archive',
    sectionName: 'Buried',
  });

  assert.deepEqual(resolved.section, { id: 'sec-buried', displayName: 'Buried' });
  assert.equal(resolved.deepSearchUsed, true, 'the caller is told this cost extra requests');
  // Breadth-first from the named group: the group itself, then the one below it.
  assert.deepEqual(structure.calls.children, ['grp-deep', 'grp-deeper']);
});

test('the fallback walk runs only when the expanded tree came back empty-handed', async () => {
  const structure = fake();
  await resolveSection(structure, {
    notebookName: 'Bullet Journal - 2026',
    sectionGroupName: '062 - February',
    sectionName: 'Monthly Log',
  });
  assert.deepEqual(structure.calls.children, []);
});

test('a section missing everywhere below the group lists what the walk saw', async () => {
  const error = await resolveSection(fake(), {
    notebookName: 'Bullet Journal - 2026',
    sectionGroupName: 'Archive',
    sectionName: 'Nowhere',
  }).catch((err: unknown) => err);

  assert.ok(error instanceof NameLookupError);
  assert.equal(error.argument, 'sectionName');
  assert.match(error.message, /Buried/);
  // The named group's own sections and the walked ones are listed once, not twice.
  assert.equal(error.message.match(/Buried/g)?.length, 1);
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
