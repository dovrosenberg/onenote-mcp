// The four browsing tools, driven through their own `handle` with a fake structure
// client. What is asserted is the contract a calling model sees: the declared schemas,
// the JSON each tool answers with, and what a bounded search says about itself.

import test from 'node:test';
import assert from 'node:assert/strict';

import type { CallToolResult } from '@modelcontextprotocol/sdk/types.js';

import type {
  ContainerChildren,
  ContainerKind,
  ExpandedNotebook,
  Notebook,
  NotebookTree,
  PageSummary,
} from '../src/graph-structure.ts';
import { ToolInputError, indexTools, type ToolDefinition } from '../src/mcp-tools.ts';
import { NameLookupError } from '../src/name-lookup.ts';
import { createStructureTools, type StructureClient } from '../src/structure-tools.ts';

const NOTEBOOKS: Notebook[] = [
  { id: 'nb-2026', displayName: '2026' },
  { id: 'nb-2025', displayName: '2025' },
];

const CHILDREN: Record<string, ContainerChildren> = {
  'notebooks:nb-2026': {
    sections: [{ id: 'sec-inbox', displayName: 'Inbox' }],
    sectionGroups: [{ id: 'grp-march', displayName: 'March' }],
  },
  'sectionGroups:grp-march': {
    sections: [{ id: 'sec-daily', displayName: 'Daily todo' }],
    sectionGroups: [],
  },
};

const TREE: NotebookTree[] = [
  {
    id: 'nb-2026',
    displayName: '2026',
    sections: [{ id: 'sec-inbox', displayName: 'Inbox' }],
    sectionGroups: [
      {
        id: 'grp-march',
        displayName: 'March',
        sections: [{ id: 'sec-daily', displayName: 'Daily todo' }],
        sectionGroups: [],
      },
    ],
  },
];

/**
 * The one-request tree the `_by_name` tools resolve against. It carries a notebook-level
 * section, a section group with its own sections, and a second notebook whose section
 * shares a name with the first — the ambiguity the lookup must refuse to resolve is
 * across notebooks, and a lookup scoped to one notebook must not trip over it.
 */
const EXPANDED: ExpandedNotebook[] = [
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
          { id: 'sec-log', displayName: 'Monthly Log' },
        ],
      },
      { id: 'grp-april', displayName: 'April', sections: [] },
    ],
  },
  {
    id: 'nb-2025',
    displayName: '2025',
    sections: [{ id: 'sec-inbox-2025', displayName: 'Inbox' }],
    sectionGroups: [],
  },
];

const PAGES: Record<string, PageSummary[]> = {
  'sec-inbox': [
    { id: 'p-1', title: 'Budget review', lastModifiedDateTime: '2026-03-04T10:00:00Z' },
    { id: 'p-2', title: 'Standup', lastModifiedDateTime: '2026-03-03T10:00:00Z' },
  ],
  'sec-daily': [{ id: 'p-3', title: 'budget notes', lastModifiedDateTime: '2026-03-05T10:00:00Z' }],
  'sec-log': [
    { id: 'p-log', title: 'Monthly Log', lastModifiedDateTime: '2026-03-06T10:00:00Z' },
    { id: 'p-other', title: 'monthly log archive', lastModifiedDateTime: '2026-03-02T10:00:00Z' },
  ],
  'sec-nested': [{ id: 'p-4', title: 'Deep', lastModifiedDateTime: '2026-03-07T10:00:00Z' }],
};

interface Fake extends StructureClient {
  readonly containerCalls: string[];
  readonly pageCalls: { sectionId: string; top: number | undefined }[];
  readonly treeCalls: { expanded: number; full: number };
  readonly titleCalls: { sectionId: string; title: string }[];
}

function fakeStructure(): Fake {
  const containerCalls: string[] = [];
  const pageCalls: { sectionId: string; top: number | undefined }[] = [];
  const treeCalls = { expanded: 0, full: 0 };
  const titleCalls: { sectionId: string; title: string }[] = [];

  return {
    containerCalls,
    pageCalls,
    treeCalls,
    titleCalls,
    getExpandedTree: () => {
      treeCalls.expanded += 1;
      return Promise.resolve(EXPANDED);
    },
    findSectionsByName: () => Promise.resolve([]),
    findPagesMatchingTitle: (sectionId: string, query: string) => {
      pageCalls.push({ sectionId, top: undefined });
      return Promise.resolve(
        (PAGES[sectionId] ?? []).filter((page) =>
          page.title.toLocaleLowerCase().includes(query.trim().toLocaleLowerCase()),
        ),
      );
    },
    findPagesByTitle: (sectionId: string, title: string) => {
      titleCalls.push({ sectionId, title });
      // Graph does this comparison; the fake does the same thing the service would.
      return Promise.resolve(
        (PAGES[sectionId] ?? []).filter(
          (page) => page.title.trim().toLowerCase() === title.trim().toLowerCase(),
        ),
      );
    },
    listNotebooks: () => Promise.resolve(NOTEBOOKS),
    listContainerChildren: (kind: ContainerKind, containerId: string) => {
      containerCalls.push(`${kind}:${containerId}`);
      return Promise.resolve(
        CHILDREN[`${kind}:${containerId}`] ?? { sections: [], sectionGroups: [] },
      );
    },
    listPagesInSection: (sectionId: string, top?: number) => {
      pageCalls.push({ sectionId, top });
      return Promise.resolve(PAGES[sectionId] ?? []);
    },
    getFullTree: () => {
      treeCalls.full += 1;
      return Promise.resolve(TREE);
    },
  };
}

function byName(tools: ToolDefinition[], name: string): ToolDefinition {
  const tool = tools.find((candidate) => candidate.name === name);
  assert.ok(tool !== undefined, `no tool named ${name}`);
  return tool;
}

/** Every tool here answers with one JSON text block; this is that object. */
function payload(result: CallToolResult): Record<string, unknown> {
  assert.notEqual(result.isError, true, JSON.stringify(result.content));
  const first = result.content[0];
  assert.equal(first?.type, 'text');
  return JSON.parse(first?.type === 'text' ? first.text : '{}') as Record<string, unknown>;
}

async function call(
  name: string,
  args: Record<string, unknown> = {},
  structure: StructureClient = fakeStructure(),
): Promise<Record<string, unknown>> {
  return payload(await byName(createStructureTools(structure), name).handle(args));
}

test('the browsing tools are registered under the names the spec gives them', () => {
  const tools = createStructureTools(fakeStructure());
  assert.deepEqual(
    tools.map((tool) => tool.name),
    [
      'list_notebooks',
      'list_sections',
      'list_pages',
      'search_pages',
      'find_page_by_name',
      'list_pages_by_name',
    ],
  );
  // Duplicate names would shadow silently in the JSON-RPC layer.
  assert.doesNotThrow(() => indexTools(tools));
});

test('every browsing tool declares itself read-only and describes its arguments', () => {
  for (const tool of createStructureTools(fakeStructure())) {
    assert.equal(tool.annotations?.readOnlyHint, true, tool.name);
    assert.ok(tool.description.length > 100, `${tool.name} has a thin description`);
    // A model that invents an argument should be told, not silently ignored.
    assert.equal(tool.inputSchema['additionalProperties'], false, tool.name);
  }
});

test('list_notebooks returns each notebook with the id list_sections takes', async () => {
  const body = await call('list_notebooks');
  assert.deepEqual(body['notebooks'], [
    { id: 'nb-2026', displayName: '2026' },
    { id: 'nb-2025', displayName: '2025' },
  ]);
  assert.equal(body['count'], 2);
});

test('list_sections returns sections and section groups in one tagged list', async () => {
  const structure = fakeStructure();
  const body = await call(
    'list_sections',
    { containerType: 'notebook', containerId: 'nb-2026' },
    structure,
  );

  assert.deepEqual(structure.containerCalls, ['notebooks:nb-2026']);
  assert.deepEqual(body['children'], [
    { type: 'section', id: 'sec-inbox', displayName: 'Inbox' },
    { type: 'sectionGroup', id: 'grp-march', displayName: 'March' },
  ]);
  assert.equal(body['sectionCount'], 1);
  assert.equal(body['sectionGroupCount'], 1);
});

test('list_sections descends into a section group with the same call', async () => {
  const structure = fakeStructure();
  const body = await call(
    'list_sections',
    { containerType: 'sectionGroup', containerId: 'grp-march' },
    structure,
  );

  assert.deepEqual(structure.containerCalls, ['sectionGroups:grp-march']);
  assert.deepEqual(body['children'], [
    { type: 'section', id: 'sec-daily', displayName: 'Daily todo' },
  ]);
});

test('list_sections rejects a containerType it does not know, naming the argument', async () => {
  const tool = byName(createStructureTools(fakeStructure()), 'list_sections');
  await assert.rejects(
    tool.handle({ containerType: 'section', containerId: 'sec-inbox' }),
    (err: unknown) =>
      err instanceof ToolInputError && err.argument === 'containerType' && /sectionGroup/.test(err.message),
  );
});

test('list_pages defaults top and reports that more pages exist behind it', async () => {
  const structure = fakeStructure();
  const full = await call('list_pages', { sectionId: 'sec-inbox', top: 2 }, structure);

  assert.deepEqual(structure.pageCalls, [{ sectionId: 'sec-inbox', top: 2 }]);
  assert.equal(full['count'], 2);
  // Exactly `top` pages came back, so the section holds at least one more.
  assert.equal(full['moreAvailable'], true);

  const defaulted = await call('list_pages', { sectionId: 'sec-inbox' });
  assert.equal(defaulted['top'], 50);
  assert.equal(defaulted['moreAvailable'], false);
});

test('list_pages rejects a top outside the declared range', async () => {
  const tool = byName(createStructureTools(fakeStructure()), 'list_pages');
  await assert.rejects(tool.handle({ sectionId: 'sec-inbox', top: 0 }), ToolInputError);
  await assert.rejects(tool.handle({ sectionId: 'sec-inbox', top: 500 }), ToolInputError);
  await assert.rejects(tool.handle({}), ToolInputError);
});

test('search_pages scoped to a section reads only that section', async () => {
  const structure = fakeStructure();
  const body = await call(
    'search_pages',
    { query: 'BUDGET', sectionId: 'sec-inbox' },
    structure,
  );

  assert.deepEqual(
    structure.pageCalls.map((entry) => entry.sectionId),
    ['sec-inbox'],
  );
  assert.equal(body['scope'], 'section');
  assert.deepEqual((body['matches'] as { pageId: string }[]).map((match) => match.pageId), ['p-1']);
  assert.equal(body['stoppedEarly'], false);
});

test('an unscoped search walks the tree and reports full coverage', async () => {
  const structure = fakeStructure();
  const body = await call('search_pages', { query: 'budget' }, structure);

  // Every section in the expanded tree: the notebooks' own sections and the sections of
  // their section groups.
  assert.deepEqual(
    structure.pageCalls.map((entry) => entry.sectionId).sort(),
    ['sec-daily', 'sec-inbox', 'sec-inbox-2025', 'sec-log'],
  );
  assert.equal(body['scope'], 'account');
  assert.equal(body['sectionsSearched'], 4);
  assert.equal(body['sectionsFound'], 4);
  assert.equal(body['stoppedEarly'], false);
  assert.equal(body['stoppedBecause'], null);
  assert.match(String(body['note']), /Searched all 4/);

  // Newest first, and each match says which section it is in.
  const matches = body['matches'] as { pageId: string; sectionPath: string }[];
  assert.deepEqual(matches.map((match) => match.pageId), ['p-3', 'p-1']);
  assert.equal(matches[0]?.sectionPath, '2026 / March / Daily todo');
});

test('a truncated search says so in the note rather than reading as an answer', async () => {
  const tools = createStructureTools(fakeStructure(), { maxSections: 1 });
  const body = payload(await byName(tools, 'search_pages').handle({ query: 'budget' }));

  assert.equal(body['stoppedEarly'], true);
  assert.equal(body['stoppedBecause'], 'section-limit');
  assert.equal(body['sectionsSearched'], 1);
  assert.equal(body['sectionsFound'], 4);
  assert.match(String(body['note']), /1 of 4 sections/);
  assert.match(String(body['note']), /sectionId/);
});

test('search_pages requires a query', async () => {
  const tool = byName(createStructureTools(fakeStructure()), 'search_pages');
  await assert.rejects(tool.handle({}), ToolInputError);
  await assert.rejects(tool.handle({ query: '   ' }), ToolInputError);
});

// ---------------------------------------------------------------------------
// The name-based lookups. What matters here is the tool surface: which client calls the
// resolution costs, what the result carries, and that a bad name is an error rather than
// an empty list. The matching rules themselves are covered in test/name-lookup.test.ts.
// ---------------------------------------------------------------------------

test('find_page_by_name matches names in full, ignoring case', async () => {
  const structure = fakeStructure();
  const body = await call(
    'find_page_by_name',
    {
      notebookName: '2026',
      sectionGroupName: 'march',
      sectionName: 'monthly log',
      pageTitle: 'MONTHLY LOG',
    },
    structure,
  );

  assert.deepEqual(body['notebook'], { id: 'nb-2026', displayName: '2026' });
  assert.deepEqual(body['sectionGroup'], { id: 'grp-march', displayName: 'March' });
  assert.deepEqual(body['section'], { id: 'sec-log', displayName: 'Monthly Log' });
  assert.equal(body['matchCount'], 1);
  assert.deepEqual(
    (body['matches'] as { id: string }[]).map((page) => page.id),
    ['p-log'],
    "'monthly log archive' is not a full-string match and must not come back",
  );

  // One tree request plus one filtered page request. The point of the tool is that it
  // walks nothing, so a regression to listContainerChildren or to reading every title in
  // the section shows up here.
  assert.equal(structure.treeCalls.expanded, 1);
  assert.deepEqual(structure.containerCalls, []);
  assert.deepEqual(structure.pageCalls, [], 'titles are filtered by Graph, not scanned');
  assert.deepEqual(structure.titleCalls, [{ sectionId: 'sec-log', title: 'MONTHLY LOG' }]);
});

test('a title that matched nothing is a complete answer, not a bounded one', async () => {
  const body = await call('find_page_by_name', {
    notebookName: '2026',
    sectionGroupName: 'March',
    sectionName: 'Monthly Log',
    pageTitle: 'Weekly Log',
  });

  assert.equal(body['matchCount'], 0);
  assert.deepEqual(body['matches'], []);
  assert.match(String(body['note']), /No page in that section has that title/);
});

test('a section name that matches nothing is an error listing what was there', async () => {
  const tool = byName(createStructureTools(fakeStructure()), 'find_page_by_name');
  const result = await tool
    .handle({
      notebookName: '2026',
      sectionGroupName: 'April',
      sectionName: 'Monthly Log',
      pageTitle: 'anything',
    })
    .catch((err: unknown) => err);

  assert.ok(result instanceof NameLookupError);
  assert.equal(result.argument, 'sectionName');
  assert.equal(result.kind, 'not-found');
});

test('a notebook name that matches two notebooks is refused, not guessed', async () => {
  const structure = fakeStructure();
  const ambiguous: StructureClient = {
    ...structure,
    getExpandedTree: () =>
      Promise.resolve([
        { id: 'nb-a', displayName: 'Journal', sections: [], sectionGroups: [] },
        { id: 'nb-b', displayName: 'journal', sections: [], sectionGroups: [] },
      ]),
  };

  const tool = byName(createStructureTools(ambiguous), 'list_pages_by_name');
  const result = await tool
    .handle({ notebookName: 'Journal', sectionName: 'Inbox' })
    .catch((err: unknown) => err);

  assert.ok(result instanceof NameLookupError);
  assert.equal(result.kind, 'ambiguous');
  assert.equal(result.argument, 'notebookName');
  assert.equal(result.candidates.length, 2);
});

test('list_pages_by_name lists a notebook-level section without a group name', async () => {
  const structure = fakeStructure();
  const body = await call('list_pages_by_name', { notebookName: '2026', sectionName: 'Inbox' }, structure);

  assert.deepEqual(body['section'], { id: 'sec-inbox', displayName: 'Inbox' });
  assert.equal(body['sectionGroup'], null);
  assert.equal(body['count'], 2);
  assert.equal(body['top'], 50);
  assert.equal(body['moreAvailable'], false);
  assert.deepEqual(structure.pageCalls, [{ sectionId: 'sec-inbox', top: 50 }]);
});

test('list_pages_by_name honours top and reports that more pages exist', async () => {
  const body = await call('list_pages_by_name', {
    notebookName: '2026',
    sectionName: 'Inbox',
    top: 2,
  });

  assert.equal(body['count'], 2);
  assert.equal(body['moreAvailable'], true);
});

test('omitting sectionGroupName does not search inside section groups', async () => {
  // 'Daily todo' exists, but inside a group. The lookup must not find it: a caller that
  // named no group said the section is a direct child of the notebook.
  const tool = byName(createStructureTools(fakeStructure()), 'list_pages_by_name');
  const result = await tool
    .handle({ notebookName: '2026', sectionName: 'Daily todo' })
    .catch((err: unknown) => err);

  assert.ok(result instanceof NameLookupError);
  assert.equal(result.argument, 'sectionName');
});

test('a name that is not a string is a ToolInputError naming the argument', async () => {
  const tool = byName(createStructureTools(fakeStructure()), 'find_page_by_name');
  await assert.rejects(
    () => tool.handle({ notebookName: '2026', sectionName: '', pageTitle: 'x' }),
    (err: unknown) => err instanceof ToolInputError && err.argument === 'sectionName',
  );
});
