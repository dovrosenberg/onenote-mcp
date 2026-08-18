// The four browsing tools, driven through their own `handle` with a fake structure
// client. What is asserted is the contract a calling model sees: the declared schemas,
// the JSON each tool answers with, and what a bounded search says about itself.

import test from 'node:test';
import assert from 'node:assert/strict';

import type { CallToolResult } from '@modelcontextprotocol/sdk/types.js';

import type {
  ContainerChildren,
  ContainerKind,
  Notebook,
  NotebookTree,
  PageSummary,
} from '../src/graph-structure.ts';
import { ToolInputError, indexTools, type ToolDefinition } from '../src/mcp-tools.ts';
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

const PAGES: Record<string, PageSummary[]> = {
  'sec-inbox': [
    { id: 'p-1', title: 'Budget review', lastModifiedDateTime: '2026-03-04T10:00:00Z' },
    { id: 'p-2', title: 'Standup', lastModifiedDateTime: '2026-03-03T10:00:00Z' },
  ],
  'sec-daily': [{ id: 'p-3', title: 'budget notes', lastModifiedDateTime: '2026-03-05T10:00:00Z' }],
};

interface Fake extends StructureClient {
  readonly containerCalls: string[];
  readonly pageCalls: { sectionId: string; top: number | undefined }[];
}

function fakeStructure(): Fake {
  const containerCalls: string[] = [];
  const pageCalls: { sectionId: string; top: number | undefined }[] = [];

  return {
    containerCalls,
    pageCalls,
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
    getFullTree: () => Promise.resolve(TREE),
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

test('the four browsing tools are registered under the names the spec gives them', () => {
  const tools = createStructureTools(fakeStructure());
  assert.deepEqual(
    tools.map((tool) => tool.name),
    ['list_notebooks', 'list_sections', 'list_pages', 'search_pages'],
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

  assert.deepEqual(
    structure.pageCalls.map((entry) => entry.sectionId).sort(),
    ['sec-daily', 'sec-inbox'],
  );
  assert.equal(body['scope'], 'account');
  assert.equal(body['sectionsSearched'], 2);
  assert.equal(body['sectionsFound'], 2);
  assert.equal(body['stoppedEarly'], false);
  assert.equal(body['stoppedBecause'], null);
  assert.match(String(body['note']), /Searched all 2/);

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
  assert.equal(body['sectionsFound'], 2);
  assert.match(String(body['note']), /1 of 2 sections/);
  assert.match(String(body['note']), /sectionId/);
});

test('search_pages requires a query', async () => {
  const tool = byName(createStructureTools(fakeStructure()), 'search_pages');
  await assert.rejects(tool.handle({}), ToolInputError);
  await assert.rejects(tool.handle({ query: '   ' }), ToolInputError);
});
