// The seven covered tools with a mirror supplied, driven through their own `handle`.
//
// Separate from test/structure-tools.test.ts and test/page-tools.test.ts on purpose:
// those two build their tools with no mirror argument, and every assertion in them still
// holds unchanged. That is the property worth protecting — an absent mirror means "always
// Graph", which is what makes `MIRROR_READ_ENABLED=false` a complete rollback rather than
// a code path with its own bugs.
//
// Both fakes count their calls, because most of what matters here is what does *not*
// happen: `useLiveData` must not touch the mirror, a mirror hit must not touch Graph, and
// a mirror failure must still answer.

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
  SectionWithParents,
} from '../src/graph-structure.ts';
import { setEventSink } from '../src/logging.ts';
import type { MirrorBlobReader } from '../src/mirror-blobs.ts';
import type {
  MirrorNotebook,
  MirrorPage,
  MirrorPageContent,
  MirrorSection,
  MirrorSectionGroup,
} from '../src/mirror-schema.ts';
import { MirrorReader, type MirrorReadStore } from '../src/mirror-reader.ts';
import type { ScanResult } from '../src/mirror-store.ts';
import type { ToolDefinition } from '../src/mcp-tools.ts';
import { createPageTools, type PageContentClient } from '../src/page-tools.ts';
import { createStructureTools, type StructureClient } from '../src/structure-tools.ts';
import {
  createWriteTools,
  type MirrorWriteSync,
  type PageLayoutReader,
  type PageWriteClient,
  type WriteLookupStructure,
} from '../src/write-tools.ts';

setEventSink(() => {});

// ---------------------------------------------------------------------------
// Graph side: the same shape the existing tool tests use, with a call counter.
// ---------------------------------------------------------------------------

const GRAPH_NOTEBOOKS: Notebook[] = [
  { id: 'nb-2026', displayName: '2026' },
  { id: 'nb-2025', displayName: '2025' },
];

const GRAPH_PAGES: PageSummary[] = [
  { id: 'p-live', title: 'Live Monday', lastModifiedDateTime: '2026-08-19T12:00:00Z' },
];

// Graph knows a notebook the mirror has not synced yet, which is the realistic `_by_name`
// miss: a container created since the last run.
const GRAPH_EXPANDED: ExpandedNotebook[] = [
  {
    id: 'nb-2026',
    displayName: '2026',
    sections: [{ id: 'sec-daily', displayName: 'Daily' }],
    sectionGroups: [],
  },
  {
    id: 'nb-2027',
    displayName: '2027',
    sections: [{ id: 'sec-new', displayName: 'Brand new' }],
    sectionGroups: [],
  },
];

interface GraphCalls {
  total: number;
}

function fakeGraph(calls: GraphCalls): StructureClient {
  const count = <T>(value: T): Promise<T> => {
    calls.total += 1;
    return Promise.resolve(value);
  };

  return {
    listNotebooks: () => count(GRAPH_NOTEBOOKS),
    listContainerChildren: (_kind: ContainerKind, _id: string) =>
      count<ContainerChildren>({
        sections: [{ id: 'sec-live', displayName: 'Live section' }],
        sectionGroups: [],
      }),
    listPagesInSection: () => count(GRAPH_PAGES),
    getFullTree: () => count<NotebookTree[]>([]),
    getExpandedTree: () => count(GRAPH_EXPANDED),
    findSectionsByName: () => count<SectionWithParents[]>([]),
    findPagesByTitle: () => count(GRAPH_PAGES),
    findPagesMatchingTitle: () => count(GRAPH_PAGES),
  };
}

// ---------------------------------------------------------------------------
// Mirror side
// ---------------------------------------------------------------------------

const MIRROR_NOTEBOOKS: MirrorNotebook[] = [
  {
    id: 'nb-2026',
    displayName: '2026',
    mirrored: true,
    sectionCount: 1,
    sectionGroupCount: 0,
    graphLastModifiedDateTime: null,
  },
  {
    id: 'nb-2025',
    displayName: '2025',
    mirrored: false,
    sectionCount: 0,
    sectionGroupCount: 0,
    graphLastModifiedDateTime: null,
  },
];

const MIRROR_SECTIONS: MirrorSection[] = [
  {
    id: 'sec-daily',
    displayName: 'Daily',
    notebookId: 'nb-2026',
    parentId: 'nb-2026',
    parentKind: 'notebook',
    path: '2026 / Daily',
    mirrored: true,
    graphLastModifiedDateTime: null,
    pagesSyncedThrough: '2026-08-19T12:00:00Z',
    pageCount: 1,
  },
];

const MIRROR_PAGES: MirrorPage[] = [
  {
    id: 'p-held',
    title: 'Held Monday',
    titleLower: 'held monday',
    sectionId: 'sec-daily',
    notebookId: 'nb-2026',
    sectionPath: '2026 / Daily',
    lastModifiedDateTime: '2026-08-19T11:00:00Z',
    contentState: 'present',
    contentHash: 'abc',
    htmlLocation: 'firestore',
    htmlObject: null,
    htmlBytes: 12,
    ink: null,
    contentSyncedAt: '2026-08-19T12:30:00.000Z',
  },
];

interface MirrorCalls {
  total: number;
}

function fakeMirror(
  calls: MirrorCalls,
  options: { broken?: boolean; empty?: boolean } = {},
): MirrorReader {
  const count = <T>(value: T): Promise<T> => {
    calls.total += 1;
    if (options.broken === true) return Promise.reject(new Error('firestore down'));
    return Promise.resolve(value);
  };

  const groups: MirrorSectionGroup[] = [];
  const content = new Map<string, MirrorPageContent>([
    ['p-held', { pageId: 'p-held', html: '<p>held content</p>', bytes: 18, contentHash: 'abc' }],
  ]);

  const store: MirrorReadStore = {
    listNotebooks: () => count(options.empty === true ? [] : MIRROR_NOTEBOOKS),
    listSectionsUnder: (parentId) =>
      count(MIRROR_SECTIONS.filter((s) => s.parentId === parentId)),
    listSectionGroupsUnder: () => count(groups),
    listAllSections: () => count(options.empty === true ? [] : MIRROR_SECTIONS),
    listAllSectionGroups: () => count(groups),
    getNotebook: (id) => count(MIRROR_NOTEBOOKS.find((n) => n.id === id) ?? null),
    getSection: (id) => count(MIRROR_SECTIONS.find((s) => s.id === id) ?? null),
    getSectionGroup: () => count(null),
    getPage: (id) => count(MIRROR_PAGES.find((p) => p.id === id) ?? null),
    getPageContent: (id) => count(content.get(id) ?? null),
    listPagesInSection: (sectionId, limit) =>
      count(MIRROR_PAGES.filter((p) => p.sectionId === sectionId).slice(0, limit)),
    scanPages: () => count<ScanResult>({ pages: MIRROR_PAGES, truncated: false }),
  };

  const blobs: MirrorBlobReader = {
    getInk: () => Promise.resolve(null),
    getHtml: () => Promise.resolve(null),
  };

  return new MirrorReader(store, blobs);
}

// ---------------------------------------------------------------------------
// Write side: the narrowest fakes that let a write succeed.
// ---------------------------------------------------------------------------

function fakeWriteClient(): PageWriteClient {
  return {
    appendToPage: () => Promise.resolve(),
    updatePageTitle: () => Promise.resolve(),
    createPage: () =>
      Promise.resolve({ id: 'p-new', title: 'Fresh', webUrl: null, clientUrl: null }),
  };
}

function fakeLayoutReader(): PageLayoutReader {
  // No ink and no HTML: the append path reads the page to plan ink clearance, and a
  // typed page needs none.
  return { fetchRaw: () => Promise.resolve({ raw: '', contentType: null, parts: [] }) };
}

function fakeWriteLookup(): WriteLookupStructure {
  return {
    getExpandedTree: () => Promise.resolve(GRAPH_EXPANDED),
    findSectionsByName: () => Promise.resolve([]),
    findPagesByTitle: () => Promise.resolve([]),
  };
}

// ---------------------------------------------------------------------------
// Harness
// ---------------------------------------------------------------------------

function payload(result: CallToolResult): Record<string, unknown> {
  const block = result.content[0];
  assert.ok(block !== undefined && block.type === 'text');
  return JSON.parse(block.text) as Record<string, unknown>;
}

function byName(tools: ToolDefinition[], name: string): ToolDefinition {
  const tool = tools.find((candidate) => candidate.name === name);
  assert.ok(tool !== undefined, `no tool named ${name}`);
  return tool;
}

// Counted into the same tally as the structure client: for these assertions "reached
// Graph" is one fact, whichever client did it.
function fakeContent(calls: GraphCalls): PageContentClient {
  return {
    fetchContent: () => {
      calls.total += 1;
      return Promise.resolve({ html: '<p>live content</p>', ink: null });
    },
  };
}

function tools(
  graphCalls: GraphCalls,
  mirrorCalls: MirrorCalls,
  options: { broken?: boolean; empty?: boolean; noMirror?: boolean } = {},
): ToolDefinition[] {
  const mirror = options.noMirror === true ? undefined : fakeMirror(mirrorCalls, options);
  return [
    ...createStructureTools(fakeGraph(graphCalls), {}, mirror),
    ...createPageTools(fakeContent(graphCalls), undefined, mirror),
  ];
}

async function run(
  name: string,
  args: Record<string, unknown>,
  options: { broken?: boolean; empty?: boolean; noMirror?: boolean } = {},
): Promise<{ body: Record<string, unknown>; graph: number; mirror: number }> {
  const graphCalls: GraphCalls = { total: 0 };
  const mirrorCalls: MirrorCalls = { total: 0 };
  const body = payload(await byName(tools(graphCalls, mirrorCalls, options), name).handle(args));
  return { body, graph: graphCalls.total, mirror: mirrorCalls.total };
}

const COVERED = [
  ['list_notebooks', {}],
  ['list_sections', { containerType: 'notebook', containerId: 'nb-2026' }],
  ['list_pages', { sectionId: 'sec-daily' }],
  ['search_pages', { query: 'monday' }],
  ['find_page_by_name', { notebookName: '2026', sectionName: 'Daily', pageTitle: 'Held Monday' }],
  ['list_pages_by_name', { notebookName: '2026', sectionName: 'Daily' }],
  ['get_page_content', { pageId: 'p-held' }],
] as const;

// ---------------------------------------------------------------------------

test('every covered tool declares useLiveData, and none of the others do', () => {
  const all = tools({ total: 0 }, { total: 0 });
  const covered = new Set(COVERED.map(([name]) => name as string));

  for (const tool of all) {
    const properties = (tool.inputSchema.properties ?? {}) as Record<string, unknown>;
    assert.equal(
      'useLiveData' in properties,
      covered.has(tool.name),
      `${tool.name} has useLiveData: ${String('useLiveData' in properties)}`,
    );
  }

  // Seven, and the list above is the specification of which.
  assert.equal(all.filter((t) => 'useLiveData' in (t.inputSchema.properties ?? {})).length, 7);
});

test('useLiveData is never required, so the argument can always be omitted', () => {
  for (const tool of tools({ total: 0 }, { total: 0 })) {
    const required = (tool.inputSchema.required ?? []) as string[];
    assert.equal(required.includes('useLiveData'), false, tool.name);
  }
});

test('every covered tool answers from the mirror and reports it', async () => {
  for (const [name, args] of COVERED) {
    const { body, graph } = await run(name, args);
    assert.equal(body['source'], 'mirror', name);
    assert.equal(graph, 0, `${name} reached Graph on a mirror hit`);
  }
});

test('useLiveData true sends every covered tool to Graph, untouched by the mirror', async () => {
  for (const [name, args] of COVERED) {
    const { body, graph, mirror } = await run(name, { ...args, useLiveData: true });
    assert.equal(body['source'], 'graph', name);
    assert.ok(graph > 0, `${name} did not reach Graph`);
    // The point of the argument: a caller needing an edit from thirty seconds ago must
    // not pay a Firestore read to be told the mirror does not have it yet.
    assert.equal(mirror, 0, `${name} touched the mirror despite useLiveData`);
  }
});

test('a mirror miss falls through to Graph on every covered tool', async () => {
  const misses: Array<readonly [string, Record<string, unknown>]> = [
    ['list_sections', { containerType: 'notebook', containerId: 'nb-nope' }],
    ['list_pages', { sectionId: 'sec-nope' }],
    ['get_page_content', { pageId: 'p-nope' }],
  ];

  for (const [name, args] of misses) {
    const { body, graph } = await run(name, args);
    assert.equal(body['source'], 'graph', name);
    assert.ok(graph > 0, `${name} did not fall through`);
  }
});

test('a container the mirror has not synced yet is retried against Graph', async () => {
  // The realistic `_by_name` miss. resolveSection over the mirror raises a
  // NameLookupError listing the siblings the mirror happens to hold, which reads to a
  // model as "no such notebook". The retry is what stops that being the answer, and it
  // costs one Graph request only on a failure.
  for (const [name, args] of [
    ['find_page_by_name', { notebookName: '2027', sectionName: 'Brand new', pageTitle: 'x' }],
    ['list_pages_by_name', { notebookName: '2027', sectionName: 'Brand new' }],
  ] as const) {
    const { body, graph } = await run(name, args);
    assert.equal(body['source'], 'graph', name);
    assert.ok(graph > 0, `${name} did not retry against Graph`);
  }
});

test('a name that exists nowhere is still an error, not a silent empty answer', async () => {
  // The retry must not turn "this section does not exist" into something softer. Both
  // paths raise, and Graph's error is the one that reaches the caller.
  const all = tools({ total: 0 }, { total: 0 });
  await assert.rejects(
    () =>
      byName(all, 'list_pages_by_name').handle({
        notebookName: 'Nope',
        sectionName: 'Daily',
      }),
    /notebookName/,
  );
});

test('a Firestore outage answers from Graph rather than failing the call', async () => {
  // Strictly better than the behaviour before the mirror existed, which is the bar.
  for (const [name, args] of COVERED) {
    const { body, graph } = await run(name, args, { broken: true });
    assert.equal(body['source'], 'graph', name);
    assert.ok(graph > 0, name);
  }
});

test('with no mirror configured every tool is Graph-only, and says so', async () => {
  // MIRROR_READ_ENABLED=false. This is the rollback.
  for (const [name, args] of COVERED) {
    const { body, graph, mirror } = await run(name, args, { noMirror: true });
    assert.equal(body['source'], 'graph', name);
    assert.equal(mirror, 0, name);
    assert.ok(graph > 0, name);
  }
});

test('list_notebooks reports every notebook and which have their pages held', async () => {
  const { body } = await run('list_notebooks', {});
  assert.deepEqual(body['notebooks'], [
    { id: 'nb-2026', displayName: '2026', pagesMirrored: true },
    { id: 'nb-2025', displayName: '2025', pagesMirrored: false },
  ]);
});

test('list_notebooks from Graph omits pagesMirrored rather than claiming false', async () => {
  // Graph knows nothing about the mirror. A `pagesMirrored: false` on every notebook
  // would be a claim this path cannot make.
  const { body } = await run('list_notebooks', { useLiveData: true });
  const notebooks = body['notebooks'] as Record<string, unknown>[];
  assert.equal('pagesMirrored' in (notebooks[0] ?? {}), false);
});

test('an unscoped search from the mirror reports its scope, not a walk that never happened', async () => {
  // The Graph path reports sectionsSearched and stoppedEarly about a walk. There is no
  // walk here — but page content is held for one of two notebooks, so the scope is
  // narrower than the account and a model has to be told.
  const { body } = await run('search_pages', { query: 'held' });

  assert.equal(body['source'], 'mirror');
  assert.equal(body['notebooksSearched'], 1);
  assert.equal(body['notebooksInAccount'], 2);
  assert.equal(body['stoppedEarly'], undefined, 'no walk, so no walk bound to report');
  assert.match(String(body['note']), /1 of 2 notebooks/);
  assert.match(String(body['note']), /useLiveData/);
});

test('a search covering every notebook says so plainly', async () => {
  const { body } = await run('search_pages', { query: 'held', useLiveData: false });
  assert.equal(body['notebooksSearched'], 1);

  // And the Graph path keeps its own vocabulary untouched.
  const live = await run('search_pages', { query: 'held', useLiveData: true });
  assert.equal(live.body['notebooksSearched'], undefined);
  assert.notEqual(live.body['sectionsSearched'], undefined);
});

test('list_pages reports an exact moreAvailable from the mirror and the heuristic from Graph', async () => {
  const held = await run('list_pages', { sectionId: 'sec-daily', top: 1 });
  assert.equal(held.body['count'], 1);
  assert.equal(held.body['moreAvailable'], false, 'one page held, one asked for');

  // Graph returned exactly `top`, which is all it can ever do, so the heuristic stands.
  const live = await run('list_pages', { sectionId: 'sec-daily', top: 1, useLiveData: true });
  assert.equal(live.body['moreAvailable'], true);
});

test('get_page_content carries the source in the JSON and in the note', async () => {
  const { body } = await run('get_page_content', { pageId: 'p-held' });

  assert.equal(body['source'], 'mirror');
  assert.equal(body['mirroredAt'], '2026-08-19T12:30:00.000Z');
  assert.ok(String(body['html']).includes('held content'));
  // Said in prose too, so a model that ignores keys it does not recognise still learns
  // the answer may be minutes old.
  assert.match(String(body['note']), /local copy/);
  assert.match(String(body['note']), /useLiveData/);
});

test('a Graph-answered page carries no mirroredAt and no staleness note', async () => {
  const { body } = await run('get_page_content', { pageId: 'p-held', useLiveData: true });

  assert.equal(body['source'], 'graph');
  assert.equal(body['mirroredAt'], undefined);
  assert.equal(String(body['note']).includes('local copy'), false);
});

test('an empty mirror is a miss for the _by_name tools, not a NameLookupError', async () => {
  // resolveSection over an empty tree would raise "sectionName matched nothing" listing
  // no siblings, which reads to a model as "no such section". The retry against Graph is
  // what stops that being the answer.
  for (const [name, args] of [
    ['find_page_by_name', { notebookName: '2026', sectionName: 'Daily', pageTitle: 'x' }],
    ['list_pages_by_name', { notebookName: '2026', sectionName: 'Daily' }],
  ] as const) {
    const { body, graph } = await run(name, args, { empty: true });
    assert.equal(body['source'], 'graph', name);
    assert.ok(graph > 0, name);
  }
});

test('useLiveData must be a boolean', async () => {
  const all = tools({ total: 0 }, { total: 0 });
  await assert.rejects(
    () => byName(all, 'list_pages').handle({ sectionId: 'sec-daily', useLiveData: 'true' }),
    /useLiveData/,
  );
});

// ---------------------------------------------------------------------------
// Write invalidation
//
// A write goes to Graph and then tells the mirror its copy is superseded. Without it, a
// `get_page_content` straight after an `append_to_page` serves pre-write content with
// nothing saying so — the write moved the page's timestamp, but the mirror does not know
// that until the next sync run.
// ---------------------------------------------------------------------------

interface SyncCalls {
  resynced: { pageId: string; hint: { title?: string; sectionId?: string } }[];
  staled: string[];
}

function fakeWriteSync(
  calls: SyncCalls,
  options: { resyncFails?: boolean; staleFails?: boolean } = {},
): MirrorWriteSync {
  return {
    resyncPage: (pageId, hint) => {
      calls.resynced.push({ pageId, hint });
      return options.resyncFails === true
        ? Promise.reject(new Error('graph down'))
        : Promise.resolve();
    },
    markPageStale: (pageId) => {
      calls.staled.push(pageId);
      return options.staleFails === true
        ? Promise.reject(new Error('firestore down'))
        : Promise.resolve();
    },
  };
}

test('every write resyncs its page, including create_page', async () => {
  // The point of resyncing rather than marking stale: a get_page_content straight after
  // an append answers from the mirror with the appended text, instead of falling through
  // to Graph until the next scheduled run.
  const calls: SyncCalls = { resynced: [], staled: [] };
  const writeTools = createWriteTools(
    fakeWriteClient(),
    fakeLayoutReader(),
    fakeWriteLookup(),
    fakeWriteSync(calls),
  );

  await byName(writeTools, 'append_to_page').handle({
    pageId: 'p-held',
    htmlFragment: '<p>added</p>',
  });
  await byName(writeTools, 'update_page_title').handle({
    pageId: 'p-held',
    newTitle: 'Renamed',
  });
  await byName(writeTools, 'create_page').handle({
    sectionId: 'sec-daily',
    title: 'Fresh',
    htmlFragment: '<p>new</p>',
  });

  assert.deepEqual(calls.resynced.map((r) => r.pageId), ['p-held', 'p-held', 'p-new']);
  assert.deepEqual(calls.staled, [], 'the fallback is not used when the resync worked');
});

test('the hints carry what a metadata read cannot be trusted for', async () => {
  // Measured 2026-08-19: GET /pages/{id}?$select=title returned "" for pages created
  // seconds earlier. So the title travels from the caller, which just set it.
  const calls: SyncCalls = { resynced: [], staled: [] };
  const writeTools = createWriteTools(
    fakeWriteClient(),
    fakeLayoutReader(),
    fakeWriteLookup(),
    fakeWriteSync(calls),
  );

  await byName(writeTools, 'append_to_page').handle({
    pageId: 'p-held',
    htmlFragment: '<p>added</p>',
  });
  await byName(writeTools, 'update_page_title').handle({
    pageId: 'p-held',
    newTitle: 'Renamed',
  });
  await byName(writeTools, 'create_page').handle({
    sectionId: 'sec-daily',
    title: 'Fresh',
    htmlFragment: '<p>new</p>',
  });

  // An append cannot change a title, so it sends none and the stored one stands.
  assert.deepEqual(calls.resynced[0]?.hint, {});
  assert.equal(calls.resynced[1]?.hint.title, 'Renamed');
  // A created page has no stored placement to read a section from, so it must be told.
  assert.equal(calls.resynced[2]?.hint.title, 'Fresh');
  assert.equal(calls.resynced[2]?.hint.sectionId, 'sec-daily');
});

test('a failed resync falls back to marking the page stale', async () => {
  const calls: SyncCalls = { resynced: [], staled: [] };
  const writeTools = createWriteTools(
    fakeWriteClient(),
    fakeLayoutReader(),
    fakeWriteLookup(),
    fakeWriteSync(calls, { resyncFails: true }),
  );

  const result = await byName(writeTools, 'append_to_page').handle({
    pageId: 'p-held',
    htmlFragment: '<p>added</p>',
  });

  // Stale is correct, just slower: the next read is a miss and goes to Graph.
  assert.deepEqual(calls.staled, ['p-held']);
  assert.equal(payload(result)['appended'], true);
});

test('neither a failed resync nor a failed fallback fails the write', async () => {
  // The write is the thing that mattered and it has already happened. Reporting an error
  // would send the caller to retry a change that is already made. It is self-healing:
  // the write moved the page's lastModifiedDateTime, so the next sync repairs it.
  const calls: SyncCalls = { resynced: [], staled: [] };
  const writeTools = createWriteTools(
    fakeWriteClient(),
    fakeLayoutReader(),
    fakeWriteLookup(),
    fakeWriteSync(calls, { resyncFails: true, staleFails: true }),
  );

  const result = await byName(writeTools, 'append_to_page').handle({
    pageId: 'p-held',
    htmlFragment: '<p>added</p>',
  });

  assert.notEqual(result.isError, true);
  assert.equal(payload(result)['appended'], true);
});

test('with no mirror configured the write tools behave exactly as before', async () => {
  const writeTools = createWriteTools(fakeWriteClient(), fakeLayoutReader(), fakeWriteLookup());

  const result = await byName(writeTools, 'append_to_page').handle({
    pageId: 'p-held',
    htmlFragment: '<p>added</p>',
  });

  assert.equal(payload(result)['appended'], true);
});
