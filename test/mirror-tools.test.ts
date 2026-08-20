// The seven covered tools with a mirror supplied, driven through their own `handle`.
//
// Separate from test/structure-tools.test.ts and test/page-tools.test.ts on purpose:
// those two build their tools with no mirror argument, and every assertion in them still
// holds unchanged. That is the property worth protecting — an absent mirror means "always
// Graph", which is what makes `MIRROR_READ_ENABLED=false` a complete rollback rather than
// a code path with its own bugs.
//
// Both fakes count their calls, because most of what matters here is what does *not*
// happen: a mirror hit must not touch Graph, a mirror failure must still answer, and an
// absent mirror must not attempt a refresh.
//
// The third fake is the refresh. Every read is refresh-then-read, and whether the refresh
// finished is what decides between the two labels a tool may report — `onenote` for an
// answer that equals what OneNote holds, `mirror` for a copy that may be behind. The
// harness below drives it to either answer, because the label is the claim the calling
// model acts on and a wrong one is a confident wrong answer.

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
import type { ReadFreshness, ReadSync } from '../src/read-sync.ts';
import type { ToolDefinition } from '../src/mcp-tools.ts';
import { createPageTools, type PageContentClient } from '../src/page-tools.ts';
import { createStructureTools, type StructureClient } from '../src/structure-tools.ts';
import type { ResyncOutcome } from '../src/mirror-sync.ts';
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
  options: { broken?: boolean; empty?: boolean; inactive?: string[] } = {},
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
    // `inactive` names the notebooks the operator froze; absent means the selection
    // document names no active set at all, which is every mirrored notebook active.
    getSelection: () =>
      count({
        notebookIds: MIRROR_NOTEBOOKS.filter((n) => n.mirrored).map((n) => n.id),
        activeNotebookIds:
          options.inactive === undefined
            ? null
            : MIRROR_NOTEBOOKS.filter(
                (n) => n.mirrored && !options.inactive?.includes(n.id),
              ).map((n) => n.id),
      }),
    listNotebooks: () => count(options.empty === true ? [] : MIRROR_NOTEBOOKS),
    listSectionsUnder: (parentId) =>
      count(MIRROR_SECTIONS.filter((s) => s.parentId === parentId)),
    listSectionGroupsUnder: () => count(groups),
    listAllSections: () => count(options.empty === true ? [] : MIRROR_SECTIONS),
    listHeldSections: () => count(MIRROR_SECTIONS.filter((s) => (s.pendingWrites ?? 0) > 0)),
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

function fakeWriteClient(order: string[] = []): PageWriteClient {
  return {
    appendToPage: () => {
      order.push('graph-write');
      return Promise.resolve();
    },
    updatePageTitle: () => {
      order.push('graph-write');
      return Promise.resolve();
    },
    createPage: () => {
      order.push('graph-write');
      return Promise.resolve({ id: 'p-new', title: 'Fresh', webUrl: null, clientUrl: null });
    },
  };
}

/** A write client whose every call fails, for the release-on-failure path. */
function failingWriteClient(): PageWriteClient {
  const fail = (): Promise<never> => Promise.reject(new Error('graph rejected the write'));
  return { appendToPage: fail, updatePageTitle: fail, createPage: fail };
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

/** How the refresh behaves for one run. `current` is the ordinary case. */
interface RunOptions {
  broken?: boolean;
  empty?: boolean;
  noMirror?: boolean;
  freshness?: ReadFreshness;
  /** Mirrored notebooks the operator marked inactive. */
  inactive?: string[];
}

interface SyncCounter {
  refreshes: number;
}

function fakeReadSync(counter: SyncCounter, freshness: ReadFreshness): ReadSync {
  return {
    refresh: () => {
      counter.refreshes += 1;
      return Promise.resolve(freshness);
    },
  };
}

function tools(
  graphCalls: GraphCalls,
  mirrorCalls: MirrorCalls,
  syncCounter: SyncCounter,
  options: RunOptions = {},
): ToolDefinition[] {
  const mirror = options.noMirror === true ? undefined : fakeMirror(mirrorCalls, options);
  const sync = fakeReadSync(syncCounter, options.freshness ?? 'current');
  return [
    ...createStructureTools(fakeGraph(graphCalls), {}, mirror, sync),
    ...createPageTools(fakeContent(graphCalls), undefined, mirror, sync),
  ];
}

async function run(
  name: string,
  args: Record<string, unknown>,
  options: RunOptions = {},
): Promise<{
  body: Record<string, unknown>;
  graph: number;
  mirror: number;
  refreshes: number;
}> {
  const graphCalls: GraphCalls = { total: 0 };
  const mirrorCalls: MirrorCalls = { total: 0 };
  const syncCounter: SyncCounter = { refreshes: 0 };
  const body = payload(
    await byName(tools(graphCalls, mirrorCalls, syncCounter, options), name).handle(args),
  );
  return {
    body,
    graph: graphCalls.total,
    mirror: mirrorCalls.total,
    refreshes: syncCounter.refreshes,
  };
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

test('no covered tool takes a live-data argument, because every read refreshes first', () => {
  // The argument existed to force a read past a mirror that might be behind. A read that
  // refreshes the mirror first and reports whether the refresh finished answers the same
  // question without a schema field the model has to reason about.
  for (const tool of tools({ total: 0 }, { total: 0 }, { refreshes: 0 })) {
    const properties = (tool.inputSchema.properties ?? {}) as Record<string, unknown>;
    assert.equal('useLiveData' in properties, false, tool.name);
  }
});

test('every covered tool refreshes the mirror exactly once before it reads', async () => {
  for (const [name, args] of COVERED) {
    const { refreshes } = await run(name, args);
    assert.equal(refreshes, 1, name);
  }
});

test('a refreshed mirror answers every covered tool as OneNote, untouched by Graph', async () => {
  for (const [name, args] of COVERED) {
    const { body, graph } = await run(name, args);
    assert.equal(body['source'], 'onenote', name);
    assert.equal(graph, 0, `${name} reached Graph on a mirror hit`);
  }
});

test('a refresh that did not finish downgrades the label to mirror', async () => {
  // The budget ran out, the lease was held, the tree read failed. The data is the same
  // data; what changed is that nothing checked it against OneNote on this call, and
  // saying `onenote` anyway would be a claim the server cannot make.
  for (const [name, args] of COVERED) {
    if (name === 'get_page_content') continue;
    const { body, graph } = await run(name, args, { freshness: 'behind' });
    assert.equal(body['source'], 'mirror', name);
    assert.equal(graph, 0, `${name} fell through instead of labelling the answer`);
  }
});

test('get_page_content stays OneNote through a refresh that did not finish', async () => {
  // The one exception, and the reason is per-page rather than per-account: a page
  // document carries its own contentState, and every write marks its page stale before it
  // touches OneNote. A hit is a page nothing has superseded.
  const { body, graph } = await run('get_page_content', { pageId: 'p-held' }, { freshness: 'behind' });
  assert.equal(body['source'], 'onenote');
  assert.equal(graph, 0);
});

test('a mirror miss falls through to Graph on every covered tool', async () => {
  const misses: Array<readonly [string, Record<string, unknown>]> = [
    ['list_sections', { containerType: 'notebook', containerId: 'nb-nope' }],
    ['list_pages', { sectionId: 'sec-nope' }],
    ['get_page_content', { pageId: 'p-nope' }],
  ];

  for (const [name, args] of misses) {
    const { body, graph } = await run(name, args);
    assert.equal(body['source'], 'onenote', name);
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
    assert.equal(body['source'], 'onenote', name);
    assert.ok(graph > 0, `${name} did not retry against Graph`);
  }
});

test('a name that exists nowhere is still an error, not a silent empty answer', async () => {
  // The retry must not turn "this section does not exist" into something softer. Both
  // paths raise, and Graph's error is the one that reaches the caller.
  const all = tools({ total: 0 }, { total: 0 }, { refreshes: 0 });
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
    assert.equal(body['source'], 'onenote', name);
    assert.ok(graph > 0, name);
  }
});

test('with no mirror configured every tool is Graph-only, and refreshes nothing', async () => {
  // MIRROR_READ_ENABLED=false. This is the rollback, and it includes the promise that no
  // tool call spends a Graph request on a sync whose result it has no use for.
  for (const [name, args] of COVERED) {
    const { body, graph, mirror, refreshes } = await run(name, args, { noMirror: true });
    assert.equal(body['source'], 'onenote', name);
    assert.equal(mirror, 0, name);
    assert.equal(refreshes, 0, `${name} refreshed a mirror it does not read`);
    assert.ok(graph > 0, name);
  }
});

test('list_notebooks reports every notebook and which have their pages held', async () => {
  const { body } = await run('list_notebooks', {});
  assert.deepEqual(body['notebooks'], [
    { id: 'nb-2026', displayName: '2026', pagesMirrored: true, pagesActive: true },
    { id: 'nb-2025', displayName: '2025', pagesMirrored: false, pagesActive: false },
  ]);
});

test('pagesActive is false for a mirrored notebook the operator froze', async () => {
  // pagesMirrored says a local copy exists; pagesActive says whether anything keeps it
  // current. A model choosing where to look needs both, and they are not the same fact.
  const { body } = await run('list_notebooks', {}, { inactive: ['nb-2026'] });
  assert.deepEqual(body['notebooks'], [
    { id: 'nb-2026', displayName: '2026', pagesMirrored: true, pagesActive: false },
    { id: 'nb-2025', displayName: '2025', pagesMirrored: false, pagesActive: false },
  ]);
});

test('list_notebooks from Graph omits pagesMirrored rather than claiming false', async () => {
  // Graph knows nothing about the mirror. A `pagesMirrored: false` on every notebook
  // would be a claim this path cannot make. An empty mirror is the miss that gets here.
  const { body } = await run('list_notebooks', {}, { empty: true });
  assert.equal(body['source'], 'onenote');
  const notebooks = body['notebooks'] as Record<string, unknown>[];
  assert.equal('pagesMirrored' in (notebooks[0] ?? {}), false);
});

test('an unscoped search from the mirror reports its scope, not a walk that never happened', async () => {
  // The Graph path reports sectionsSearched and stoppedEarly about a walk. There is no
  // walk here — but page content is held for one of two notebooks, so the scope is
  // narrower than the account and a model has to be told.
  const { body } = await run('search_pages', { query: 'held' });

  assert.equal(body['source'], 'onenote');
  assert.equal(body['notebooksSearched'], 1);
  assert.equal(body['notebooksInAccount'], 2);
  assert.equal(body['stoppedEarly'], undefined, 'no walk, so no walk bound to report');
  assert.match(String(body['note']), /1 of 2 notebooks/);
  // The escape hatch is naming a section, not an argument: a scoped search misses in the
  // mirror and walks that one section in OneNote.
  assert.match(String(body['note']), /sectionId/);
});

// ---------------------------------------------------------------------------
// Frozen notebooks, per tool
//
// The point of these is which tools ask and which do not. `list_notebooks` and
// `list_sections` answer about *structure*, which every sync run rewrites whatever the
// active set says, so weakening their label would report a staleness that does not exist.
// ---------------------------------------------------------------------------

const FROZEN = { inactive: ['nb-2026'] };

test('every page-reading tool reports best-available from a frozen notebook', async () => {
  const cases: Array<[string, Record<string, unknown>]> = [
    ['list_pages', { sectionId: 'sec-daily' }],
    ['search_pages', { query: 'held', sectionId: 'sec-daily' }],
    ['search_pages', { query: 'held' }],
    ['get_page_content', { pageId: 'p-held' }],
    ['list_pages_by_name', { notebookName: '2026', sectionName: 'Daily' }],
    ['find_page_by_name', { notebookName: '2026', sectionName: 'Daily', pageTitle: 'Held Monday' }],
  ];

  for (const [name, args] of cases) {
    const { body, graph } = await run(name, args, FROZEN);
    assert.equal(body['source'], 'best-available', name);
    assert.equal(graph, 0, `${name} still answered from the mirror`);
  }
});

test('the structure tools are unaffected by activity', async () => {
  // Structure is stored for the whole account on every run, active or not — the tree read
  // returns it all for the same one request. Reporting best-available here would describe
  // a staleness these two do not have.
  for (const [name, args] of [
    ['list_notebooks', {}],
    ['list_sections', { containerType: 'notebook', containerId: 'nb-2026' }],
  ] as const) {
    const { body } = await run(name, args, FROZEN);
    assert.equal(body['source'], 'onenote', name);
  }
});

test('a frozen notebook stays best-available when the refresh fails too', async () => {
  // The row that looks wrong and is not: that refresh was never going to check this
  // notebook, so reporting `mirror` would describe a failure that could not have changed
  // the answer.
  const { body } = await run('list_pages', { sectionId: 'sec-daily' }, {
    ...FROZEN,
    freshness: 'behind',
  });
  assert.equal(body['source'], 'best-available');
});

test('an unscoped search says how many notebooks are frozen', async () => {
  // best-available says a skip happened; the count says how large it was.
  const { body } = await run('search_pages', { query: 'held' }, FROZEN);
  assert.equal(body['inactiveNotebooks'], 1);
  assert.equal(body['notebooksSearched'], 1);

  const active = await run('search_pages', { query: 'held' });
  assert.equal(active.body['inactiveNotebooks'], 0);
  assert.equal(active.body['source'], 'onenote');
});

test('get_page_content warns in prose that a frozen notebook is not re-checked', async () => {
  const { body } = await run('get_page_content', { pageId: 'p-held' }, FROZEN);
  assert.equal(body['source'], 'best-available');
  assert.match(String(body['note']), /not re-checked/);
});

test('the Graph search path keeps its own vocabulary', async () => {
  const { body } = await run('search_pages', { query: 'held' }, { empty: true });
  assert.equal(body['notebooksSearched'], undefined);
  assert.notEqual(body['sectionsSearched'], undefined);
});

test('list_pages reports an exact moreAvailable from the mirror and the heuristic from Graph', async () => {
  const held = await run('list_pages', { sectionId: 'sec-daily', top: 1 });
  assert.equal(held.body['count'], 1);
  assert.equal(held.body['moreAvailable'], false, 'one page held, one asked for');

  // Graph returned exactly `top`, which is all it can ever do, so the heuristic stands.
  const live = await run('list_pages', { sectionId: 'sec-nope', top: 1 });
  assert.equal(live.body['source'], 'onenote');
  assert.equal(live.body['moreAvailable'], true);
});

test('get_page_content carries the source and the sync time in the JSON and in the note', async () => {
  const { body } = await run('get_page_content', { pageId: 'p-held' });

  assert.equal(body['source'], 'onenote');
  assert.equal(body['mirroredAt'], '2026-08-19T12:30:00.000Z');
  assert.ok(String(body['html']).includes('held content'));
  // Said in prose too, so a model that ignores keys it does not recognise still learns
  // the page came from a stored copy and when it was taken.
  assert.match(String(body['note']), /local copy/);
});

test('a Graph-answered page carries no mirroredAt and no local-copy note', async () => {
  const { body } = await run('get_page_content', { pageId: 'p-nope' });

  assert.equal(body['source'], 'onenote');
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
    assert.equal(body['source'], 'onenote', name);
    assert.ok(graph > 0, name);
  }
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
  held: string[];
  released: string[];
}

function fakeWriteSync(
  calls: SyncCalls,
  options: {
    resyncFails?: boolean;
    staleFails?: boolean;
    holdFails?: boolean;
    outcome?: ResyncOutcome;
    order?: string[];
    sectionOf?: Record<string, string>;
  } = {},
): MirrorWriteSync {
  return {
    resyncPage: (pageId, hint) => {
      options.order?.push('resync');
      calls.resynced.push({ pageId, hint });
      return options.resyncFails === true
        ? Promise.reject(new Error('graph down'))
        : Promise.resolve(options.outcome ?? 'updated');
    },
    markPageStale: (pageId) => {
      options.order?.push('mark-stale');
      calls.staled.push(pageId);
      return options.staleFails === true
        ? Promise.reject(new Error('firestore down'))
        : Promise.resolve();
    },
    sectionOfPage: (pageId) =>
      Promise.resolve(
        options.sectionOf?.[pageId] ??
          MIRROR_PAGES.find((page) => page.id === pageId)?.sectionId ??
          null,
      ),
    holdSectionListing: (sectionId) => {
      options.order?.push('hold');
      calls.held.push(sectionId);
      return options.holdFails === true
        ? Promise.reject(new Error('firestore down'))
        : Promise.resolve();
    },
    releaseSectionListing: (sectionId) => {
      options.order?.push('release');
      calls.released.push(sectionId);
      return Promise.resolve();
    },
  };
}

test('every write resyncs its page, including create_page', async () => {
  // The point of resyncing rather than marking stale: a get_page_content straight after
  // an append answers from the mirror with the appended text, instead of falling through
  // to Graph until the next scheduled run.
  const calls: SyncCalls = { resynced: [], staled: [], held: [], released: [] };
  const order: string[] = [];
  const writeTools = createWriteTools(
    fakeWriteClient(order),
    fakeLayoutReader(),
    fakeWriteLookup(),
    fakeWriteSync(calls, { order }),
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
  // The whole sequence, rather than counts. Every one of the three invalidates first,
  // then writes to OneNote, then resyncs — and no resync is followed by a fallback,
  // because all three succeeded. What each one invalidates differs: an append marks the
  // page's content stale, a rename does that *and* holds the section's page listing
  // because the title is what a listing shows, and a create holds only the listing
  // because it has no page document to mark.
  assert.deepEqual(order, [
    'mark-stale', 'graph-write', 'resync',
    'mark-stale', 'hold', 'graph-write', 'resync', 'release',
    'hold', 'graph-write', 'resync', 'release',
  ]);
});

test('the hints carry what a metadata read cannot be trusted for', async () => {
  // Measured 2026-08-19: GET /pages/{id}?$select=title returned "" for pages created
  // seconds earlier. So the title travels from the caller, which just set it.
  const calls: SyncCalls = { resynced: [], staled: [], held: [], released: [] };
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
  const calls: SyncCalls = { resynced: [], staled: [], held: [], released: [] };
  const order: string[] = [];
  const writeTools = createWriteTools(
    fakeWriteClient(order),
    fakeLayoutReader(),
    fakeWriteLookup(),
    fakeWriteSync(calls, { resyncFails: true, order }),
  );

  const result = await byName(writeTools, 'append_to_page').handle({
    pageId: 'p-held',
    htmlFragment: '<p>added</p>',
  });

  // Stale is correct, just slower: the next read is a miss and goes to Graph.
  assert.deepEqual(order, ['mark-stale', 'graph-write', 'resync', 'mark-stale']);
  assert.equal(payload(result)['appended'], true);
});

test('neither a failed resync nor a failed fallback fails the write', async () => {
  // The write is the thing that mattered and it has already happened. Reporting an error
  // would send the caller to retry a change that is already made. It is self-healing:
  // the write moved the page's lastModifiedDateTime, so the next sync repairs it.
  const calls: SyncCalls = { resynced: [], staled: [], held: [], released: [] };
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

test('an append that resynced to "unchanged" is marked stale, not left as current', async () => {
  // An append always changes content, so a resync finding nothing to write means the
  // read did not see the write. Leaving the page `present` would serve pre-write content
  // as current, with nothing saying so. Marking it stale sends the next read to Graph,
  // which cannot be wrong.
  const calls: SyncCalls = { resynced: [], staled: [], held: [], released: [] };
  const order: string[] = [];
  const writeTools = createWriteTools(
    fakeWriteClient(order),
    fakeLayoutReader(),
    fakeWriteLookup(),
    fakeWriteSync(calls, { outcome: 'unchanged', order }),
  );

  await byName(writeTools, 'append_to_page').handle({
    pageId: 'p-held',
    htmlFragment: '<p>added</p>',
  });
  await byName(writeTools, 'create_page').handle({
    sectionId: 'sec-daily',
    title: 'Fresh',
    htmlFragment: '<p>new</p>',
  });

  // Each write's resync reported nothing to store, so each is followed by a fallback.
  assert.deepEqual(order, [
    'mark-stale', 'graph-write', 'resync', 'mark-stale',
    'hold', 'graph-write', 'resync', 'mark-stale', 'release',
  ]);
});

test('a rename that resynced to "unchanged" is left alone', async () => {
  // A rename changes no content by design, so `unchanged` is not evidence of a lost
  // write here — and the title comparison in writePageFromRaw is what makes a real
  // rename come back `updated` anyway.
  const calls: SyncCalls = { resynced: [], staled: [], held: [], released: [] };
  const order: string[] = [];
  const writeTools = createWriteTools(
    fakeWriteClient(order),
    fakeLayoutReader(),
    fakeWriteLookup(),
    fakeWriteSync(calls, { outcome: 'unchanged', order }),
  );

  await byName(writeTools, 'update_page_title').handle({
    pageId: 'p-held',
    newTitle: 'Renamed',
  });

  assert.deepEqual(order, ['mark-stale', 'hold', 'graph-write', 'resync', 'release']);
});

test('a page outside the mirrored set needs no stale marker', async () => {
  // `not-mirrored` means there is no copy to be wrong.
  const calls: SyncCalls = { resynced: [], staled: [], held: [], released: [] };
  const order: string[] = [];
  const writeTools = createWriteTools(
    fakeWriteClient(order),
    fakeLayoutReader(),
    fakeWriteLookup(),
    fakeWriteSync(calls, { outcome: 'not-mirrored', order }),
  );

  await byName(writeTools, 'append_to_page').handle({
    pageId: 'p-elsewhere',
    htmlFragment: '<p>added</p>',
  });

  // The pre-mark is issued blind -- the tool does not know whether a copy exists, and
  // MirrorStore.markPageStale no-ops on NOT_FOUND rather than creating a stub. What must
  // not happen is a second one after the resync reported there was nothing to hold.
  assert.deepEqual(order, ['mark-stale', 'graph-write', 'resync']);
});

test('the page is marked stale BEFORE the Graph write, not only after it', async () => {
  // The window that ordering closes: between the PATCH succeeding and the resync
  // completing, OneNote and the mirror disagree. If the process merely errors, resync's
  // catch marks the page stale. If the process *stops* -- Cloud Run cutting the request,
  // or the instance being reclaimed -- the resync is sitting in the request gate's queue,
  // the queue goes with the instance, and no catch runs because nothing threw. The mirror
  // would keep serving pre-write content as `present` until the next scheduled sync.
  //
  // Marking first makes the window pessimistic: a death anywhere in it leaves a miss, and
  // a miss goes to Graph.
  for (const [name, args] of [
    ['append_to_page', { pageId: 'p-held', htmlFragment: '<p>added</p>' }],
    ['update_page_title', { pageId: 'p-held', newTitle: 'Renamed' }],
  ] as const) {
    const order: string[] = [];
    const calls: SyncCalls = { resynced: [], staled: [], held: [], released: [] };
    const writeTools = createWriteTools(
      fakeWriteClient(order),
      fakeLayoutReader(),
      fakeWriteLookup(),
      fakeWriteSync(calls, { order }),
    );

    await byName(writeTools, name).handle(args);

    assert.equal(order.indexOf('mark-stale') < order.indexOf('graph-write'), true, name);
    assert.equal(order.indexOf('graph-write') < order.indexOf('resync'), true, name);
  }
});

test('create_page marks no page stale, and holds its section listing instead', async () => {
  // A page the mirror has never seen is already a miss, so a stale marker for it would be
  // a Firestore write for nothing — and there is no document to write it to. What a
  // create *does* make wrong is the section's page listing: `list_pages`,
  // `list_pages_by_name`, `find_page_by_name` and `search_pages` all answer from stored
  // page documents, so between the create and its resync the mirror reports a section
  // that does not contain the page. To a model that reads as "the page was not created".
  const order: string[] = [];
  const calls: SyncCalls = { resynced: [], staled: [], held: [], released: [] };
  const writeTools = createWriteTools(
    fakeWriteClient(order),
    fakeLayoutReader(),
    fakeWriteLookup(),
    fakeWriteSync(calls, { order }),
  );

  await byName(writeTools, 'create_page').handle({
    sectionId: 'sec-daily',
    title: 'Fresh',
    htmlFragment: '<p>new</p>',
  });

  assert.deepEqual(order, ['hold', 'graph-write', 'resync', 'release']);
  assert.deepEqual(calls.staled, []);
  assert.deepEqual(calls.held, ['sec-daily']);
  assert.deepEqual(calls.released, ['sec-daily']);
});

test('a Graph write that failed still releases the listing hold', async () => {
  // The hold is released in a `finally`. A write that never happened left the listing
  // correct, and a hold nothing lowers costs every later listing for that section a Graph
  // request until it expires on age.
  for (const [name, args] of [
    ['create_page', { sectionId: 'sec-daily', title: 'Fresh', htmlFragment: '<p>new</p>' }],
    ['update_page_title', { pageId: 'p-held', newTitle: 'Renamed' }],
  ] as const) {
    const calls: SyncCalls = { resynced: [], staled: [], held: [], released: [] };
    const writeTools = createWriteTools(
      failingWriteClient(),
      fakeLayoutReader(),
      fakeWriteLookup(),
      fakeWriteSync(calls, {}),
    );

    await assert.rejects(() => byName(writeTools, name).handle(args));

    assert.deepEqual(calls.held, ['sec-daily'], name);
    assert.deepEqual(calls.released, ['sec-daily'], name);
    assert.deepEqual(calls.resynced, [], name);
  }
});

test('a failed listing hold does not stop the write', async () => {
  // Same rule as the stale marker: it narrows an existing window rather than being a
  // precondition for writing.
  const calls: SyncCalls = { resynced: [], staled: [], held: [], released: [] };
  const writeTools = createWriteTools(
    fakeWriteClient(),
    fakeLayoutReader(),
    fakeWriteLookup(),
    fakeWriteSync(calls, { holdFails: true }),
  );

  const result = await byName(writeTools, 'create_page').handle({
    sectionId: 'sec-daily',
    title: 'Fresh',
    htmlFragment: '<p>new</p>',
  });

  assert.equal(payload(result)['pageId'], 'p-new');
  assert.deepEqual(calls.resynced.map((r) => r.pageId), ['p-new']);
});

test('a failed pre-mark does not stop the write', async () => {
  // It narrows an existing window; it is not a precondition for writing.
  const calls: SyncCalls = { resynced: [], staled: [], held: [], released: [] };
  const writeTools = createWriteTools(
    fakeWriteClient(),
    fakeLayoutReader(),
    fakeWriteLookup(),
    fakeWriteSync(calls, { staleFails: true }),
  );

  const result = await byName(writeTools, 'append_to_page').handle({
    pageId: 'p-held',
    htmlFragment: '<p>added</p>',
  });

  assert.equal(payload(result)['appended'], true);
});
