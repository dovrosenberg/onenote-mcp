// The sync algorithm, driven through fakes that count their calls.
//
// No `fetch` anywhere in this file. Every URL the sync builds is already asserted in
// test/graph-structure.test.ts against an exact-URL fake, so routing them again here
// would test the same thing twice and make each behavioural assertion pay for a routing
// table. What this file asserts is the algorithm: which sections get visited, when a
// watermark moves, what a failure costs, and what a budget stops.
//
// Most of the properties worth having are about what does *not* happen, which is why the
// fakes count calls rather than just answering. In order of how much damage the absence
// would do:
//
// 1. **A failed sweep enumeration deletes nothing.** An auth failure or a 500 would
//    otherwise empty the mirror one section at a time, and nothing about a deleted page
//    is recoverable from here.
// 2. **A section whose page listing failed keeps its old watermark**, so the next run
//    retries it rather than skipping every page it never reached.
// 3. **A budget-exhausted run keeps the advances it earned**, so a five-hour backfill
//    makes progress across slices instead of restarting.
// 4. **An unchanged content hash writes nothing and renders no ink**, which is what makes
//    the hour of watermark overlap nearly free.
//
// What no test here covers is whether Graph's timestamps behave as the algorithm assumes.
// That a page create, edit and delete each move the parent section's
// `lastModifiedDateTime` — and that nothing else does — was measured on 2026-08-19 and is
// recorded in api-overview.md; it is an assumption here, not a fact this file can check.

import test from 'node:test';
import assert from 'node:assert/strict';

import { GraphRequestError } from '../src/graph-structure.ts';
import type {
  ContainerChildren,
  ExpandedNotebook,
  PageSummary,
} from '../src/graph-structure.ts';
import { setEventSink } from '../src/logging.ts';
import {
  initialSyncState,
  type MirrorPage,
  type MirrorPageContent,
  type MirrorSection,
  type MirrorSectionGroup,
  type MirrorSyncState,
  type MirrorTombstone,
  type NotebookSelection,
  type SyncMode,
} from '../src/mirror-schema.ts';
import {
  buildStructure,
  pickCandidates,
  runFullSweep,
  runIncremental,
  runSweep,
  runSweepAll,
  resyncPage,
  splitByActivity,
  activeSelectionHashOf,
  structureHashOf,
  type ResyncDeps,
  type SyncDeps,
  type SyncStore,
} from '../src/mirror-sync.ts';
import type { RawPageContent } from '../src/page-content.ts';

setEventSink(() => {});

const NB = 'nb-1';
const NOW = Date.parse('2026-08-19T12:00:00.000Z');
const NOW_ISO = new Date(NOW).toISOString();

function graphError(status: number): GraphRequestError {
  return new GraphRequestError('https://graph.microsoft.com/v1.0/x', status, 'Error', '{}');
}

function rawHtml(html: string): RawPageContent {
  return { raw: html, contentType: 'text/html', parts: [] };
}

function section(overrides: Partial<MirrorSection> = {}): MirrorSection {
  return {
    id: 'sec-1',
    displayName: 'Daily',
    notebookId: NB,
    parentId: NB,
    parentKind: 'notebook',
    path: '2026 / Daily',
    mirrored: true,
    graphLastModifiedDateTime: '2026-08-19T11:00:00Z',
    pagesSyncedThrough: '2026-08-19T10:00:00Z',
    pageCount: 0,
    ...overrides,
  };
}

function summary(id: string, modified = '2026-08-19T11:30:00Z'): PageSummary {
  return { id, title: `Page ${id}`, lastModifiedDateTime: modified };
}

// ---------------------------------------------------------------------------
// Fakes
// ---------------------------------------------------------------------------

interface GraphScript {
  tree?: ExpandedNotebook[] | (() => never);
  changed?: Record<string, PageSummary[] | (() => never)>;
  ids?: Record<string, string[] | (() => never)>;
  children?: ContainerChildren;
}

interface GraphCalls {
  tree: number;
  changedSince: string[];
  unfiltered: string[];
  pageIds: string[];
  children: string[];
}

function fakeGraph(script: GraphScript): { graph: SyncDeps['graph']; calls: GraphCalls } {
  const calls: GraphCalls = { tree: 0, changedSince: [], unfiltered: [], pageIds: [], children: [] };

  // A scripted value, or a thunk that throws to simulate a Graph failure.
  const resolve = <T>(value: T | (() => never) | undefined, fallback: T): T => {
    if (typeof value === 'function') (value as () => never)();
    return (value as T | undefined) ?? fallback;
  };

  return {
    calls,
    graph: {
      getExpandedTree: () => {
        calls.tree += 1;
        return Promise.resolve(resolve(script.tree, []));
      },
      listContainerChildren: (_kind, id) => {
        calls.children.push(id);
        return Promise.resolve(script.children ?? { sections: [], sectionGroups: [] });
      },
      listPagesChangedSince: (sectionId) => {
        calls.changedSince.push(sectionId);
        return Promise.resolve(resolve(script.changed?.[sectionId], []));
      },
      listPagesInSection: (sectionId) => {
        calls.unfiltered.push(sectionId);
        return Promise.resolve(resolve(script.changed?.[sectionId], []));
      },
      listPageIds: (sectionId) => {
        calls.pageIds.push(sectionId);
        return Promise.resolve(resolve(script.ids?.[sectionId], []));
      },
    },
  };
}

/**
 * A selection, with "every selected notebook is active" as the default.
 *
 * `activeNotebookIds: null` is what a document written before this feature existed reads
 * as, so it is what every test that is not about activity should use.
 */
function sel(notebookIds: string[], activeNotebookIds: string[] | null = null): NotebookSelection {
  return { notebookIds, activeNotebookIds };
}

interface StoreState {
  selection: NotebookSelection;
  state: MirrorSyncState;
  sections: MirrorSection[];
  groups: MirrorSectionGroup[];
  pages: Map<string, MirrorPage>;
  pageIdsBySection: Map<string, string[]>;
}

interface StoreCalls {
  watermarks: { sectionId: string; watermark: string }[];
  patches: Partial<MirrorSyncState>[];
  puts: { page: MirrorPage; hasContent: boolean }[];
  deletes: MirrorTombstone[];
  structures: number;
  leases: SyncMode[];
  releases: string[];
  childGroupsKnown: string[];
  sweepResults: string[];
}

function fakeStore(initial: Partial<StoreState> = {}): {
  store: SyncStore;
  calls: StoreCalls;
  data: StoreState;
} {
  const data: StoreState = {
    selection: sel([NB]),
    state: initialSyncState(),
    sections: [section()],
    groups: [],
    pages: new Map(),
    pageIdsBySection: new Map(),
    ...initial,
  };

  const calls: StoreCalls = {
    watermarks: [],
    patches: [],
    puts: [],
    deletes: [],
    structures: 0,
    leases: [],
    releases: [],
    childGroupsKnown: [],
    sweepResults: [],
  };

  const store: SyncStore = {
    getSelection: () => Promise.resolve(data.selection),
    getSyncState: () => Promise.resolve(data.state),
    patchSyncState: (patch) => {
      calls.patches.push(patch);
      data.state = { ...data.state, ...patch };
      return Promise.resolve();
    },
    acquireLease: (mode) => {
      calls.leases.push(mode);
      return Promise.resolve();
    },
    releaseLease: (heldSince) => {
      calls.releases.push(heldSince);
      return Promise.resolve();
    },
    putStructure: () => {
      calls.structures += 1;
      return Promise.resolve();
    },
    listSectionsToSync: () => Promise.resolve(data.sections.filter((s) => s.mirrored)),
    listAllSectionGroups: () => Promise.resolve(data.groups),
    setSectionWatermark: (sectionId, watermark) => {
      calls.watermarks.push({ sectionId, watermark });
      return Promise.resolve();
    },
    setSectionSweepResult: (sectionId) => {
      calls.sweepResults.push(sectionId);
      return Promise.resolve();
    },
    setChildGroupsKnown: (groupId) => {
      calls.childGroupsKnown.push(groupId);
      return Promise.resolve();
    },
    getPage: (pageId) => Promise.resolve(data.pages.get(pageId) ?? null),
    putPage: (page: MirrorPage, content: MirrorPageContent | null) => {
      calls.puts.push({ page, hasContent: content !== null });
      data.pages.set(page.id, page);
      return Promise.resolve();
    },
    deletePage: (tombstone) => {
      calls.deletes.push(tombstone);
      return Promise.resolve();
    },
    listPageIdsInSection: (sectionId) =>
      Promise.resolve(data.pageIdsBySection.get(sectionId) ?? []),
  };

  return { store, calls, data };
}

interface BlobCalls {
  ink: string[];
  inkml: string[];
  html: string[];
  deleted: string[];
}

function fakeBlobs(): { blobs: SyncDeps['blobs']; calls: BlobCalls } {
  const calls: BlobCalls = { ink: [], inkml: [], html: [], deleted: [] };
  return {
    calls,
    blobs: {
      putInk: (pageId) => {
        calls.ink.push(pageId);
        return Promise.resolve('ink');
      },
      putInkml: (pageId) => {
        calls.inkml.push(pageId);
        return Promise.resolve('inkml');
      },
      putHtml: (pageId) => {
        calls.html.push(pageId);
        return Promise.resolve('html');
      },
      deleteForPage: (pageId) => {
        calls.deleted.push(pageId);
        return Promise.resolve();
      },
    },
  };
}

interface Harness {
  deps: SyncDeps;
  graphCalls: GraphCalls;
  storeCalls: StoreCalls;
  blobCalls: BlobCalls;
  data: StoreState;
}

function harness(
  script: GraphScript = {},
  storeInit: Partial<StoreState> = {},
  content?: SyncDeps['content'],
): Harness {
  const { graph, calls: graphCalls } = fakeGraph(script);
  const { store, calls: storeCalls, data } = fakeStore(storeInit);
  const { blobs, calls: blobCalls } = fakeBlobs();

  return {
    graphCalls,
    storeCalls,
    blobCalls,
    data,
    deps: {
      graph,
      store,
      blobs,
      now: () => NOW,
      content: content ?? { fetchRaw: () => Promise.resolve(rawHtml('<p>typed</p>')) },
    },
  };
}

const BUDGET = { requestBudget: 100 };

// ---------------------------------------------------------------------------
// buildStructure and the hash
// ---------------------------------------------------------------------------

const TREE: ExpandedNotebook[] = [
  {
    id: NB,
    displayName: '2026',
    lastModifiedDateTime: '2026-08-19T11:00:00Z',
    sections: [{ id: 'sec-1', displayName: 'Index', lastModifiedDateTime: '2026-08-19T11:00:00Z' }],
    sectionGroups: [
      {
        id: 'grp-1',
        displayName: '062 - February',
        sections: [
          { id: 'sec-2', displayName: 'Daily', lastModifiedDateTime: '2026-08-18T09:00:00Z' },
        ],
      },
    ],
  },
  { id: 'nb-2', displayName: 'Other', sections: [], sectionGroups: [] },
];

test('every notebook is stored, and only the selected ones are marked mirrored', () => {
  // The tree read returns the whole account for the same one request. Storing only the
  // selected notebooks would make list_notebooks — which takes no arguments — answer
  // confidently and partially, and a partial answer that cannot be detected as partial is
  // the failure CLAUDE.md names about truncated searches.
  const built = buildStructure(TREE, sel([NB]));

  assert.deepEqual(
    built.notebooks.map((n) => [n.id, n.mirrored]),
    [[NB, true], ['nb-2', false]],
  );
  assert.equal(built.sections.length, 2);
  assert.deepEqual(built.sections.map((s) => s.mirrored), [true, true]);
});

test('a section inside a section group gets the group in its path and as its parent', () => {
  const built = buildStructure(TREE, sel([NB]));
  const nested = built.sections.find((s) => s.id === 'sec-2');

  assert.equal(nested?.parentId, 'grp-1');
  assert.equal(nested?.parentKind, 'sectionGroup');
  assert.equal(nested?.path, '2026 / 062 - February / Daily');
  assert.equal(nested?.notebookId, NB, 'the notebook is denormalised so no query joins');
});

test('childGroupsKnown is false on every group the expanded tree produced', () => {
  // $expand nests two levels, so a first-level group's nested groups are absent from the
  // response rather than known to be empty. The read path treats false as a mirror miss.
  const built = buildStructure(TREE, sel([NB]));
  assert.deepEqual(built.sectionGroups.map((g) => g.childGroupsKnown), [false]);
});

test('a selected notebook id matching nothing is reported rather than ignored', () => {
  // A mistyped id is a notebook that silently never syncs, and the ids are opaque strings
  // nobody can eyeball. This count is the only thing that says so.
  const built = buildStructure(TREE, sel([NB, 'nb-typo']));
  assert.deepEqual(built.unknownNotebookIds, ['nb-typo']);
});

test('the structure hash ignores timestamps and notices everything else', () => {
  // Timestamps move constantly and are read from the live tree rather than the stored
  // copy. Including them would make the hash differ on every run and defeat the skip.
  const base = structureHashOf(buildStructure(TREE, sel([NB])));

  const restamped = structureHashOf(
    buildStructure(
      TREE.map((n) => ({ ...n, lastModifiedDateTime: '2099-01-01T00:00:00Z' })),
      sel([NB]),
    ),
  );
  assert.equal(restamped, base, 'a moved timestamp is not a structure change');

  const renamed = structureHashOf(
    buildStructure(TREE.map((n) => ({ ...n, displayName: `${n.displayName}!` })), sel([NB])),
  );
  assert.notEqual(renamed, base, 'a rename is');

  const reselected = structureHashOf(buildStructure(TREE, sel([NB, 'nb-2'])));
  assert.notEqual(reselected, base, 'and so is a change to the selection');
});

// ---------------------------------------------------------------------------
// pickCandidates
// ---------------------------------------------------------------------------

test('with the roll-up trusted, only sections whose timestamp moved are candidates', () => {
  const state: MirrorSyncState = {
    ...initialSyncState(),
    sectionsScannedThrough: '2026-08-19T11:00:00.000Z',
  };
  const sections = [
    section({ id: 'moved', graphLastModifiedDateTime: '2026-08-19T11:30:00Z' }),
    section({ id: 'still', graphLastModifiedDateTime: '2026-08-01T00:00:00Z' }),
  ];

  assert.deepEqual(
    pickCandidates(sections, state, true).map((s) => s.id),
    ['moved'],
  );
});

test('a never-synced section is always a candidate, however old its timestamp', () => {
  const state: MirrorSyncState = {
    ...initialSyncState(),
    sectionsScannedThrough: '2026-08-19T11:00:00.000Z',
  };
  const sections = [
    section({ id: 'fresh', graphLastModifiedDateTime: '2020-01-01T00:00:00Z', pagesSyncedThrough: null }),
  ];

  assert.deepEqual(pickCandidates(sections, state, true).map((s) => s.id), ['fresh']);
});

test('an absent timestamp behaves exactly like a distrusted one', () => {
  // "The field is absent" and "the timestamp cannot be relied on" must be the same
  // branch, or a service that quietly stopped returning it would silently stop the
  // mirror updating.
  const state: MirrorSyncState = {
    ...initialSyncState(),
    sectionsScannedThrough: '2026-08-19T11:00:00.000Z',
  };
  const sections = [section({ id: 'no-stamp', graphLastModifiedDateTime: null })];

  assert.deepEqual(pickCandidates(sections, state, true).map((s) => s.id), ['no-stamp']);
});

test('with the roll-up distrusted, or the tree stale, every section is a candidate', () => {
  const sections = [
    section({ id: 'a', graphLastModifiedDateTime: '2020-01-01T00:00:00Z' }),
    section({ id: 'b', graphLastModifiedDateTime: '2020-01-01T00:00:00Z' }),
  ];
  const scanned = { ...initialSyncState(), sectionsScannedThrough: '2026-08-19T11:00:00.000Z' };

  assert.equal(pickCandidates(sections, { ...scanned, sectionRollUpTrusted: false }, true).length, 2);
  assert.equal(pickCandidates(sections, scanned, false).length, 2, 'stale tree, so visit all');
});

test('the overlap window is applied, so a section on the boundary is not skipped', () => {
  // Graph's clock and this service's are not the same clock. A section modified fifty
  // minutes before the last pass started still gets visited.
  const state: MirrorSyncState = {
    ...initialSyncState(),
    sectionsScannedThrough: '2026-08-19T12:00:00.000Z',
  };
  const sections = [section({ id: 'just-before', graphLastModifiedDateTime: '2026-08-19T11:10:00Z' })];

  assert.deepEqual(pickCandidates(sections, state, true).map((s) => s.id), ['just-before']);
});

// ---------------------------------------------------------------------------
// Incremental
// ---------------------------------------------------------------------------

test('an empty selection costs no Graph request at all', async () => {
  // The selection document is hand-edited and may legitimately be empty. That is a real
  // state, not an error, and it should not walk the account to discover it.
  const h = harness({}, { selection: sel([]) });

  const report = await runIncremental(h.deps, BUDGET);

  assert.equal(report.graphRequests, 0);
  assert.equal(report.outcome, 'complete');
  assert.equal(h.graphCalls.tree, 0);
});

test('a quiet run reads the tree, writes no structure, and visits nothing', async () => {
  const built = buildStructure(TREE, sel([NB]));
  const h = harness(
    { tree: TREE },
    {
      state: {
        ...initialSyncState(),
        structureHash: structureHashOf(built),
        sectionsScannedThrough: '2026-08-19T11:55:00.000Z',
      },
      sections: [section({ graphLastModifiedDateTime: '2026-08-01T00:00:00Z' })],
    },
  );

  const report = await runIncremental(h.deps, BUDGET);

  assert.equal(report.graphRequests, 1, 'the tree, and nothing else');
  assert.equal(h.storeCalls.structures, 0, 'an unchanged hash skips every structure write');
  assert.deepEqual(h.graphCalls.changedSince, []);
});

test('a changed page is fetched and stored, and the watermark moves to the pass start', async () => {
  const h = harness({ tree: TREE, changed: { 'sec-1': [summary('p-1')] } });

  const report = await runIncremental(h.deps, BUDGET);

  assert.equal(report.pagesUpdated, 1);
  assert.equal(h.storeCalls.puts.length, 1);
  assert.equal(h.storeCalls.puts[0]?.hasContent, true, 'small HTML stays in Firestore');

  // Not the newest page's timestamp. Graph's clock and this service's are not the same
  // clock, and the hour of overlap in overlapFrom is what covers the difference.
  assert.deepEqual(h.storeCalls.watermarks, [{ sectionId: 'sec-1', watermark: NOW_ISO }]);
  assert.notEqual(NOW_ISO, summary('p-1').lastModifiedDateTime);
});

test('an unchanged content hash writes nothing and renders no ink', async () => {
  // What makes the hour of watermark overlap nearly free: a page re-read because it fell
  // inside the window costs one request and no write.
  const first = harness({ tree: TREE, changed: { 'sec-1': [summary('p-1')] } });
  await runIncremental(first.deps, BUDGET);
  const stored = first.data.pages.get('p-1');
  assert.ok(stored);

  const second = harness(
    { tree: TREE, changed: { 'sec-1': [summary('p-1')] } },
    { pages: new Map([['p-1', stored]]) },
  );

  const report = await runIncremental(second.deps, BUDGET);

  assert.equal(second.storeCalls.puts.length, 0, 'nothing written');
  assert.deepEqual(second.blobCalls.ink, [], 'and nothing rendered');
  assert.equal(report.pagesUpdated, 0);
  // The watermark still advances: the section was visited and found unchanged.
  assert.deepEqual(second.storeCalls.watermarks, [{ sectionId: 'sec-1', watermark: NOW_ISO }]);
});

test('a section whose listing failed keeps its old watermark', async () => {
  // The next run has to retry it. Advancing here would skip every page the failed listing
  // never reported, permanently.
  const h = harness({
    tree: TREE,
    changed: {
      'sec-1': () => {
        throw graphError(500);
      },
      'sec-2': [summary('p-2')],
    },
  });
  h.data.sections = [section({ id: 'sec-1' }), section({ id: 'sec-2', path: '2026 / Other' })];

  const report = await runIncremental(h.deps, BUDGET);

  assert.deepEqual(
    h.storeCalls.watermarks.map((w) => w.sectionId),
    ['sec-2'],
    'only the section that succeeded advances',
  );
  assert.equal(report.pagesFailed, 1);
  assert.equal(report.pagesUpdated, 1, 'and the run carries on past the failure');
});

test('the tree read failing is survived, not fatal', async () => {
  // $expand on /notebooks was measured unavailable for minutes at a time on 2026-08-19
  // while un-expanded calls answered 200. Skipping a whole poll cycle over the
  // slowest-changing thing in the account would be the more expensive mistake.
  const h = harness({
    tree: () => {
      throw graphError(500);
    },
    changed: { 'sec-1': [summary('p-1')] },
  });

  const report = await runIncremental(h.deps, BUDGET);

  assert.equal(report.treeRead, false, 'and the report says the structure is stale');
  assert.equal(report.outcome, 'complete');
  assert.equal(report.pagesUpdated, 1, 'the page pass ran against stored structure');
  assert.equal(h.storeCalls.structures, 0);
  assert.ok(h.storeCalls.patches.some((p) => p.lastTreeFailureAt === NOW_ISO));
});

test('a stale tree makes every section a candidate, whatever its stored timestamp', async () => {
  const h = harness(
    {
      tree: () => {
        throw graphError(500);
      },
    },
    {
      state: { ...initialSyncState(), sectionsScannedThrough: '2026-08-19T11:59:00.000Z' },
      sections: [section({ id: 'ancient', graphLastModifiedDateTime: '2020-01-01T00:00:00Z' })],
    },
  );

  await runIncremental(h.deps, BUDGET);

  assert.deepEqual(h.graphCalls.changedSince, ['ancient']);
});

test('sectionsScannedThrough advances only when every candidate completed', async () => {
  const partial = harness({ tree: TREE, changed: { 'sec-1': [summary('p-1'), summary('p-2')] } });

  // Budget: tree + list + one page. The second page does not fit.
  const stopped = await runIncremental(partial.deps, { requestBudget: 3 });

  assert.equal(stopped.done, false);
  assert.equal(stopped.outcome, 'budget-exhausted');
  assert.equal(
    partial.storeCalls.patches.some((p) => p.sectionsScannedThrough !== undefined),
    false,
    'advancing it would make the next run skip what this one never reached',
  );
});

test('a budget-exhausted run keeps the watermarks it earned', async () => {
  // The backfill is five hours of slices. Losing a slice's progress would make it never
  // finish.
  const h = harness({
    tree: TREE,
    changed: { 'sec-1': [summary('p-1')], 'sec-2': [summary('p-2')] },
  });
  h.data.sections = [section({ id: 'sec-1' }), section({ id: 'sec-2', path: '2026 / Other' })];

  // tree + sec-1's list + p-1 = 3, leaving nothing for sec-2.
  const report = await runIncremental(h.deps, { requestBudget: 3 });

  assert.equal(report.done, false);
  assert.deepEqual(h.storeCalls.watermarks, [{ sectionId: 'sec-1', watermark: NOW_ISO }]);
  assert.deepEqual(h.graphCalls.changedSince, ['sec-1'], 'sec-2 was never started');
});

test('a 404 during a content fetch tombstones the page rather than failing it', async () => {
  // The page was deleted between the listing and the fetch. This is the lazy tombstoning
  // the issue describes and it costs nothing extra on this path.
  const h = harness({ tree: TREE, changed: { 'sec-1': [summary('p-gone')] } }, {}, {
    fetchRaw: () => Promise.reject(graphError(404)),
  });

  const report = await runIncremental(h.deps, BUDGET);

  assert.equal(report.pagesDeleted, 1);
  assert.equal(report.pagesFailed, 0);
  assert.deepEqual(h.storeCalls.deletes, [
    { id: 'p-gone', sectionId: 'sec-1', notebookId: NB, reason: 'not-found' },
  ]);
  assert.deepEqual(h.blobCalls.deleted, ['p-gone'], 'and its blobs go with it');
});

test('any other content-fetch status leaves the stored copy alone', async () => {
  for (const status of [400, 401, 429, 500, 503]) {
    const h = harness({ tree: TREE, changed: { 'sec-1': [summary('p-1')] } }, {}, {
      fetchRaw: () => Promise.reject(graphError(status)),
    });

    const report = await runIncremental(h.deps, BUDGET);

    assert.equal(report.pagesDeleted, 0, `status ${status}`);
    assert.equal(report.pagesFailed, 1, `status ${status}`);
    assert.deepEqual(h.storeCalls.deletes, [], `status ${status}`);
  }
});

test('oversized HTML spills to the bucket and the document records where it went', async () => {
  const huge = `<p>${'x'.repeat(800_000)}</p>`;
  const h = harness({ tree: TREE, changed: { 'sec-1': [summary('p-big')] } }, {}, {
    fetchRaw: () => Promise.resolve(rawHtml(huge)),
  });

  await runIncremental(h.deps, BUDGET);

  assert.deepEqual(h.blobCalls.html, ['p-big']);
  const put = h.storeCalls.puts[0];
  assert.equal(put?.page.htmlLocation, 'gcs');
  assert.equal(put?.hasContent, false, 'nothing that large goes in a Firestore document');
  assert.ok((put?.page.htmlBytes ?? 0) > 700_000);
});

test('the datetime filter failing falls back to the unfiltered list, once', async () => {
  // Measured accepted on 2026-08-19, so this path should never run. If it does, the
  // fallback costs the same one request and every later section skips straight to it.
  let filtered = 0;
  const { graph, calls } = fakeGraph({ tree: TREE, changed: { 'sec-1': [summary('p-1')] } });
  const failing: SyncDeps['graph'] = {
    ...graph,
    listPagesChangedSince: () => {
      filtered += 1;
      return Promise.reject(graphError(400));
    },
  };

  const { store, calls: storeCalls, data } = fakeStore();
  const { blobs } = fakeBlobs();
  data.sections = [section({ id: 'sec-1' }), section({ id: 'sec-2', path: '2026 / Other' })];

  await runIncremental(
    { graph: failing, store, blobs, now: () => NOW, content: { fetchRaw: () => Promise.resolve(rawHtml('<p>x</p>')) } },
    BUDGET,
  );

  assert.equal(filtered, 1, 'tried once, then never again this run');
  assert.deepEqual(calls.unfiltered, ['sec-1', 'sec-2']);
  assert.ok(storeCalls.patches.some((p) => p.datetimeFilterSupported === false));
});

test('the lease is taken before any work and released after', async () => {
  const h = harness({ tree: TREE });

  await runIncremental(h.deps, BUDGET);

  assert.deepEqual(h.storeCalls.leases, ['incremental']);
  // Released with the timestamp it took the lease with, so a run that has been superseded
  // -- frozen past Cloud Run's 300s cut, resumed after the lease expired on age -- cannot
  // clear whichever run holds it now.
  assert.deepEqual(h.storeCalls.releases, [NOW_ISO]);
});


// ---------------------------------------------------------------------------
// Active notebooks
//
// The rule in one line: a section in an inactive notebook is visited exactly once, by
// the backfill, and never again by an incremental or a scoped or full sweep. Every test
// below is about something that does *not* happen, because a filter that silently
// stopped filtering costs Graph requests against a 400-per-hour budget and nothing in a
// report would look wrong.
// ---------------------------------------------------------------------------

const NB2 = 'nb-2';

/** A section in the second notebook, which the tests below mark inactive. */
function otherSection(overrides: Partial<MirrorSection> = {}): MirrorSection {
  return section({
    id: 'sec-other',
    notebookId: NB2,
    parentId: NB2,
    path: 'Other / Daily',
    ...overrides,
  });
}

test('splitByActivity backfills an inactive section once and then leaves it alone', () => {
  const never = otherSection({ pagesSyncedThrough: null });
  const filled = otherSection({ id: 'sec-filled', pagesSyncedThrough: '2026-08-19T10:00:00Z' });
  const selection = sel([NB, NB2], [NB]);

  const first = splitByActivity([section(), never, filled], selection, true);
  assert.deepEqual(first.eligible.map((s) => s.id), ['sec-1', 'sec-other']);
  assert.equal(first.skippedInactive, 1);

  // A sweep never backfills, so it declines both.
  const swept = splitByActivity([section(), never, filled], selection, false);
  assert.deepEqual(swept.eligible.map((s) => s.id), ['sec-1']);
  assert.equal(swept.skippedInactive, 2);
});

test('a null active list makes every section eligible, which is the rollback', () => {
  const split = splitByActivity([section(), otherSection()], sel([NB, NB2]), false);
  assert.equal(split.eligible.length, 2);
  assert.equal(split.skippedInactive, 0);
});

test('an incremental run backfills an inactive notebook and then stops listing it', async () => {
  const inactive = { selection: sel([NB, NB2], [NB]) };

  const first = harness(
    { tree: TREE, changed: { 'sec-other': [] } },
    { ...inactive, sections: [otherSection({ pagesSyncedThrough: null })] },
  );
  await runIncremental(first.deps, BUDGET);
  assert.deepEqual(first.graphCalls.changedSince, ['sec-other'], 'the backfill runs');

  const second = harness(
    { tree: TREE },
    { ...inactive, sections: [otherSection({ pagesSyncedThrough: '2026-08-19T10:00:00Z' })] },
  );
  const report = await runIncremental(second.deps, BUDGET);
  assert.deepEqual(second.graphCalls.changedSince, [], 'and never runs again');
  assert.equal(report.sectionsSkippedInactive, 1);
});

test('the activity filter still applies when the timestamps cannot be trusted', async () => {
  // pickCandidates returns early on !sectionRollUpTrusted, so a filter folded into it
  // would have to be applied on both sides of that branch. This is the side a fold would
  // most likely miss, and missing it means visiting every archived section every run.
  const h = harness(
    { tree: TREE },
    {
      selection: sel([NB, NB2], [NB]),
      sections: [otherSection({ pagesSyncedThrough: '2026-08-19T10:00:00Z' })],
      state: { ...initialSyncState(), sectionRollUpTrusted: false },
    },
  );

  await runIncremental(h.deps, BUDGET);
  assert.deepEqual(h.graphCalls.changedSince, []);
});

test('a scoped and a full sweep skip inactive sections; sweep-all visits them', async () => {
  const init = {
    selection: sel([NB, NB2], [NB]),
    sections: [otherSection()],
    state: { ...initialSyncState(), sectionsScannedThrough: '2020-01-01T00:00:00.000Z' },
  };
  const script = { tree: TREE, ids: { 'sec-other': [] } };

  const scoped = harness(script, structuredClone(init));
  await runSweep(scoped.deps, BUDGET);
  assert.deepEqual(scoped.graphCalls.pageIds, []);

  const full = harness(script, structuredClone(init));
  await runFullSweep(full.deps, BUDGET);
  assert.deepEqual(full.graphCalls.pageIds, []);

  // The only thing that reaches a frozen notebook, which is what it is for: nothing else
  // would ever notice a page deleted there in the OneNote client.
  const all = harness(script, structuredClone(init));
  const report = await runSweepAll(all.deps, BUDGET);
  assert.deepEqual(all.graphCalls.pageIds, ['sec-other']);
  assert.equal(report.mode, 'sweep-all');
  assert.equal(report.sectionsSkippedInactive, 0);
});

test('changing the active set nulls sectionsScannedThrough, so a re-activation is seen', async () => {
  // Tier 1 skips a section older than that cutoff, and the cutoff advances on every
  // completed run. Without the reset, a notebook re-activated after three months would
  // have every section older than the cutoff and would never be re-checked at all.
  const h = harness(
    { tree: TREE, changed: { 'sec-1': [] } },
    {
      selection: sel([NB, NB2], [NB]),
      sections: [section({ graphLastModifiedDateTime: '2020-01-01T00:00:00Z' })],
      state: {
        ...initialSyncState(),
        // Set, so a structure rewrite is not what widens this run: an unchanged hash
        // makes the timestamps trusted, and the reset is then the only thing that can
        // turn a section last touched in 2020 into a candidate.
        structureHash: structureHashOf(buildStructure(TREE, sel([NB, NB2], [NB]))),
        sectionsScannedThrough: '2026-08-19T11:55:00.000Z',
        activeSelectionHash: activeSelectionHashOf(sel([NB, NB2], [NB2])),
      },
    },
  );

  await runIncremental(h.deps, BUDGET);

  assert.equal(
    h.storeCalls.patches.some((p) => p.sectionsScannedThrough === null),
    true,
    'the watermark is cleared so every active section is a candidate on this run',
  );
  assert.deepEqual(
    h.graphCalls.changedSince,
    ['sec-1'],
    'and the widening applies to the run that noticed, not the one after it',
  );
});

test('a state document that predates the field records the hash and widens nothing', async () => {
  // Null is "written before this existed", not "the set changed". Treating it as a change
  // would make the first run after every deploy a full-width scan.
  const h = harness(
    { tree: TREE },
    {
      selection: sel([NB]),
      sections: [section({ graphLastModifiedDateTime: '2020-01-01T00:00:00Z' })],
      state: {
        ...initialSyncState(),
        structureHash: structureHashOf(buildStructure(TREE, sel([NB]))),
        sectionsScannedThrough: '2026-08-19T11:55:00.000Z',
      },
    },
  );

  await runIncremental(h.deps, BUDGET);

  assert.deepEqual(h.graphCalls.changedSince, []);
  assert.equal(
    h.storeCalls.patches.some((p) => p.activeSelectionHash === 'all'),
    true,
    'the hash is recorded, so the next real change is detectable',
  );
});

test('the structure hash does not move when only the active set changed', () => {
  // If it did, putStructure would rewrite every section with pagesSyncedThrough null and
  // an activation edit would trigger a full re-backfill — hours of the request budget.
  const base = structureHashOf(buildStructure(TREE, sel([NB])));
  assert.equal(structureHashOf(buildStructure(TREE, sel([NB], []))), base);
  assert.equal(structureHashOf(buildStructure(TREE, sel([NB], [NB]))), base);

  assert.notEqual(activeSelectionHashOf(sel([NB], [])), activeSelectionHashOf(sel([NB], [NB])));
  assert.equal(
    activeSelectionHashOf(sel([NB], ['a', 'b'])),
    activeSelectionHashOf(sel([NB], ['b', 'a'])),
    'the order an operator typed them in is not a change',
  );
});

test('an active id naming no mirrored notebook is counted rather than ignored', () => {
  // Two ways to get here and both are silent: a typo, and a real id never added to
  // notebookIds, whose pages are not mirrored so marking it active reaches nothing.
  const built = buildStructure(TREE, sel([NB], [NB, 'nb-typo', 'nb-2']));
  assert.deepEqual(built.unknownActiveNotebookIds, ['nb-typo', 'nb-2']);
  assert.deepEqual(built.unknownNotebookIds, []);
});

test('a resync writes a page in an inactive notebook', async () => {
  // Activity gates the sync, never the write path. Someone adding a symmetrical check to
  // resyncPage would leave every write to a frozen notebook serving pre-write content
  // until a sweep-all ran, which may be never.
  const h = resyncDeps({
    selection: sel([NB, NB2], [NB]),
    sections: [otherSection()],
  });

  const outcome = await resyncPage(h.deps, 'p-1', { title: 'Written', sectionId: 'sec-other' });

  assert.equal(outcome, 'updated');
  assert.equal(h.storeCalls.puts[0]?.page.title, 'Written');
});

// ---------------------------------------------------------------------------
// Sweep
// ---------------------------------------------------------------------------

test('the sweep deletes exactly the ids Graph no longer returns', async () => {
  const h = harness({ tree: TREE, ids: { 'sec-1': ['p-1', 'p-2'] } });
  h.data.pageIdsBySection.set('sec-1', ['p-1', 'p-2', 'p-gone']);

  const report = await runSweep(h.deps, BUDGET);

  assert.deepEqual(h.storeCalls.deletes, [
    { id: 'p-gone', sectionId: 'sec-1', notebookId: NB, reason: 'sweep' },
  ]);
  assert.equal(report.pagesDeleted, 1);
});

test('a failed enumeration deletes NOTHING', async () => {
  // The single most important assertion in this file. An auth failure or a 500 would
  // otherwise empty the mirror one section at a time, and nothing about a deleted page is
  // recoverable from here.
  for (const status of [401, 429, 500, 503]) {
    const h = harness({
      tree: TREE,
      ids: {
        'sec-1': () => {
          throw graphError(status);
        },
      },
    });
    h.data.pageIdsBySection.set('sec-1', ['p-1', 'p-2', 'p-3']);

    const report = await runSweep(h.deps, BUDGET);

    assert.deepEqual(h.storeCalls.deletes, [], `status ${status}`);
    assert.equal(report.pagesDeleted, 0, `status ${status}`);
    assert.equal(report.pagesFailed, 1, `status ${status}`);
  }
});

test('an empty enumeration on a section that really is empty still deletes', async () => {
  // The guard above is on the *failure*, not on emptiness. A section whose pages were all
  // deleted must be reconciled, or the mirror keeps them forever.
  const h = harness({ tree: TREE, ids: { 'sec-1': [] } });
  h.data.pageIdsBySection.set('sec-1', ['p-1']);

  const report = await runSweep(h.deps, BUDGET);

  assert.equal(report.pagesDeleted, 1);
});

test('the sweep queues ids Graph has that the mirror lacks', async () => {
  // A page moved *into* a mirrored section may not have its own lastModifiedDateTime
  // bumped by the move — the same class of unknown as the section roll-up — so without
  // this it would be invisible until someone next edited it.
  const h = harness({ tree: TREE, ids: { 'sec-1': ['p-1', 'p-new'] } });
  h.data.pageIdsBySection.set('sec-1', ['p-1']);

  const report = await runSweep(h.deps, BUDGET);

  assert.deepEqual(h.storeCalls.puts.map((p) => p.page.id), ['p-new']);
  assert.equal(report.pagesUpdated, 1);
  assert.deepEqual(h.storeCalls.deletes, []);
});

test('the sweep does not advance a section watermark', async () => {
  // Anything it could not fetch inside the budget has to be picked up by the next
  // incremental, which only happens if the watermark stays where it is.
  const h = harness({ tree: TREE, ids: { 'sec-1': ['p-new'] } });

  await runSweep(h.deps, BUDGET);

  assert.deepEqual(h.storeCalls.watermarks, []);
  assert.deepEqual(h.storeCalls.sweepResults, ['sec-1']);
});

test('a budget-exhausted sweep records where to resume', async () => {
  const h = harness({ tree: TREE, ids: { 'sec-1': [], 'sec-2': [] } });
  h.data.sections = [section({ id: 'sec-1' }), section({ id: 'sec-2', path: '2026 / Other' })];

  // tree + sec-1's enumeration = 2, leaving nothing for sec-2.
  const report = await runSweep(h.deps, { requestBudget: 2 });

  assert.equal(report.done, false);
  assert.ok(h.storeCalls.patches.some((p) => p.sweepCursorSectionId === 'sec-2'));
});

test('a scoped sweep visits only moved sections; a full sweep visits all of them', async () => {
  const stale = section({ id: 'still', graphLastModifiedDateTime: '2020-01-01T00:00:00Z' });
  const state = { ...initialSyncState(), sectionsScannedThrough: '2026-08-19T11:00:00.000Z' };

  const scoped = harness({ tree: TREE, ids: { still: [] } }, { sections: [stale], state });
  await runSweep(scoped.deps, BUDGET);
  assert.deepEqual(scoped.graphCalls.pageIds, [], 'nothing moved, so nothing is swept');

  const full = harness({ tree: TREE, ids: { still: [] } }, { sections: [stale], state });
  await runFullSweep(full.deps, BUDGET);
  assert.deepEqual(full.graphCalls.pageIds, ['still'], 'the weekly backstop visits everything');
});

test('the sweep learns the nested section groups $expand could not reach', async () => {
  const h = harness({ tree: TREE }, {
    groups: [
      {
        id: 'grp-1',
        displayName: '062 - February',
        notebookId: NB,
        parentId: NB,
        parentKind: 'notebook',
        mirrored: true,
        path: '2026 / 062 - February',
        childGroupsKnown: false,
      },
    ],
  });

  await runSweep(h.deps, BUDGET);

  assert.deepEqual(h.graphCalls.children, ['grp-1']);
  assert.deepEqual(h.storeCalls.childGroupsKnown, ['grp-1']);
});

// ---------------------------------------------------------------------------
// Resync after a write
//
// One Graph request, triggered by a write tool, so a read immediately after a write
// answers from the mirror instead of falling through until the next scheduled run.
//
// It re-reads content and nothing else. Measured 2026-08-19 (api-overview.md): a PATCH
// is visible to the next content read at 3.7 seconds including both round trips, but
// page metadata is weaker — GET /pages/{id}?$select=title returned "" for pages created
// seconds earlier. So the title comes from the caller and the timestamp is stamped
// locally, and the tests below are what stop either being read back from Graph.
// ---------------------------------------------------------------------------

function resyncDeps(
  storeInit: Partial<StoreState> = {},
  content?: SyncDeps['content'],
): {
  deps: ResyncDeps;
  storeCalls: StoreCalls;
  blobCalls: BlobCalls;
  data: StoreState;
} {
  const { store, calls: storeCalls, data } = fakeStore(storeInit);
  const { blobs, calls: blobCalls } = fakeBlobs();

  return {
    storeCalls,
    blobCalls,
    data,
    deps: {
      store: {
        getPage: store.getPage,
        putPage: store.putPage,
        deletePage: store.deletePage,
        getSection: (sectionId) =>
          Promise.resolve(data.sections.find((s) => s.id === sectionId) ?? null),
      },
      blobs,
      now: () => NOW,
      content: content ?? { fetchRaw: () => Promise.resolve(rawHtml('<p>after the write</p>')) },
    },
  };
}

test('a resync stores the page and stamps the time locally', async () => {
  const h = resyncDeps();

  const outcome = await resyncPage(h.deps, 'p-1', { title: 'Renamed', sectionId: 'sec-1' });

  assert.equal(outcome, 'updated');
  const put = h.storeCalls.puts[0];
  assert.equal(put?.page.id, 'p-1');
  assert.equal(put?.page.title, 'Renamed');
  assert.equal(put?.page.contentState, 'present');
  // Stamped locally rather than read back, because a metadata read here is unreliable.
  assert.equal(put?.page.lastModifiedDateTime, NOW_ISO);
});

test('an append sends no title, and the stored one survives', async () => {
  // An append cannot change a title. Sending an empty one would blank it in the mirror
  // until the next sync corrected it, and a title search would miss the page meanwhile.
  const stored: MirrorPage = {
    id: 'p-1',
    title: 'Monday',
    titleLower: 'monday',
    sectionId: 'sec-1',
    notebookId: NB,
    sectionPath: '2026 / Daily',
    lastModifiedDateTime: '2026-08-19T10:00:00Z',
    contentState: 'present',
    contentHash: 'old',
    htmlLocation: 'firestore',
    htmlObject: null,
    htmlBytes: 5,
    ink: null,
  };
  const h = resyncDeps({ pages: new Map([['p-1', stored]]) });

  await resyncPage(h.deps, 'p-1', {});

  assert.equal(h.storeCalls.puts[0]?.page.title, 'Monday');
});

test('a page in a notebook the mirror does not hold costs no Graph request', async () => {
  // The write tools reach the whole account; only the selection is mirrored.
  let fetched = 0;
  const h = resyncDeps(
    { sections: [section({ mirrored: false })] },
    {
      fetchRaw: () => {
        fetched += 1;
        return Promise.resolve(rawHtml('<p>x</p>'));
      },
    },
  );

  assert.equal(await resyncPage(h.deps, 'p-1', { sectionId: 'sec-1' }), 'not-mirrored');
  assert.equal(fetched, 0);
  assert.deepEqual(h.storeCalls.puts, []);
});

test('a page with no hint and no stored placement is not mirrored', async () => {
  // Nothing says where it belongs, so there is nothing to write and nothing to read.
  const h = resyncDeps();
  assert.equal(await resyncPage(h.deps, 'p-unknown', {}), 'not-mirrored');
});

test('a resync whose page has already gone deletes the mirrored copy', async () => {
  const h = resyncDeps({}, { fetchRaw: () => Promise.reject(graphError(404)) });

  assert.equal(await resyncPage(h.deps, 'p-1', { sectionId: 'sec-1' }), 'deleted');
  assert.deepEqual(h.storeCalls.deletes, [
    { id: 'p-1', sectionId: 'sec-1', notebookId: NB, reason: 'not-found' },
  ]);
  assert.deepEqual(h.blobCalls.deleted, ['p-1']);
});

test('any other Graph failure propagates, for the caller to fall back on', async () => {
  // write-tools catches this and marks the page stale instead, which is correct and just
  // slower. Swallowing it here would leave the mirror holding pre-write content with
  // nothing recording that it is wrong.
  const h = resyncDeps({}, { fetchRaw: () => Promise.reject(graphError(500)) });

  await assert.rejects(() => resyncPage(h.deps, 'p-1', { sectionId: 'sec-1' }), GraphRequestError);
});

test('a resync of unchanged content writes nothing', async () => {
  // An append that produced identical HTML, or a rename, which does not touch content.
  const first = resyncDeps();
  await resyncPage(first.deps, 'p-1', { sectionId: 'sec-1', title: 'Monday' });
  const stored = first.data.pages.get('p-1');
  assert.ok(stored);

  const second = resyncDeps({ pages: new Map([['p-1', stored]]) });
  assert.equal(
    await resyncPage(second.deps, 'p-1', { sectionId: 'sec-1', title: 'Monday' }),
    'unchanged',
  );
  assert.deepEqual(second.storeCalls.puts, []);
});

test('the resync and the sync build the same document from the same response', async () => {
  // One writer, deliberately. A second copy that skipped the ink render or spilled at a
  // different threshold would make a page's stored form depend on which path last touched
  // it, and the difference would surface as a wrong answer to a model days later.
  const raw = rawHtml('<p>identical</p>');

  const viaSync = harness({ tree: TREE, changed: { 'sec-1': [summary('p-1')] } }, {}, {
    fetchRaw: () => Promise.resolve(raw),
  });
  await runIncremental(viaSync.deps, BUDGET);

  const viaWrite = resyncDeps({}, { fetchRaw: () => Promise.resolve(raw) });
  await resyncPage(viaWrite.deps, 'p-1', { sectionId: 'sec-1', title: summary('p-1').title });

  const a = viaSync.storeCalls.puts[0]?.page;
  const b = viaWrite.storeCalls.puts[0]?.page;
  assert.ok(a && b);
  assert.equal(a.contentHash, b.contentHash);
  assert.equal(a.htmlLocation, b.htmlLocation);
  assert.equal(a.htmlBytes, b.htmlBytes);
  assert.equal(a.sectionPath, b.sectionPath);
  assert.equal(a.notebookId, b.notebookId);
});

test('a rename reaches the mirror even though it changes no content', async () => {
  // This was a bug. writePageFromRaw short-circuited on the content hash alone, so
  // update_page_title -- which changes a title and nothing else -- wrote nothing and the
  // mirror kept serving the old title. find_page_by_name and search_pages both match on
  // that title, so the wrong one was live until the next scheduled sync corrected it.
  const first = resyncDeps();
  await resyncPage(first.deps, 'p-1', { sectionId: 'sec-1', title: 'Old Title' });
  const stored = first.data.pages.get('p-1');
  assert.ok(stored);
  assert.equal(stored.title, 'Old Title');

  const renamed = resyncDeps({ pages: new Map([['p-1', stored]]) });
  const outcome = await resyncPage(renamed.deps, 'p-1', { sectionId: 'sec-1', title: 'New Title' });

  assert.equal(outcome, 'updated');
  assert.equal(renamed.storeCalls.puts[0]?.page.title, 'New Title');
  assert.equal(renamed.storeCalls.puts[0]?.page.titleLower, 'new title');
});

test('a page moved to another section is rewritten even with identical content', async () => {
  // Same shape of miss as the rename: page ids are stable across a move, so the content
  // hash matches and only the placement changed.
  const first = resyncDeps();
  await resyncPage(first.deps, 'p-1', { sectionId: 'sec-1', title: 'Monday' });
  const stored = first.data.pages.get('p-1');
  assert.ok(stored);

  const moved = resyncDeps({
    pages: new Map([['p-1', stored]]),
    sections: [section({ id: 'sec-2', displayName: 'Elsewhere', path: '2026 / Elsewhere' })],
  });
  const outcome = await resyncPage(moved.deps, 'p-1', { sectionId: 'sec-2', title: 'Monday' });

  assert.equal(outcome, 'updated');
  assert.equal(moved.storeCalls.puts[0]?.page.sectionId, 'sec-2');
  assert.equal(moved.storeCalls.puts[0]?.page.sectionPath, '2026 / Elsewhere');
});

test('the timestamp alone never forces a rewrite', async () => {
  // lastModifiedDateTime moves on every write, so comparing it would rewrite every page
  // the watermark overlap re-read and defeat the point of the short-circuit.
  const first = harness({ tree: TREE, changed: { 'sec-1': [summary('p-1', '2026-08-19T11:00:00Z')] } });
  await runIncremental(first.deps, BUDGET);
  const stored = first.data.pages.get('p-1');
  assert.ok(stored);

  const later = harness(
    { tree: TREE, changed: { 'sec-1': [summary('p-1', '2026-08-19T11:59:00Z')] } },
    { pages: new Map([['p-1', stored]]) },
  );
  await runIncremental(later.deps, BUDGET);

  assert.deepEqual(later.storeCalls.puts, []);
});
