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
import { createHash } from 'node:crypto';

import { GraphRequestError } from '../src/graph-structure.ts';
import type {
  ContainerChildren,
  ExpandedNotebook,
  PageSummary,
  Section,
} from '../src/graph-structure.ts';
import { setEventSink } from '../src/logging.ts';
import {
  encodeMirrorId,
  groupIdentity,
  initialSyncState,
  NEW_GROUP_DEFAULTS,
  NEW_SECTION_DEFAULTS,
  notebookIdentity,
  planStructureWrite,
  sectionIdentity,
  type MirrorPage,
  type MirrorPageContent,
  type MirrorNotebook,
  type MirrorSection,
  type MirrorSectionGroup,
  type MirrorSyncState,
  type MirrorTombstone,
  type PageStamp,
  type NotebookSelection,
  type StructureIdentity,
  type SyncMode,
} from '../src/mirror-schema.ts';
import {
  buildStructure,
  pickCandidates,
  type CandidateOptions,
  runFullSweep,
  runIncremental,
  writePageFromRaw,
  runSweep,
  runSweepAll,
  resyncPage,
  splitByActivity,
  structureHashOf,
  withLiveMtimes,
  type ResyncDeps,
  type SectionMtimes,
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

/**
 * A stored structure document, which carries an `identity` a fresh fixture does not.
 *
 * `identity` is optional because a document written before that field existed has none,
 * and the store reads its absence as "does not match" — one rewrite each, once. A test
 * that wants the steady state seeds it through the helpers below.
 */
type StoredNotebook = MirrorNotebook & Partial<StructureIdentity>;
type StoredSectionGroup = MirrorSectionGroup & Partial<StructureIdentity>;
type StoredSection = MirrorSection & Partial<StructureIdentity>;

/** A stored section that a matching tree would skip rather than rewrite. */
function settled(doc: MirrorSection): StoredSection {
  return { ...doc, identity: sectionIdentity(doc) };
}

function notebookDoc(overrides: Partial<MirrorNotebook> = {}): StoredNotebook {
  const doc: MirrorNotebook = {
    id: NB,
    displayName: '2026',
    mirrored: true,
    sectionCount: 1,
    sectionGroupCount: 0,
    graphLastModifiedDateTime: null,
    ...overrides,
  };
  return { ...doc, identity: notebookIdentity(doc) };
}

function summary(id: string, modified = '2026-08-19T11:30:00Z'): PageSummary {
  return { id, title: `Page ${id}`, lastModifiedDateTime: modified };
}

/**
 * A stored page as the sweep reads it back.
 *
 * The default timestamp is `summary`'s, so a page seeded in both places *agrees* with
 * Graph unless a test says otherwise — every sweep test written before the stamp
 * comparison existed asserts about deletion and discovery, and a mismatching default
 * would add a content fetch to all of them.
 */
function digest(id: string, modified = '2026-08-19T11:30:00Z'): PageStamp {
  return { id, title: `Page ${id}`, lastModifiedDateTime: modified };
}

const TYPED_HTML = '<p>typed</p>';

/** The sha256 `writePageFromRaw` computes over a page's HTML. */
function hashOf(html: string): string {
  return createHash('sha256').update(html).digest('hex');
}

/**
 * A stored page document, hashed against the harness's default content.
 *
 * A test that wants the short-circuit to fire seeds one of these; one that wants a full
 * write overrides `contentHash`.
 */
function storedPage(overrides: Partial<MirrorPage> = {}): MirrorPage {
  return {
    id: 'p1',
    title: 'Page p1',
    titleLower: 'page p1',
    sectionId: 'sec-1',
    notebookId: NB,
    sectionPath: '2026 / Daily',
    lastModifiedDateTime: '2026-08-19T11:30:00Z',
    contentState: 'present',
    contentHash: hashOf(TYPED_HTML),
    htmlLocation: 'firestore',
    htmlObject: null,
    htmlBytes: TYPED_HTML.length,
    ink: null,
    ...overrides,
  };
}

// ---------------------------------------------------------------------------
// Fakes
// ---------------------------------------------------------------------------

interface GraphScript {
  /**
   * A fixed tree, or a thunk. A thunk that returns is how a test varies the tree between
   * runs without reassigning `deps.graph.getExpandedTree`, which would drop the call
   * counter; a thunk that throws simulates a Graph failure.
   */
  tree?: ExpandedNotebook[] | (() => ExpandedNotebook[]);
  changed?: Record<string, PageSummary[] | (() => never)>;
  summaries?: Record<string, PageSummary[] | (() => never)>;
  children?: ContainerChildren;
}

interface GraphCalls {
  tree: number;
  changedSince: string[];
  unfiltered: string[];
  pageSummaries: string[];
  children: string[];
}

function fakeGraph(script: GraphScript): { graph: SyncDeps['graph']; calls: GraphCalls } {
  const calls: GraphCalls = {
    tree: 0,
    changedSince: [],
    unfiltered: [],
    pageSummaries: [],
    children: [],
  };

  // A scripted value, or a thunk called for it. A thunk that throws simulates a Graph
  // failure; one that returns lets a test change the answer between runs.
  const resolve = <T>(value: T | (() => T) | undefined, fallback: T): T => {
    if (typeof value === 'function') return (value as () => T)();
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
      listPageSummaries: (sectionId) => {
        calls.pageSummaries.push(sectionId);
        return Promise.resolve(resolve(script.summaries?.[sectionId], []));
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
  notebooks: StoredNotebook[];
  sections: StoredSection[];
  groups: StoredSectionGroup[];
  pages: Map<string, MirrorPage>;
  digestsBySection: Map<string, PageStamp[]>;
}

interface StoreCalls {
  watermarks: { sectionId: string; watermark: string }[];
  patches: Partial<MirrorSyncState>[];
  puts: { page: MirrorPage; hasContent: boolean }[];
  deletes: MirrorTombstone[];
  /** One entry per `putPageMetadata`, which is the short-circuit correcting a stamp. */
  metadata: { id: string; lastModifiedDateTime: string }[];
  structures: number;
  /** One entry per structure document actually written. A skipped one is absent. */
  structureWrites: { id: string }[];
  /** Documents the creation defaults were applied to. On an existing one that is a reset. */
  structureResets: string[];
  structureDeletes: string[];
  leases: SyncMode[];
  releases: string[];
  childGroupsKnown: string[];
  sweepResults: string[];
}

/**
 * `#replaceCollection`, over an array in memory.
 *
 * It calls the real `planStructureWrite` rather than reimplementing the decision, so this
 * fake cannot drift from the store it stands in for — and every assertion the tests below
 * make about a skipped or merged document is an assertion about the shipped function. What
 * is faked is only what Firestore does with the plan: `set(…, { merge: true })` is a spread
 * onto the stored document, and a delete removes it.
 *
 * So this asserts the plan and **not** the write mode. The spread below models the merge
 * rather than checking it, which means deleting `{ merge: true }` from the real
 * `batch.set` leaves every test here green while clearing `pagesSyncedThrough` on every
 * section whose identity moved. Nothing on this machine can close that; the check is a
 * live one, on the first post-deploy sync — a section renamed in the OneNote client must
 * come out of it with its watermark still set.
 */
function mergeStructure<S extends { id: string }, T extends { id: string; identity: string }>(
  stored: readonly S[],
  incoming: readonly T[],
  createDefaults: object,
  calls: StoreCalls,
): S[] {
  const byDocumentId = new Map(
    stored.map((doc) => [encodeMirrorId(doc.id), doc as S & Partial<StructureIdentity>]),
  );
  const plan = planStructureWrite(
    new Map([...byDocumentId].map(([documentId, doc]) => [documentId, doc.identity])),
    incoming,
    createDefaults,
  );

  for (const write of plan.writes) {
    const existing = byDocumentId.get(write.documentId);
    calls.structureWrites.push({ id: String(write.fields['id']) });
    if (write.created) calls.structureResets.push(String(write.fields['id']));
    byDocumentId.set(write.documentId, { ...existing, ...write.fields } as S &
      Partial<StructureIdentity>);
  }

  for (const documentId of plan.deletes) {
    calls.structureDeletes.push(byDocumentId.get(documentId)?.id ?? documentId);
    byDocumentId.delete(documentId);
  }

  return [...byDocumentId.values()];
}

function fakeStore(initial: Partial<StoreState> = {}): {
  store: SyncStore;
  calls: StoreCalls;
  data: StoreState;
} {
  const data: StoreState = {
    selection: sel([NB]),
    state: initialSyncState(),
    notebooks: [],
    sections: [section()],
    groups: [],
    pages: new Map(),
    digestsBySection: new Map(),
    ...initial,
  };

  const calls: StoreCalls = {
    watermarks: [],
    patches: [],
    puts: [],
    deletes: [],
    metadata: [],
    structures: 0,
    structureWrites: [],
    structureResets: [],
    structureDeletes: [],
    leases: [],
    releases: [],
    childGroupsKnown: [],
    sweepResults: [],
  };

  /**
   * What Firestore does to the projection when a page document is written.
   *
   * The title as well as the stamp, because `listPageDigestsInSection` reads both back off
   * the same document and the sweep compares both. A fake that re-projected only the stamp
   * left a swept rename disagreeing on every later run, which is the shape of the bug the
   * comparison exists to fix.
   */
  const reproject = (pageId: string, fields: Omit<PageStamp, 'id'>): void => {
    for (const [sectionId, digests] of data.digestsBySection) {
      data.digestsBySection.set(
        sectionId,
        digests.map((d) => (d.id === pageId ? { ...d, ...fields } : d)),
      );
    }
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
    // `#replaceCollection` in src/mirror-store.ts has three behaviours and this fake has
    // to have all three, because that module has no test of its own: a document whose
    // stored `identity` equals the incoming one is skipped, one whose identity differs is
    // merged so the sync-owned fields under it survive, and one the tree no longer holds
    // is deleted. A fake that replaced the collections wholesale — which is what the real
    // store did before the identity split — would let every blast-radius test below pass
    // against nothing.
    putStructure: (structure) => {
      calls.structures += 1;
      data.notebooks = mergeStructure(data.notebooks, structure.notebooks, {}, calls);
      data.groups = mergeStructure(data.groups, structure.sectionGroups, NEW_GROUP_DEFAULTS, calls);
      data.sections = mergeStructure(data.sections, structure.sections, NEW_SECTION_DEFAULTS, calls);
      return Promise.resolve();
    },
    listSectionsToSync: () => Promise.resolve(data.sections.filter((s) => s.mirrored)),
    listAllSectionGroups: () => Promise.resolve(data.groups),
    // Merged onto the stored section, the way the real store's `set(…, { merge: true })`
    // does. A fake that only counted the call left every section looking never-synced, so
    // tier 1 of `pickCandidates` never applied and no test could reach the timestamp
    // filter across two runs.
    setSectionWatermark: (sectionId, watermark) => {
      calls.watermarks.push({ sectionId, watermark });
      data.sections = data.sections.map((s) =>
        s.id === sectionId ? { ...s, pagesSyncedThrough: watermark } : s,
      );
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
    // Both writes apply their effect to the stamp a later `listPageDigestsInSection`
    // reads back, not merely record the call. A fake that only pushed to `calls` would
    // let "a second sweep after a re-fetch disagrees about nothing" pass against a mirror
    // that had forgotten the correction — which is the one property saying the re-fetch
    // converges rather than repeating every run.
    // Both writes stamp `contentSyncedAt`, which is what the real store does with
    // `FieldValue.serverTimestamp()` (src/mirror-store.ts). Modelled with the harness's
    // fixed clock rather than a real one. Without it the settle guard in
    // `storedPageIsCurrent` fails on every document this fake ever wrote, so no test could
    // show a page becoming skippable — which is the property the metadata write buys.
    putPage: (page: MirrorPage, content: MirrorPageContent | null) => {
      calls.puts.push({ page, hasContent: content !== null });
      data.pages.set(page.id, { ...page, contentSyncedAt: NOW_ISO });
      reproject(page.id, { title: page.title, lastModifiedDateTime: page.lastModifiedDateTime });
      return Promise.resolve();
    },
    putPageMetadata: (page) => {
      calls.metadata.push({ id: page.id, lastModifiedDateTime: page.lastModifiedDateTime });
      const existing = data.pages.get(page.id);
      if (existing !== undefined) {
        data.pages.set(page.id, { ...existing, ...page, contentSyncedAt: NOW_ISO });
      }
      reproject(page.id, { title: page.title, lastModifiedDateTime: page.lastModifiedDateTime });
      return Promise.resolve();
    },
    deletePage: (tombstone) => {
      calls.deletes.push(tombstone);
      return Promise.resolve();
    },
    listPageDigestsInSection: (sectionId) =>
      Promise.resolve(data.digestsBySection.get(sectionId) ?? []),
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

/**
 * The expanded tree that would produce `data`'s sections and groups.
 *
 * A test that scripts no tree used to get `[]`, which was harmless while the store fake
 * ignored `putStructure`. With a faithful fake an empty tree deletes every section, so
 * the default has to describe what the store was seeded with.
 *
 * It describes `NB` alone, with every stored group placed directly under it and every
 * `parentKind: 'notebook'` section too, whatever the section's own `notebookId` says. A
 * test seeding a second notebook, a group inside a group, or a section under a group whose
 * own parent is a group has to script the tree instead. Two tests do seed
 * `sel([NB, NB2], [NB])` against this default and so report `unknownNotebookIds: 1` where
 * they used to report 0 — the selection names a notebook the tree does not return, which
 * is what that counter is for.
 */
function treeFrom(data: StoreState): ExpandedNotebook[] {
  // A conditional spread rather than `?? undefined`: `exactOptionalPropertyTypes` is on,
  // so writing `lastModifiedDateTime: undefined` on an optional property is a type error.
  const asGraphSection = (s: MirrorSection): Section => ({
    id: s.id,
    displayName: s.displayName,
    ...(s.graphLastModifiedDateTime === null
      ? {}
      : { lastModifiedDateTime: s.graphLastModifiedDateTime }),
  });

  return [
    {
      id: NB,
      displayName: '2026',
      sections: data.sections.filter((s) => s.parentKind === 'notebook').map(asGraphSection),
      sectionGroups: data.groups.map((g) => ({
        id: g.id,
        displayName: g.displayName,
        sections: data.sections
          .filter((s) => s.parentKind === 'sectionGroup' && s.parentId === g.id)
          .map(asGraphSection),
      })),
    },
  ];
}

function harness(
  script: GraphScript = {},
  storeInit: Partial<StoreState> = {},
  content?: SyncDeps['content'],
): Harness {
  const { store, calls: storeCalls, data } = fakeStore(storeInit);
  const tree = script.tree ?? treeFrom(data);
  const { graph, calls: graphCalls } = fakeGraph({ ...script, tree });
  const { blobs, calls: blobCalls } = fakeBlobs();

  // A harness starts in the steady state: the stored structure already matches the
  // account, so `reconcileStructure` reads the tree and writes nothing. Deleting this
  // block fails several tests, for two reasons:
  //
  // 1. A structure write happening at all is visible: it counts `structures`, patches
  //    `structureHash`, and a test asserting a quiet run wrote nothing sees it.
  // 2. A seeded document the incoming tree does not name is deleted. Merging protects the
  //    sync-owned fields of a document the tree still holds; it does nothing for one the
  //    tree has dropped, and `treeFrom` describes `NB` alone.
  //
  // What it seeds is the hash of the tree under **this** test's selection, which is the
  // state a run leaves behind, not the state one starts from when the selection has just
  // been edited. A test about a selection change has to seed the hash the *previous*
  // selection left, or it asserts against a state the running system cannot be in — the
  // structure write and the selection record happen one line apart in the same pass.
  if (data.state.structureHash === null && typeof tree !== 'function') {
    data.state = {
      ...data.state,
      structureHash: structureHashOf(buildStructure(tree, data.selection)),
    };
  }

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

test('buildStructure emits tree-owned fields only, and no sync-owned defaults', () => {
  // The whole split. `pagesSyncedThrough`, `pageCount` and `childGroupsKnown` are learned
  // over hours of Graph requests and the tree read knows none of them, so a structure
  // write that carried them would reset the lot. They are creation defaults now, applied
  // by the store to a document that was not there and never to one that was.
  const built = buildStructure(TREE, sel([NB]));

  const asRecord = (value: unknown): Record<string, unknown> => value as Record<string, unknown>;
  const section = asRecord(built.sections[0]);
  assert.equal('pagesSyncedThrough' in section, false);
  assert.equal('pageCount' in section, false);
  assert.equal('childGroupsKnown' in asRecord(built.sectionGroups[0]), false);

  // What those three are instead is asserted on real stored output by the two
  // creation-default tests above, not against the constants: a test that reads
  // `NEW_SECTION_DEFAULTS.pagesSyncedThrough` restates the literal it is checking.
});

test('a selected notebook id matching nothing is reported rather than ignored', () => {
  // A mistyped id is a notebook that silently never syncs, and the ids are opaque strings
  // nobody can eyeball. This count is the only thing that says so.
  const built = buildStructure(TREE, sel([NB, 'nb-typo']));
  assert.deepEqual(built.unknownNotebookIds, ['nb-typo']);

  // And it is absent from the mirrored set, which is why `reconcileStructure` reports that
  // set rather than `selection.notebookIds`: the two differ by exactly the ids above, and
  // a caller expanding "every mirrored notebook" from the selection would name one that
  // does not exist.
  assert.deepEqual(
    built.notebooks.filter((n) => n.mirrored).map((n) => n.id),
    [NB],
  );
});

test('the structure hash ignores timestamps and notices everything else', () => {
  // Timestamps move constantly, so including them would make the hash differ on every run
  // and defeat the skip. `reconcileStructure` returns them separately and `pickCandidates`
  // overlays them, which is what stops the exclusion freezing the stored copies.
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
// What a structure write costs
//
// The requirement these three are written against: a change to one notebook must not
// start work on any other. Before the identity split a structure write set every document
// in the account from the tree alone, so renaming one section — or creating a notebook in
// a corner of the account that is not mirrored — cleared `pagesSyncedThrough` on all 202
// mirrored sections and triggered a full re-backfill, one content fetch per page across
// about 2000 pages. Each test below runs with `requestBudget: 1`, which buys the tree read
// and nothing after it, so what they assert is the structure pass alone.
// ---------------------------------------------------------------------------

const TREE_READ_ONLY = { requestBudget: 1 };

test('renaming one section writes that section and nothing else', async () => {
  const renamed: ExpandedNotebook[] = [
    {
      id: NB,
      displayName: '2026',
      sections: [
        { id: 'sec-1', displayName: 'Daily log', lastModifiedDateTime: '2026-08-19T11:00:00Z' },
        { id: 'sec-2', displayName: 'Weekly', lastModifiedDateTime: '2026-08-19T11:00:00Z' },
      ],
      sectionGroups: [],
    },
  ];

  const h = harness(
    { tree: renamed },
    {
      notebooks: [notebookDoc({ sectionCount: 2 })],
      sections: [
        {
          // A section mid-write, with a page count and a listing hold on it. None of the
          // three is a tree field, and all three are gone from the document if the write
          // stops merging.
          ...settled(section({ id: 'sec-1', displayName: 'Daily', path: '2026 / Daily' })),
          pageCount: 7,
          pendingWrites: 1,
          pendingWritesSince: '2026-08-19T11:59:00.000Z',
        },
        settled(section({ id: 'sec-2', displayName: 'Weekly', path: '2026 / Weekly' })),
      ],
      // A hash that does not match, which is what makes the run write the structure at all.
      state: { ...initialSyncState(), structureHash: 'stale' },
    },
  );

  await runIncremental(h.deps, TREE_READ_ONLY);

  assert.deepEqual(h.storeCalls.structureWrites.map((w) => w.id), ['sec-1']);
  assert.deepEqual(h.storeCalls.structureDeletes, []);
  assert.deepEqual(h.storeCalls.structureResets, [], 'nothing that already existed was created');

  const byId = new Map(h.data.sections.map((s) => [s.id, s]));
  assert.equal(byId.get('sec-1')?.displayName, 'Daily log', 'the rename did land');
  assert.equal(byId.get('sec-1')?.path, '2026 / Daily log');

  // The failure the whole task is about: a watermark cleared here is a full re-backfill of
  // the section, and clearing all 202 of them is five hours of the hourly request budget.
  assert.equal(byId.get('sec-1')?.pagesSyncedThrough, '2026-08-19T10:00:00Z');
  assert.equal(byId.get('sec-2')?.pagesSyncedThrough, '2026-08-19T10:00:00Z');
  assert.equal(byId.get('sec-1')?.pageCount, 7);

  // A cleared hold is narrower damage and just as invisible: `list_pages` on this section
  // would answer from the mirror again while a `create_page` or a rename is still in
  // flight against it.
  assert.equal(byId.get('sec-1')?.pendingWrites, 1);
  assert.equal(byId.get('sec-1')?.pendingWritesSince, '2026-08-19T11:59:00.000Z');
  assert.deepEqual(h.storeCalls.watermarks, [], 'and the run spent no budget on either');
});

test('adding a notebook to the selection leaves every other notebook alone', async () => {
  const NB2 = 'nb-2';
  const tree: ExpandedNotebook[] = [
    {
      id: NB,
      displayName: '2026',
      sections: [
        { id: 'sec-1', displayName: 'Daily', lastModifiedDateTime: '2026-08-19T11:00:00Z' },
      ],
      sectionGroups: [],
    },
    {
      id: NB2,
      displayName: '2025',
      sections: [
        { id: 'sec-2', displayName: 'Archive', lastModifiedDateTime: '2026-08-19T11:00:00Z' },
      ],
      sectionGroups: [],
    },
  ];

  const h = harness(
    { tree },
    {
      // The steady state before the edit: NB selected and backfilled, NB2 stored but not
      // mirrored, so its section has never been synced.
      selection: sel([NB, NB2]),
      notebooks: [
        notebookDoc(),
        notebookDoc({ id: NB2, displayName: '2025', mirrored: false }),
      ],
      sections: [
        settled(section({ id: 'sec-1' })),
        settled(
          section({
            id: 'sec-2',
            displayName: 'Archive',
            notebookId: NB2,
            parentId: NB2,
            path: '2025 / Archive',
            mirrored: false,
            pagesSyncedThrough: null,
          }),
        ),
      ],
      state: { ...initialSyncState(), structureHash: 'stale' },
    },
  );

  await runIncremental(h.deps, TREE_READ_ONLY);

  // `mirrored` moved on the new notebook and on its section, and on nothing else.
  assert.deepEqual(h.storeCalls.structureWrites.map((w) => w.id), ['nb-2', 'sec-2']);
  assert.deepEqual(h.storeCalls.structureResets, []);
  assert.deepEqual(h.storeCalls.structureDeletes, []);

  const byId = new Map(h.data.sections.map((s) => [s.id, s]));
  assert.equal(byId.get('sec-2')?.mirrored, true, 'the new notebook is now mirrored');
  assert.equal(
    byId.get('sec-1')?.pagesSyncedThrough,
    '2026-08-19T10:00:00Z',
    'and the notebook that was already there kept its backfill',
  );
});

test('a section that leaves the tree is still deleted', async () => {
  // Deletion is by absence from the incoming set rather than by identity, so the skip
  // must not reach it: the caller has just read the whole tree in one request, which makes
  // absence a fact rather than an inference.
  const tree: ExpandedNotebook[] = [
    {
      id: NB,
      displayName: '2026',
      sections: [
        { id: 'sec-1', displayName: 'Daily', lastModifiedDateTime: '2026-08-19T11:00:00Z' },
      ],
      sectionGroups: [],
    },
  ];

  const h = harness(
    { tree },
    {
      notebooks: [notebookDoc()],
      sections: [
        settled(section({ id: 'sec-1' })),
        settled(section({ id: 'sec-2', displayName: 'Gone', path: '2026 / Gone' })),
      ],
      state: { ...initialSyncState(), structureHash: 'stale' },
    },
  );

  await runIncremental(h.deps, TREE_READ_ONLY);

  assert.deepEqual(h.storeCalls.structureDeletes, ['sec-2']);
  assert.deepEqual(h.storeCalls.structureWrites, [], 'and the surviving section was skipped');
  assert.deepEqual(h.data.sections.map((s) => s.id), ['sec-1']);
});

test('a section the tree just gained is created ready to be backfilled', async () => {
  // `pagesSyncedThrough: null` on a create is the only thing that makes a new section
  // reachable at all. `listSectionsToSync` orders by that field and a Firestore `orderBy`
  // drops a document that does not carry it, so a section created without it is invisible
  // to every later run and nothing reports a section that was never enumerated. Before
  // this split `buildStructure` sent the field on every write; now only the create
  // default supplies it.
  //
  // Both sections carry a 2020 timestamp, so the tier-1 filter would skip either one on
  // its timestamp alone. The second run is what shows the default doing its job.
  const tree: ExpandedNotebook[] = [
    {
      id: NB,
      displayName: '2026',
      sections: [
        { id: 'sec-1', displayName: 'Daily', lastModifiedDateTime: '2020-01-01T00:00:00Z' },
        { id: 'sec-2', displayName: 'New', lastModifiedDateTime: '2020-01-01T00:00:00Z' },
      ],
      sectionGroups: [],
    },
  ];

  const h = harness(
    { tree },
    {
      notebooks: [notebookDoc({ sectionCount: 2 })],
      sections: [
        settled(section({ id: 'sec-1', graphLastModifiedDateTime: '2020-01-01T00:00:00Z' })),
      ],
      state: {
        ...initialSyncState(),
        structureHash: 'stale',
        sectionsScannedThrough: '2026-08-19T11:00:00.000Z',
      },
    },
  );

  await runIncremental(h.deps, TREE_READ_ONLY);

  assert.deepEqual(h.storeCalls.structureResets, ['sec-2'], 'and only the new one');
  const created = h.data.sections.find((s) => s.id === 'sec-2');
  assert.strictEqual(created?.pagesSyncedThrough, null);
  assert.strictEqual(created?.pageCount, 0);
  assert.equal(
    h.data.sections.find((s) => s.id === 'sec-1')?.pagesSyncedThrough,
    '2026-08-19T10:00:00Z',
    'the section that was already there kept its watermark',
  );

  // The second run writes no structure — the hash matches now — so the timestamp filter
  // applies, and a 2020 timestamp is far behind the cutoff. Only the null watermark can
  // get `sec-2` visited.
  await runIncremental(h.deps, BUDGET);
  assert.deepEqual(h.graphCalls.changedSince, ['sec-2']);
});

test('a section group the tree just gained is created not knowing its nested groups', async () => {
  // `childGroupsKnown: false` is what the read path treats as a mirror miss. Created true,
  // a `list_sections` on this group would answer from the mirror and silently omit every
  // group nested inside it — `$expand` reaches one level, so those are absent from the
  // response rather than known to be empty.
  const tree: ExpandedNotebook[] = [
    {
      id: NB,
      displayName: '2026',
      sections: [],
      sectionGroups: [{ id: 'grp-1', displayName: '062 - February', sections: [] }],
    },
  ];

  const h = harness(
    { tree },
    {
      notebooks: [notebookDoc({ sectionCount: 0, sectionGroupCount: 1 })],
      sections: [],
      state: { ...initialSyncState(), structureHash: 'stale' },
    },
  );

  await runIncremental(h.deps, TREE_READ_ONLY);

  assert.deepEqual(h.storeCalls.structureResets, ['grp-1']);
  assert.strictEqual(h.data.groups.find((g) => g.id === 'grp-1')?.childGroupsKnown, false);
});

// ---------------------------------------------------------------------------
// pickCandidates
// ---------------------------------------------------------------------------

// No observation for any section, so each keeps its stored timestamp and these tests
// assert the filter alone. The overlay itself is asserted directly below them.
const NO_LIVE: SectionMtimes = new Map();

/** No notebook is being caught up on, so the timestamp filter is the only thing deciding. */
const NO_WIDE_SCAN: ReadonlySet<string> = new Set<string>();

function opts(state: MirrorSyncState, mayFilterByTimestamp: boolean): CandidateOptions {
  return { state, mayFilterByTimestamp, wideScanNotebookIds: NO_WIDE_SCAN };
}

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
    pickCandidates(sections, NO_LIVE, opts(state, true)).map((s) => s.id),
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

  assert.deepEqual(pickCandidates(sections, NO_LIVE, opts(state, true)).map((s) => s.id), ['fresh']);
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

  assert.deepEqual(pickCandidates(sections, NO_LIVE, opts(state, true)).map((s) => s.id), ['no-stamp']);
});

test('with the roll-up distrusted, or the tree stale, every section is a candidate', () => {
  const sections = [
    section({ id: 'a', graphLastModifiedDateTime: '2020-01-01T00:00:00Z' }),
    section({ id: 'b', graphLastModifiedDateTime: '2020-01-01T00:00:00Z' }),
  ];
  const scanned = { ...initialSyncState(), sectionsScannedThrough: '2026-08-19T11:00:00.000Z' };

  assert.equal(
    pickCandidates(sections, NO_LIVE, opts({ ...scanned, sectionRollUpTrusted: false }, true)).length,
    2,
  );
  assert.equal(pickCandidates(sections, NO_LIVE, opts(scanned, false)).length, 2, 'no filter, visit all');
});

test('a wide-scan notebook is a candidate whatever the cutoff says', () => {
  // The clause that has to come first: every other one declines these sections, and the
  // cutoff only ever moves forward, so nothing else would make them candidates again.
  const state: MirrorSyncState = {
    ...initialSyncState(),
    sectionsScannedThrough: '2026-08-19T11:00:00.000Z',
  };
  const sections = [
    section({ id: 'widened', graphLastModifiedDateTime: '2020-01-01T00:00:00Z' }),
    section({
      id: 'elsewhere',
      notebookId: 'nb-other',
      graphLastModifiedDateTime: '2020-01-01T00:00:00Z',
    }),
  ];

  assert.deepEqual(
    pickCandidates(sections, NO_LIVE, {
      state,
      mayFilterByTimestamp: true,
      wideScanNotebookIds: new Set([NB]),
    }).map((s) => s.id),
    ['widened'],
    'and a section in a notebook the widening does not name is still skipped',
  );
});

test('the section scan window is fifteen minutes, not an hour', () => {
  // The two windows cost different things. This one keeps an edited section a candidate,
  // at one `listPagesChangedSince` per run for as long as it lasts; an hour of that was
  // the last term worth cutting. `outside` is the section an hour-wide cutoff would still
  // be listing.
  const state: MirrorSyncState = {
    ...initialSyncState(),
    sectionsScannedThrough: '2026-08-19T12:00:00.000Z',
  };
  const sections = [
    section({ id: 'inside', graphLastModifiedDateTime: '2026-08-19T11:50:00Z' }),
    section({ id: 'outside', graphLastModifiedDateTime: '2026-08-19T11:40:00Z' }),
    section({
      id: 'never',
      graphLastModifiedDateTime: '2020-01-01T00:00:00Z',
      pagesSyncedThrough: null,
    }),
    section({
      id: 'widened',
      notebookId: 'nb-2',
      graphLastModifiedDateTime: '2020-01-01T00:00:00Z',
    }),
  ];

  // Both clauses that outrank the cutoff are here rather than in tests of their own,
  // because narrowing the window is the change most likely to take them with it: a
  // never-synced section and a section in a notebook being caught up on are both months
  // older than any cutoff, and both must survive it.
  assert.deepEqual(
    pickCandidates(sections, NO_LIVE, {
      state,
      mayFilterByTimestamp: true,
      wideScanNotebookIds: new Set(['nb-2']),
    }).map((s) => s.id),
    ['inside', 'never', 'widened'],
  );
});

test('the scan cutoff is closed, so a section stamped exactly on it is a candidate', () => {
  // Graph's clock and this service's are not the same clock: the stamp compared is
  // Graph's and the watermark is this process's. The boundary is closed for the same
  // reason `overlapSaveAgeMs` reads a stamp equal to the watermark as inside the window.
  const state: MirrorSyncState = {
    ...initialSyncState(),
    sectionsScannedThrough: '2026-08-19T12:00:00.000Z',
  };
  const sections = [
    section({ id: 'on-it', graphLastModifiedDateTime: '2026-08-19T11:45:00.000Z' }),
    section({ id: 'a-ms-before', graphLastModifiedDateTime: '2026-08-19T11:44:59.999Z' }),
  ];

  assert.deepEqual(
    pickCandidates(sections, NO_LIVE, opts(state, true)).map((s) => s.id),
    ['on-it'],
  );
});

test('withLiveMtimes takes the observed timestamp, and keeps the stored one when there is none', () => {
  // The overlay is what stops `pickCandidates` reading a frozen stored value. Each of the
  // three cases below is a different wrong answer if it is written carelessly.
  const sections = [
    section({ id: 'seen', graphLastModifiedDateTime: '2020-01-01T00:00:00Z' }),
    section({ id: 'unseen', graphLastModifiedDateTime: '2020-01-01T00:00:00Z' }),
    section({ id: 'cleared', graphLastModifiedDateTime: '2020-01-01T00:00:00Z' }),
  ];

  const live = new Map<string, string | null>([
    ['seen', '2026-08-19T11:00:00Z'],
    ['cleared', null],
  ]);

  const byId = new Map(withLiveMtimes(sections, live).map((s) => [s.id, s]));

  assert.equal(byId.get('seen')?.graphLastModifiedDateTime, '2026-08-19T11:00:00Z');

  // Only reachable when the tree read failed, which is also when `timestampsAreFresh` is
  // false — but dropping the stored value here would make the section look never-stamped
  // to anything else that reads it.
  assert.equal(byId.get('unseen')?.graphLastModifiedDateTime, '2020-01-01T00:00:00Z');

  // A live `null` is Graph saying the section carries no timestamp, not "no observation".
  // Falling back to the stored value here would keep filtering on a frozen timestamp,
  // which is the bug the overlay exists to fix.
  assert.equal(byId.get('cleared')?.graphLastModifiedDateTime, null);

  assert.equal(sections[0]?.graphLastModifiedDateTime, '2020-01-01T00:00:00Z', 'and nothing is mutated');
});

test('an edit is still noticed after the structure has stopped changing', async () => {
  // The failure this guards: `graphLastModifiedDateTime` is written only by
  // `putStructure`, which runs only when the tree hash moves — and neither the hash nor a
  // document's identity covers timestamps, so even a run that does write the structure
  // skips a section whose timestamp is the only thing that moved. A stored copy therefore
  // freezes while `sectionsScannedThrough` advances on every completed run, and the
  // section stops being a candidate for ever.
  //
  // Three runs, because two are not enough: the first writes the structure, the second is
  // what advances the cutoff past the frozen timestamp, and only the third can be wrongly
  // skipped.
  let sectionMtime = '2026-08-19T09:00:00Z';
  let pageMtime = '2026-08-19T09:00:00Z';
  let html = '<p>before</p>';

  const tree = (): ExpandedNotebook[] => [
    {
      id: NB,
      displayName: '2026',
      lastModifiedDateTime: sectionMtime,
      sections: [{ id: 'sec-1', displayName: 'Daily', lastModifiedDateTime: sectionMtime }],
      sectionGroups: [],
    },
  ];

  const h = harness(
    // A thunk rather than a fixed value, so each run sees the tree as it is by then and
    // `graphCalls.tree` still counts. The stored hash matches nothing, so the first run
    // writes the structure and the section's timestamp reaches the store the only way it
    // ever does.
    { tree, summaries: {} },
    {
      sections: [],
      state: { ...initialSyncState(), structureHash: 'written-under-an-older-tree' },
    },
    { fetchRaw: () => Promise.resolve(rawHtml(html)) },
  );

  h.deps.graph.listPagesChangedSince = (sectionId, since) => {
    h.graphCalls.changedSince.push(sectionId);
    return Promise.resolve(
      pageMtime >= since ? [{ id: 'p1', title: 'Page', lastModifiedDateTime: pageMtime }] : [],
    );
  };

  let clock = Date.parse('2026-08-19T10:00:00Z');
  const deps = { ...h.deps, now: () => clock };

  await runIncremental(deps, { requestBudget: 100 });
  assert.equal(h.storeCalls.structures, 1, 'the structure is written once, and then never');

  // A quiet run two hours later. Nothing changed, and the cutoff advances past the
  // section's stored timestamp.
  clock += 2 * 3600_000;
  await runIncremental(deps, { requestBudget: 100 });
  assert.equal(h.storeCalls.structures, 1);

  // Two hours after that, someone edits the page in the OneNote client.
  clock += 2 * 3600_000;
  sectionMtime = '2026-08-19T14:00:00Z';
  pageMtime = '2026-08-19T14:00:00Z';
  html = '<p>after the client edit</p>';

  const listedBefore = h.graphCalls.changedSince.length;
  const report = await runIncremental(deps, { requestBudget: 100 });

  assert.ok(
    h.graphCalls.changedSince.length > listedBefore,
    'the section whose timestamp moved must be listed again',
  );
  assert.equal(report.pagesUpdated, 1, 'the edit must reach the mirror');
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
  // No scripted tree, so the harness answers one derived from the store: the account
  // agrees with the mirror, and the section really has not moved. That has to be the tree
  // rather than only the stored copy, because the timestamp the filter reads is the live
  // one.
  const h = harness(
    {},
    {
      state: {
        ...initialSyncState(),
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

test('the page listing still reaches back an hour, whatever the section scan window is', async () => {
  // Two windows, two constants, two reasons. This one decides which pages a section's
  // listing returns, and `storedPageIsCurrent` skips an unchanged one without a request —
  // so widening it costs bytes and it keeps the full hour of margin. Pointing it at
  // `SECTION_SCAN_OVERLAP_MS` would narrow it to fifteen minutes with nothing else
  // failing, and the page below would stop being listed at all.
  const sinceSeen: string[] = [];
  const edited = '2026-08-19T09:30:00Z';

  const h = harness({}, { sections: [section({ pagesSyncedThrough: '2026-08-19T10:00:00Z' })] });
  h.deps.graph.listPagesChangedSince = (sectionId, since) => {
    sinceSeen.push(since);
    return Promise.resolve(edited >= since ? [summary('p1', edited)] : []);
  };

  const report = await runIncremental(h.deps, BUDGET);

  assert.deepEqual(sinceSeen, ['2026-08-19T09:00:00.000Z'], 'the watermark, less an hour');
  assert.equal(report.pagesUpdated, 1, 'and the page only that hour surfaced was fetched');
});

/** Every `sync-overlap-save` line this run wrote, parsed. */
function overlapSaves(lines: readonly string[]): unknown[] {
  return lines
    .map((line) => JSON.parse(line) as { event: string })
    .filter((entry) => entry.event === 'sync-overlap-save');
}

test('a page only the overlap surfaced is logged with its age', async () => {
  const lines: string[] = [];
  setEventSink((line) => lines.push(line));

  // The watermark is 10:00 and the page's stamp is 09:30, so the only reason this page is
  // in the listing at all is the hour `overlapFrom` reaches back. It is also genuinely
  // changed — nothing is stored for it — so the window earned its cost here.
  const h = harness(
    { changed: { 'sec-1': [summary('p1', '2026-08-19T09:30:00Z')] } },
    { sections: [section({ pagesSyncedThrough: '2026-08-19T10:00:00Z' })] },
  );

  const report = await runIncremental(h.deps, BUDGET);
  setEventSink(() => {});

  assert.equal(report.pagesUpdated, 1, 'the page really was written');
  assert.deepEqual(overlapSaves(lines), [{ event: 'sync-overlap-save', ageMs: 1_800_000 }]);
});

test('a page inside the watermark is not logged as an overlap save', async () => {
  const lines: string[] = [];
  setEventSink((line) => lines.push(line));

  // 10:30 is after the 10:00 watermark, so a window of any width would have listed it.
  const h = harness(
    { changed: { 'sec-1': [summary('p1', '2026-08-19T10:30:00Z')] } },
    { sections: [section({ pagesSyncedThrough: '2026-08-19T10:00:00Z' })] },
  );

  const report = await runIncremental(h.deps, BUDGET);
  setEventSink(() => {});

  assert.equal(report.pagesUpdated, 1, 'it was written, and still is not a save');
  assert.deepEqual(overlapSaves(lines), []);
});

test('a page the overlap re-read and found unchanged is not a save', async () => {
  const lines: string[] = [];
  setEventSink((line) => lines.push(line));

  // This is the case the event must never claim. The page's stamp predates the watermark,
  // so the overlap is why it was listed — but the stored copy already matched, so a
  // narrower window would have lost nothing. Counting it as a save would produce evidence
  // arguing to keep a window that is doing no work.
  const h = harness(
    { changed: { 'sec-1': [summary('p1', '2026-08-19T09:30:00Z')] } },
    {
      sections: [section({ pagesSyncedThrough: '2026-08-19T10:00:00Z' })],
      pages: new Map([['p1', storedPage()]]),
    },
  );

  const report = await runIncremental(h.deps, BUDGET);
  setEventSink(() => {});

  assert.equal(report.pagesUpdated, 0, 'nothing was written');
  assert.deepEqual(overlapSaves(lines), []);
});

test('an overlap save carries a number and nothing that could name the page', async () => {
  const lines: string[] = [];
  setEventSink((line) => lines.push(line));

  // The event is read off a log-based metric as a distribution of ages. Everything else
  // in scope at the call site is user content: the page's title, the section's name, the
  // notebook path it sits under. This asserts none of it travels.
  const title = 'Contoso renewal — pricing';
  const sectionName = 'Board minutes';
  const h = harness(
    { changed: { 'sec-1': [{ id: 'p1', title, lastModifiedDateTime: '2026-08-19T09:30:00Z' }] } },
    {
      sections: [
        section({
          displayName: sectionName,
          path: `2026 / ${sectionName}`,
          pagesSyncedThrough: '2026-08-19T10:00:00Z',
        }),
      ],
    },
  );

  await runIncremental(h.deps, BUDGET);
  setEventSink(() => {});

  const raw = lines.filter((line) => line.includes('sync-overlap-save'));
  assert.equal(raw.length, 1);
  const line = raw[0] as string;

  // The exact object: a `deepEqual` against two keys is what forbids a third being added
  // later without this test noticing.
  assert.deepEqual(JSON.parse(line), { event: 'sync-overlap-save', ageMs: 1_800_000 });

  for (const secret of [title, sectionName, '2026 / ', 'p1', 'sec-1', NB, 'http']) {
    assert.equal(line.includes(secret), false, `the line quotes ${secret}`);
  }
});

// ---------------------------------------------------------------------------
// Skipping the content fetch
// ---------------------------------------------------------------------------

/** A stored page in the state the skip needs: settled well after Graph's stamp. */
function settledPage(overrides: Partial<MirrorPage> = {}): MirrorPage {
  return storedPage({
    lastModifiedDateTime: '2026-08-19T11:30:00Z',
    contentSyncedAt: '2026-08-19T11:35:00Z',
    ...overrides,
  });
}

/** A harness whose content client records every page it was asked for. */
function fetchRecorder(): { fetched: string[]; content: SyncDeps['content'] } {
  const fetched: string[] = [];
  return {
    fetched,
    content: {
      fetchRaw: (pageId) => {
        fetched.push(pageId);
        return Promise.resolve(rawHtml(TYPED_HTML));
      },
    },
  };
}

test('an unchanged page inside the overlap window costs no Graph request', async () => {
  const { fetched, content } = fetchRecorder();

  const h = harness(
    { changed: { 'sec-1': [summary('p1', '2026-08-19T11:30:00Z')] } },
    { sections: [section()], pages: new Map([['p1', settledPage()]]) },
    content,
  );

  const report = await runIncremental(h.deps, BUDGET);

  assert.deepEqual(fetched, [], 'the stamp did not move, so nothing needs re-reading');
  assert.deepEqual(h.storeCalls.puts, []);
  assert.deepEqual(h.storeCalls.metadata, [], 'and no stamp needs correcting either');
  assert.equal(report.pagesSkipped, 1);
  assert.equal(report.pagesUpdated, 0, 'a skip is not an update');
  assert.equal(report.graphRequests, 2, 'the tree and the listing, and nothing else');
  assert.deepEqual(h.storeCalls.watermarks, [{ sectionId: 'sec-1', watermark: NOW_ISO }]);

  // A skip is work completed, not work outstanding. `freshnessOf` in src/read-sync.ts
  // requires `outcome === 'complete'` and `done` before a covered read may report
  // `source: "onenote"`, so a skip that left either unset would downgrade the label on
  // every tool answer in the account — and nothing else in this file would show it.
  assert.equal(report.outcome, 'complete');
  assert.equal(report.done, true);
  assert.ok(
    h.storeCalls.patches.some((p) => p.sectionsScannedThrough === NOW_ISO),
    'and the section scan advances, which only a finished run does',
  );
});

test('a skipped page is neither an overlap save nor absent from the report line', async () => {
  const lines: string[] = [];
  setEventSink((line) => lines.push(line));

  // The page's stamp predates the watermark, so the overlap window is the only reason it
  // was listed — and now it is not even fetched. Counting that as a save would produce
  // evidence arguing to keep a window whose whole cost this pass just removed.
  const { fetched, content } = fetchRecorder();
  const h = harness(
    { changed: { 'sec-1': [summary('p1', '2026-08-19T09:30:00Z')] } },
    {
      sections: [section({ pagesSyncedThrough: '2026-08-19T10:00:00Z' })],
      pages: new Map([
        ['p1', settledPage({ lastModifiedDateTime: '2026-08-19T09:30:00Z' })],
      ]),
    },
    content,
  );

  const report = await runIncremental(h.deps, BUDGET);
  setEventSink(() => {});

  assert.deepEqual(fetched, []);
  assert.equal(report.pagesSkipped, 1);
  assert.deepEqual(overlapSaves(lines), []);

  // The count has to reach the log line, not only the returned report: the report is read
  // by a test and by `POST /sync`'s response body, and the log line is the only thing an
  // operator watching the scheduler sees. A skip that showed up in neither would make a
  // run that did all its work look identical to one that did none.
  const completed = lines
    .map((line) => JSON.parse(line) as Record<string, unknown>)
    .filter((entry) => entry['event'] === 'sync-completed');
  assert.equal(completed.length, 1);
  assert.equal(completed[0]?.['pagesSkipped'], 1);
});

test('an unsettled copy is fetched even though the stamps match', async () => {
  const { fetched, content } = fetchRecorder();

  const h = harness(
    { changed: { 'sec-1': [summary('p1', '2026-08-19T11:30:00Z')] } },
    {
      sections: [section()],
      // Five seconds after the stamp: Graph stamps whole seconds, so an edit inside the
      // second it names could have landed after this copy was read and moved nothing.
      pages: new Map([['p1', settledPage({ contentSyncedAt: '2026-08-19T11:30:05Z' })]]),
    },
    content,
  );

  const report = await runIncremental(h.deps, BUDGET);

  assert.deepEqual(fetched, ['p1']);
  assert.equal(report.pagesSkipped, 0);
});

test('a copy that has never settled is fetched, and the fetch is what settles it', async () => {
  // The loop this closes. A stored copy the settle guard refuses is fetched; the content
  // turns out identical, so `writePageFromRaw` short-circuits; the stamps agree, so the
  // metadata write used to be skipped; `contentSyncedAt` therefore never moved and the
  // guard refused the same page on the next run, and the next, for the whole life of the
  // hour-wide listing window. One Graph content request per run per page, against 400 an
  // hour, spent on the freshest pages in the account — the ones the skip exists for.
  //
  // Every page document written before `contentSyncedAt` existed has this shape, so it is
  // the common case immediately after a deploy rather than an edge case. It is also
  // invisible in a run report: `pagesUpdated` and `pagesSkipped` both stay at zero and
  // only `graphRequests` moves.
  const first = fetchRecorder();
  const h1 = harness(
    { changed: { 'sec-1': [summary('p1', '2026-08-19T11:30:00Z')] } },
    { sections: [section()], pages: new Map([['p1', storedPage()]]) },
    first.content,
  );

  const report = await runIncremental(h1.deps, BUDGET);

  assert.deepEqual(first.fetched, ['p1'], 'an unsettleable copy cannot be trusted');
  assert.deepEqual(h1.storeCalls.puts, [], 'and the content turns out identical');
  assert.deepEqual(
    h1.storeCalls.metadata,
    [{ id: 'p1', lastModifiedDateTime: '2026-08-19T11:30:00Z' }],
    'so the one write that breaks the loop is the metadata write',
  );
  assert.equal(report.pagesUpdated, 0, 'settling a copy is not a page update');

  const settledNow = h1.data.pages.get('p1');
  assert.ok(settledNow);

  const second = fetchRecorder();
  const h2 = harness(
    { changed: { 'sec-1': [summary('p1', '2026-08-19T11:30:00Z')] } },
    { sections: [section()], pages: new Map([['p1', settledNow]]) },
    second.content,
  );
  const later = await runIncremental(h2.deps, BUDGET);

  assert.deepEqual(second.fetched, [], 'the next run skips it, which is the whole point');
  assert.equal(later.pagesSkipped, 1);
  assert.deepEqual(h2.storeCalls.metadata, [], 'and spends no second write on it either');
});

test('a copy taken inside the settle window is settled by the fetch too', async () => {
  // The other way the guard fails: `contentSyncedAt` is present but within
  // TIMESTAMP_SETTLE_MS of Graph's stamp, so an edit inside the second Graph names could
  // have landed after the copy was read. Same loop, same fix.
  const first = fetchRecorder();
  const h1 = harness(
    { changed: { 'sec-1': [summary('p1', '2026-08-19T11:30:00Z')] } },
    {
      sections: [section()],
      pages: new Map([['p1', settledPage({ contentSyncedAt: '2026-08-19T11:30:05Z' })]]),
    },
    first.content,
  );

  await runIncremental(h1.deps, BUDGET);

  assert.deepEqual(first.fetched, ['p1']);
  assert.deepEqual(h1.storeCalls.puts, []);
  assert.deepEqual(h1.storeCalls.metadata, [
    { id: 'p1', lastModifiedDateTime: '2026-08-19T11:30:00Z' },
  ]);

  const settledNow = h1.data.pages.get('p1');
  assert.ok(settledNow);

  const second = fetchRecorder();
  const h2 = harness(
    { changed: { 'sec-1': [summary('p1', '2026-08-19T11:30:00Z')] } },
    { sections: [section()], pages: new Map([['p1', settledNow]]) },
    second.content,
  );
  await runIncremental(h2.deps, BUDGET);

  assert.deepEqual(second.fetched, []);
});

test('a renamed page is fetched even though Graph moved no timestamp', async () => {
  // Measured 2026-08-21 (api-overview.md): a PATCH replacing a page's title left
  // `lastModifiedDateTime` at its old value, unchanged two minutes later. So the listing
  // is the only thing that can report a rename, and a skip that trusted the stamp alone
  // would leave the mirror answering by the old title for ever — which is the field
  // `find_page_by_name` and `search_pages` match on.
  const { fetched, content } = fetchRecorder();

  const h = harness(
    {
      changed: {
        'sec-1': [{ id: 'p1', title: 'Renamed', lastModifiedDateTime: '2026-08-19T11:30:00Z' }],
      },
    },
    { sections: [section()], pages: new Map([['p1', settledPage()]]) },
    content,
  );

  const report = await runIncremental(h.deps, BUDGET);

  assert.deepEqual(fetched, ['p1'], 'the title moved, so the page is re-read');
  assert.equal(report.pagesSkipped, 0);
  assert.equal(h.data.pages.get('p1')?.title, 'Renamed', 'and the mirror now holds it');
});

test('a skipped page is read from Firestore once, not twice', async () => {
  const reads: string[] = [];
  const h = harness(
    { changed: { 'sec-1': [summary('p1', '2026-08-19T11:30:00Z')] } },
    { sections: [section()], pages: new Map([['p1', settledPage()]]) },
  );
  const inner = h.deps.store.getPage;
  h.deps.store.getPage = (pageId) => {
    reads.push(pageId);
    return inner(pageId);
  };

  await runIncremental(h.deps, BUDGET);

  assert.deepEqual(reads, ['p1']);
});

test('a page the pre-check declines is read from Firestore once too', async () => {
  // The pre-check has already read the document by the time it decides to fetch, so it
  // hands it on. Without that, every page the sync does *not* skip costs two Firestore
  // reads instead of one — the pre-check's and `writePageFromRaw`'s — and the change
  // would be a regression on the common path rather than a saving. Nothing else in this
  // file would notice, because both reads return the same document.
  const reads: string[] = [];
  const h = harness(
    // A stamp Graph moved, so the page is fetched rather than skipped.
    { changed: { 'sec-1': [summary('p1', '2026-08-19T11:45:00Z')] } },
    { sections: [section()], pages: new Map([['p1', settledPage()]]) },
  );
  const inner = h.deps.store.getPage;
  h.deps.store.getPage = (pageId) => {
    reads.push(pageId);
    return inner(pageId);
  };

  const report = await runIncremental(h.deps, BUDGET);

  assert.equal(report.pagesSkipped, 0, 'this page was fetched');
  assert.deepEqual(reads, ['p1'], 'and read from Firestore exactly once');
});

test('a page the mirror has never seen is fetched, whatever its stamp says', async () => {
  // The skip's guard is that a stored document exists. A create inside the window has
  // none, and answering "current" for it would leave the page out of the mirror
  // permanently — the watermark advances past its stamp and no later incremental lists it.
  const { fetched, content } = fetchRecorder();

  const h = harness(
    { changed: { 'sec-1': [summary('p1', '2026-08-19T11:30:00Z')] } },
    { sections: [section()], pages: new Map() },
    content,
  );

  const report = await runIncremental(h.deps, BUDGET);

  assert.deepEqual(fetched, ['p1']);
  assert.equal(report.pagesSkipped, 0);
  assert.equal(h.storeCalls.puts.length, 1);
});

test('a page marked stale by a write is fetched, not skipped', async () => {
  // `beginWrite` marks a page stale before it touches OneNote, and the mirror serves a
  // stale page as a miss. A skip that ignored `contentState` would leave the page marked
  // stale for ever: no run would ever re-fetch it, so every read of it would fall through
  // to Graph, which is the cost the mirror exists to remove.
  const { fetched, content } = fetchRecorder();

  const h = harness(
    { changed: { 'sec-1': [summary('p1', '2026-08-19T11:30:00Z')] } },
    { sections: [section()], pages: new Map([['p1', settledPage({ contentState: 'stale' })]]) },
    content,
  );

  const report = await runIncremental(h.deps, BUDGET);

  assert.deepEqual(fetched, ['p1']);
  assert.equal(report.pagesSkipped, 0);
  assert.equal(h.data.pages.get('p1')?.contentState, 'present');
});

test('the budget is still checked before a page that turns out to be skipped', async () => {
  // The skip costs no Graph request, so a run could in principle keep skipping past an
  // exhausted budget. It must not: the budget also bounds wall clock, and a section whose
  // pages were only half visited must report `done: false` so the watermark stays put and
  // the next run retries it.
  const { fetched, content } = fetchRecorder();

  const h = harness(
    {
      changed: {
        'sec-1': [summary('p1', '2026-08-19T11:30:00Z'), summary('p2', '2026-08-19T11:30:00Z')],
      },
    },
    {
      sections: [section()],
      pages: new Map([
        ['p1', settledPage()],
        ['p2', settledPage({ id: 'p2', title: 'Page p2', titleLower: 'page p2' })],
      ]),
    },
    content,
  );

  // The tree and the listing exhaust it, so no page is reached at all.
  const report = await runIncremental(h.deps, { requestBudget: 2 });

  assert.deepEqual(fetched, []);
  assert.equal(report.pagesSkipped, 0);
  assert.equal(report.done, false);
  assert.deepEqual(h.storeCalls.watermarks, [], 'and the section keeps its old watermark');
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
  // These deps are built by hand, so nothing seeds the stored hash the way `harness` does.
  // Without it the run rewrites the structure, `mayFilterByTimestamp` is false, and the
  // assertion below would hold whatever the timestamp filter did.
  data.state = { ...data.state, structureHash: structureHashOf(buildStructure(TREE, sel([NB]))) };

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
  const script = { tree: TREE, summaries: { 'sec-other': [] } };

  const scoped = harness(script, structuredClone(init));
  await runSweep(scoped.deps, BUDGET);
  assert.deepEqual(scoped.graphCalls.pageSummaries, []);

  const full = harness(script, structuredClone(init));
  await runFullSweep(full.deps, BUDGET);
  assert.deepEqual(full.graphCalls.pageSummaries, []);

  // The only thing that reaches a frozen notebook, which is what it is for: nothing else
  // would ever notice a page deleted there in the OneNote client.
  const all = harness(script, structuredClone(init));
  const report = await runSweepAll(all.deps, BUDGET);
  assert.deepEqual(all.graphCalls.pageSummaries, ['sec-other']);
  assert.equal(report.mode, 'sweep-all');
  assert.equal(report.sectionsSkippedInactive, 0);
});

test('no sweep backfills a never-synced section in a frozen notebook', async () => {
  // `sweepPass` passes `includeBackfill: false` to `splitByActivity`, and that argument is
  // the whole of the active-notebook saving on the sweep side. Flipping it to true is
  // invisible to every other test here, because they all use a section that is already
  // filled — and a section with `pagesSyncedThrough === null` is eligible whatever its
  // notebook's activity says, so it is the only shape that can show the difference.
  //
  // Flipped, every scoped sweep and every nightly `/sync/sweep/full` would backfill the
  // frozen notebooks in full: a section's worth of page enumerations and content fetches
  // per frozen section, against 400 requests an hour, which is the work freezing a notebook
  // exists to prevent.
  const init = {
    selection: sel([NB, NB2], [NB]),
    sections: [otherSection({ pagesSyncedThrough: null })],
    state: { ...initialSyncState(), sectionsScannedThrough: '2020-01-01T00:00:00.000Z' },
  };
  const script = { tree: TREE, summaries: { 'sec-other': [] } };

  const scoped = harness(script, structuredClone(init));
  await runSweep(scoped.deps, BUDGET);
  assert.deepEqual(scoped.graphCalls.pageSummaries, [], 'a scoped sweep leaves it frozen');

  const full = harness(script, structuredClone(init));
  const fullReport = await runFullSweep(full.deps, BUDGET);
  assert.deepEqual(full.graphCalls.pageSummaries, [], 'and so does the nightly backstop');
  assert.equal(fullReport.sectionsSkippedInactive, 1);

  // The incremental pass is what fills it, exactly once — `splitByActivity`'s
  // `includeBackfill` is true there and false here, and that asymmetry is the feature.
  const incremental = harness(
    { tree: TREE, changed: { 'sec-other': [] } },
    structuredClone(init),
  );
  await runIncremental(incremental.deps, BUDGET);
  assert.deepEqual(incremental.graphCalls.changedSince, ['sec-other']);
});

test('a stale tree makes every section a sweep candidate, whatever its stored timestamp', async () => {
  // The sweep's own copy of the rule the incremental pass has at
  // "a stale tree makes every section a candidate". `mayFilterByTimestamp` is
  // `ctx.tally.treeRead` on both call sites and must be: with the tree read failed,
  // `liveMtimes` is empty, so the filter would compare stored timestamps that only move
  // when a structure rewrite happens. A section really edited an hour ago would be declined
  // by a value frozen months earlier, and the sweep is what notices a page deleted or
  // renamed in the OneNote client.
  const h = harness(
    {
      tree: () => {
        throw graphError(500);
      },
      summaries: { ancient: [] },
    },
    {
      state: { ...initialSyncState(), sectionsScannedThrough: '2026-08-19T11:59:00.000Z' },
      sections: [section({ id: 'ancient', graphLastModifiedDateTime: '2020-01-01T00:00:00Z' })],
    },
  );

  const report = await runSweep(h.deps, BUDGET);

  assert.equal(report.treeRead, false);
  assert.deepEqual(h.graphCalls.pageSummaries, ['ancient']);
});

/**
 * Two mirrored notebooks, one section each, both last touched in 2020.
 *
 * `treeFrom` describes `NB` alone, so the widening tests script this instead: the property
 * they assert is that one notebook's sections are listed and the other's are not, which
 * needs the other notebook to really exist. Both timestamps are older than the
 * `sectionsScannedThrough` these tests seed, so tier 1 declines both sections and the
 * wide-scan set is the only thing that can make either a candidate.
 */
const STALE = '2020-01-01T00:00:00Z';
const TWO_NOTEBOOKS: ExpandedNotebook[] = [
  {
    id: NB,
    displayName: '2026',
    sections: [{ id: 'sec-1', displayName: 'Daily', lastModifiedDateTime: STALE }],
    sectionGroups: [],
  },
  {
    id: NB2,
    displayName: 'Other',
    sections: [{ id: 'sec-other', displayName: 'Daily', lastModifiedDateTime: STALE }],
    sectionGroups: [],
  },
];

/** The stored pair the tree above describes, both filled so neither is a backfill. */
function staleSections(): StoredSection[] {
  return [
    settled(section({ graphLastModifiedDateTime: STALE })),
    settled(otherSection({ graphLastModifiedDateTime: STALE })),
  ];
}

/**
 * A state document with the cutoff already past the sections' timestamps.
 *
 * Every caller passes `mirroredNotebookIdsSeen`, because a null one is "never recorded"
 * and takes the widen-nothing path whatever the selection says.
 */
function seenState(seen: Partial<MirrorSyncState>): MirrorSyncState {
  return {
    ...initialSyncState(),
    sectionsScannedThrough: '2026-08-19T11:55:00.000Z',
    ...seen,
  };
}

test('activating one notebook widens that notebook and lists no other', async () => {
  // The requirement in one line: a change to one notebook starts no work on any other.
  // The old mechanism nulled `sectionsScannedThrough`, which made every mirrored active
  // section a candidate — about 70 listing requests on this account for a change that
  // concerns one notebook.
  const h = harness(
    { tree: TWO_NOTEBOOKS },
    {
      selection: sel([NB, NB2], [NB, NB2]),
      sections: staleSections(),
      state: seenState({
        mirroredNotebookIdsSeen: [NB, NB2],
        activeNotebookIdsSeen: [NB],
      }),
    },
  );

  await runIncremental(h.deps, BUDGET);

  assert.deepEqual(h.graphCalls.changedSince, ['sec-other']);
  assert.equal(
    h.storeCalls.patches.some((p) => p.sectionsScannedThrough === null),
    false,
    'the global cutoff is never nulled by a selection change',
  );
  assert.deepEqual(
    h.storeCalls.patches.find((p) => p.wideScanNotebookIds !== undefined)?.wideScanNotebookIds,
    [NB2],
    'the widening names the notebook that changed and nothing else',
  );
});

test('adding one notebook to the selection widens that notebook and lists no other', async () => {
  // The same mechanism through the other list. Task 2 made watermarks survive a structure
  // write, so a notebook removed and re-added lists from where it left off — but its
  // sections' timestamps are older than the cutoff, so without this it would never become
  // a candidate at all.
  const h = harness(
    { tree: TWO_NOTEBOOKS },
    {
      selection: sel([NB, NB2]),
      sections: staleSections(),
      state: seenState({
        mirroredNotebookIdsSeen: [NB],
        activeNotebookIdsSeen: null,
        // The hash the *previous* selection left behind. The harness would otherwise seed
        // the hash of the new one, which says the structure was already written under a
        // selection the last run had not seen — two writes one line apart in the same
        // pass, so that state exists only if the process died between them.
        structureHash: structureHashOf(buildStructure(TWO_NOTEBOOKS, sel([NB]))),
      }),
    },
  );

  await runIncremental(h.deps, BUDGET);

  assert.deepEqual(h.graphCalls.changedSince, ['sec-other']);
});

test('removing a notebook from either list widens nothing', async () => {
  // Removal leaves no backlog: the notebook this run does less work for has nothing to be
  // caught up on. Widening on a removal would cost a listing request per section of the
  // whole selection for a change that only ever subtracts work.
  const deactivated = harness(
    { tree: TWO_NOTEBOOKS },
    {
      selection: sel([NB, NB2], [NB]),
      sections: staleSections(),
      state: seenState({ mirroredNotebookIdsSeen: [NB, NB2], activeNotebookIdsSeen: [NB, NB2] }),
    },
  );
  await runIncremental(deactivated.deps, BUDGET);
  assert.deepEqual(deactivated.graphCalls.changedSince, []);
  assert.equal(
    deactivated.storeCalls.patches.some(
      (p) => p.wideScanNotebookIds !== undefined && p.wideScanNotebookIds.length > 0,
    ),
    false,
    'nothing was widened',
  );
  // The removal is still recorded, or re-activating this notebook later would diff
  // against a list it is already in and widen nothing.
  assert.deepEqual(
    deactivated.storeCalls.patches.find((p) => p.mirroredNotebookIdsSeen !== undefined)
      ?.activeNotebookIdsSeen,
    [NB],
  );

  const dropped = harness(
    { tree: TWO_NOTEBOOKS },
    {
      selection: sel([NB]),
      sections: staleSections(),
      state: seenState({
        mirroredNotebookIdsSeen: [NB, NB2],
        activeNotebookIdsSeen: null,
        structureHash: structureHashOf(buildStructure(TWO_NOTEBOOKS, sel([NB, NB2]))),
      }),
    },
  );
  await runIncremental(dropped.deps, BUDGET);
  assert.deepEqual(dropped.graphCalls.changedSince, []);
});

test('deactivating then re-activating a notebook widens it the second time', async () => {
  // The cycle the recording exists for. If the deactivation were not recorded, the
  // re-activation would diff against a list still naming this notebook and widen nothing,
  // and its sections would sit below the cutoff for ever.
  const h = harness(
    { tree: TWO_NOTEBOOKS },
    {
      selection: sel([NB, NB2], [NB]),
      sections: staleSections(),
      state: seenState({ mirroredNotebookIdsSeen: [NB, NB2], activeNotebookIdsSeen: [NB, NB2] }),
    },
  );

  await runIncremental(h.deps, BUDGET);
  assert.deepEqual(h.graphCalls.changedSince, [], 'the deactivation widens nothing');

  h.data.selection = sel([NB, NB2], [NB, NB2]);
  await runIncremental(h.deps, BUDGET);
  assert.deepEqual(h.graphCalls.changedSince, ['sec-other']);

  // And a third run over an unchanged selection writes nothing about it: on a steady
  // account this is every run, and a write per run would be a Firestore write per poll.
  h.storeCalls.patches.length = 0;
  await runIncremental(h.deps, BUDGET);
  assert.equal(
    h.storeCalls.patches.some((p) => p.mirroredNotebookIdsSeen !== undefined),
    false,
  );
});

test('a budget-stopped run keeps the wide-scan set; a completed one clears it', async () => {
  // The set is stored rather than held for one run because a run may stop with sections
  // outstanding. Clearing it on a run that never reached them would leave those sections
  // below the cutoff for ever, which is the failure the widening exists to prevent.
  const seeded = {
    selection: sel([NB, NB2], [NB, NB2]),
    sections: staleSections(),
    state: seenState({ mirroredNotebookIdsSeen: [NB, NB2], activeNotebookIdsSeen: [NB] }),
  };

  // One request, spent on the tree read, so the candidate loop stops before its first
  // section.
  const stopped = harness({ tree: TWO_NOTEBOOKS }, structuredClone(seeded));
  await runIncremental(stopped.deps, TREE_READ_ONLY);
  assert.deepEqual(stopped.graphCalls.changedSince, []);
  assert.equal(
    stopped.storeCalls.patches.some(
      (p) => p.wideScanNotebookIds !== undefined && p.wideScanNotebookIds.length === 0,
    ),
    false,
    'nothing cleared the set',
  );

  const completed = harness({ tree: TWO_NOTEBOOKS }, structuredClone(seeded));
  await runIncremental(completed.deps, BUDGET);
  assert.equal(
    completed.storeCalls.patches.some(
      (p) => p.wideScanNotebookIds !== undefined && p.wideScanNotebookIds.length === 0,
    ),
    true,
    'and a run that visited every candidate does clear it',
  );
});

test('a run that read no tree records no selection, so the next one still widens', async () => {
  // `mirroredNotebookIds` is empty when the tree read did not happen, so an
  // `activeNotebookIds` of `null` would resolve to "nothing is active", widen nothing, and
  // still record `activeNotebookIdsSeen: null`. The next healthy run would then diff null
  // against null and the unfrozen notebook would stay below the cutoff for ever.
  const seeded = {
    selection: sel([NB, NB2]),
    sections: staleSections(),
    state: seenState({ mirroredNotebookIdsSeen: [NB, NB2], activeNotebookIdsSeen: [NB] }),
  };

  // Entry point one: the budget runs out before the tree read.
  const broke = harness({ tree: TWO_NOTEBOOKS }, structuredClone(seeded));
  await runIncremental(broke.deps, { requestBudget: 0 });
  assert.equal(
    broke.storeCalls.patches.some((p) => p.mirroredNotebookIdsSeen !== undefined),
    false,
    'nothing was recorded',
  );
  await runIncremental(broke.deps, BUDGET);
  assert.deepEqual(broke.graphCalls.changedSince, ['sec-other'], 'so the next run widens');

  // Entry point two: `getExpandedTree` answers the 500 measured on 2026-08-19.
  let failed = true;
  const flaky = harness(
    {
      tree: () => {
        if (failed) throw graphError(500);
        return TWO_NOTEBOOKS;
      },
    },
    structuredClone(seeded),
  );
  // A thunked tree skips the harness's steady-state seeding, so the hash the previous
  // selection left has to be seeded by hand.
  flaky.data.state = {
    ...flaky.data.state,
    structureHash: structureHashOf(buildStructure(TWO_NOTEBOOKS, sel([NB, NB2]))),
  };

  await runIncremental(flaky.deps, BUDGET);
  assert.equal(
    flaky.storeCalls.patches.some((p) => p.mirroredNotebookIdsSeen !== undefined),
    false,
  );
  // A failed tree read leaves the stored timestamps stale, so that run visited every
  // section — which is the existing rule and not the widening. Only the second run's calls
  // say anything about this test.
  flaky.graphCalls.changedSince.length = 0;

  failed = false;
  await runIncremental(flaky.deps, BUDGET);
  assert.deepEqual(flaky.graphCalls.changedSince, ['sec-other']);
});

test('a scoped sweep visits a widened notebook, and does not clear the set', async () => {
  // The sweep is what reconciles a section's pages against Graph, so a widened notebook it
  // skipped would leave pages deleted while it was frozen in the mirror with nothing to
  // notice them. It reads
  // the set and never clears it: only a completed incremental run has visited what the
  // widening is for.
  const h = harness(
    { tree: TWO_NOTEBOOKS, summaries: { 'sec-other': [], 'sec-1': [] } },
    {
      selection: sel([NB, NB2], [NB, NB2]),
      sections: staleSections(),
      state: seenState({
        mirroredNotebookIdsSeen: [NB, NB2],
        activeNotebookIdsSeen: [NB, NB2],
        wideScanNotebookIds: [NB2],
      }),
    },
  );

  await runSweep(h.deps, BUDGET);

  assert.deepEqual(h.graphCalls.pageSummaries, ['sec-other']);
  assert.equal(
    h.storeCalls.patches.some((p) => p.wideScanNotebookIds !== undefined),
    false,
  );
});

test('a state document that predates the selection fields records them and widens nothing', async () => {
  // `mirroredNotebookIdsSeen` null is "written before these fields existed", not "the
  // selection changed". Treating it as a change would make the first run after this deploy
  // widen every mirrored notebook.
  const h = harness(
    { tree: TWO_NOTEBOOKS },
    {
      selection: sel([NB, NB2], [NB, NB2]),
      sections: staleSections(),
      state: { ...initialSyncState(), sectionsScannedThrough: '2026-08-19T11:55:00.000Z' },
    },
  );

  await runIncremental(h.deps, BUDGET);

  assert.deepEqual(h.graphCalls.changedSince, []);
  const recorded = h.storeCalls.patches.find((p) => p.mirroredNotebookIdsSeen !== undefined);
  assert.deepEqual(recorded?.mirroredNotebookIdsSeen, [NB, NB2]);
  assert.deepEqual(recorded?.activeNotebookIdsSeen, [NB, NB2]);
  assert.deepEqual(recorded?.wideScanNotebookIds, []);
});

test('an active list becoming null widens every mirrored notebook that was frozen', async () => {
  // Null is "every mirrored notebook is active", so dropping the list activates whatever
  // was not in it — and only that.
  const h = harness(
    { tree: TWO_NOTEBOOKS },
    {
      selection: sel([NB, NB2]),
      sections: staleSections(),
      state: seenState({ mirroredNotebookIdsSeen: [NB, NB2], activeNotebookIdsSeen: [NB] }),
    },
  );

  await runIncremental(h.deps, BUDGET);

  assert.deepEqual(h.graphCalls.changedSince, ['sec-other']);
});

test('the structure hash does not move when only the active set changed', () => {
  // If it did, an activation edit would write a structure the documents say nothing about.
  // The two changes need different responses, and `reconcileSelection` gives the selection
  // lists their own.
  const base = structureHashOf(buildStructure(TREE, sel([NB])));
  assert.equal(structureHashOf(buildStructure(TREE, sel([NB], []))), base);
  assert.equal(structureHashOf(buildStructure(TREE, sel([NB], [NB]))), base);
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
  const h = harness({ tree: TREE, summaries: { 'sec-1': [summary('p-1'), summary('p-2')] } });
  h.data.digestsBySection.set('sec-1', [digest('p-1'), digest('p-2'), digest('p-gone')]);

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
      summaries: {
        'sec-1': () => {
          throw graphError(status);
        },
      },
    });
    h.data.digestsBySection.set('sec-1', [digest('p-1'), digest('p-2'), digest('p-3')]);

    const report = await runSweep(h.deps, BUDGET);

    assert.deepEqual(h.storeCalls.deletes, [], `status ${status}`);
    assert.equal(report.pagesDeleted, 0, `status ${status}`);
    assert.equal(report.pagesFailed, 1, `status ${status}`);
  }
});

test('an empty enumeration on a section that really is empty still deletes', async () => {
  // The guard above is on the *failure*, not on emptiness. A section whose pages were all
  // deleted must be reconciled, or the mirror keeps them forever.
  const h = harness({ tree: TREE, summaries: { 'sec-1': [] } });
  h.data.digestsBySection.set('sec-1', [digest('p-1')]);

  const report = await runSweep(h.deps, BUDGET);

  assert.equal(report.pagesDeleted, 1);
});

test('the sweep queues ids Graph has that the mirror lacks', async () => {
  // Measured 2026-08-21: a page moved between sections keeps the lastModifiedDateTime it
  // already had, so it is below every later section watermark and listPagesChangedSince
  // never returns it. Without this loop it is invisible until someone next edits it.
  const h = harness({ tree: TREE, summaries: { 'sec-1': [summary('p-1'), summary('p-new')] } });
  h.data.digestsBySection.set('sec-1', [digest('p-1')]);

  const report = await runSweep(h.deps, BUDGET);

  assert.deepEqual(h.storeCalls.puts.map((p) => p.page.id), ['p-new']);
  assert.equal(report.pagesUpdated, 1);
  assert.deepEqual(h.storeCalls.deletes, []);
});

test('the sweep stores a discovered page with its real title and timestamp', async () => {
  // Before listPageSummaries the sweep synthesized title '' and 1970-01-01 here, and
  // neither self-heals through the incremental pass: both fields reach the calling model,
  // and a page moved into a section keeps its old timestamp (measured 2026-08-21), so no
  // later incremental lists it.
  const h = harness(
    { summaries: { 'sec-1': [summary('new-page', '2026-08-19T11:45:00Z')] } },
    { sections: [section()] },
  );

  await runFullSweep(h.deps, BUDGET);

  const written = h.storeCalls.puts.find((p) => p.page.id === 'new-page');
  assert.ok(written, 'the page Graph has and the mirror lacks must be fetched');
  assert.equal(written.page.title, 'Page new-page');
  assert.equal(written.page.lastModifiedDateTime, '2026-08-19T11:45:00Z');
});

test('stamps a second apart send the sweep to Graph and leave the content in place', async () => {
  // The old rule marked a disagreeing page stale, and `markPageStale` deletes the
  // page-content document. Nothing re-fetches a stale page: the incremental will not list
  // a page whose Graph stamp is behind the section watermark, no read path writes to the
  // mirror, and the sweep skipped a non-`present` copy. So a mark was permanent, and the
  // tolerance that guarded against it discarded every edit inside its window. A re-fetch
  // has neither problem — the content hash decides, and a false positive costs one
  // request.
  const h = harness(
    { summaries: { 'sec-1': [summary('p1', '2026-08-19T13:00:01Z')] } },
    {
      sections: [section()],
      digestsBySection: new Map([['sec-1', [digest('p1', '2026-08-19T13:00:00Z')]]]),
    },
  );

  await runFullSweep(h.deps, BUDGET);

  assert.deepEqual(h.graphCalls.pageSummaries, ['sec-1']);
  assert.deepEqual(h.storeCalls.puts.map((p) => p.page.id), ['p1'], 'the page is re-fetched');
  assert.deepEqual(h.storeCalls.deletes, [], 'and nothing is deleted over a stamp');
});

test('stamps a day apart re-fetch the same way', async () => {
  const h = harness(
    { summaries: { 'sec-1': [summary('p1', '2026-08-20T13:00:00Z')] } },
    {
      sections: [section()],
      digestsBySection: new Map([['sec-1', [digest('p1', '2026-08-19T13:00:00Z')]]]),
    },
  );

  await runFullSweep(h.deps, BUDGET);

  assert.deepEqual(h.storeCalls.puts.map((p) => p.page.id), ['p1']);
});

test('a page stored ahead of Graph is re-fetched, and that is the repair', async () => {
  // `resyncPage` stamps `lastModifiedDateTime` from this process's clock after the write
  // returns, so every page written through this server is stored *ahead* of the stamp
  // Graph recorded for the same edit. Under the old mark that direction had to be excluded
  // or the mirror lost its most-used pages permanently. Under a re-fetch it is the case
  // that fixes itself: the page comes back with Graph's own string in place of the local
  // one, and the next sweep finds the two agreeing.
  const h = harness(
    { summaries: { 'sec-1': [summary('p1', '2026-08-19T13:00:00Z')] } },
    {
      sections: [section()],
      digestsBySection: new Map([['sec-1', [digest('p1', '2026-08-19T13:00:02.481Z')]]]),
    },
  );

  const report = await runFullSweep(h.deps, BUDGET);

  const written = h.storeCalls.puts.find((put) => put.page.id === 'p1');
  assert.ok(written, 'stored-ahead is a disagreement like any other');
  assert.equal(written.page.lastModifiedDateTime, '2026-08-19T13:00:00Z', "Graph's string wins");
  assert.equal(report.pagesFailed, 0, 'and it is not an error');
  assert.deepEqual(h.storeCalls.deletes, []);
});

test('a page renamed outside this server is swept, though Graph moved no timestamp', async () => {
  // Measured 2026-08-21 (api-overview.md, "A page rename does not move the page's
  // `lastModifiedDateTime`"): a PATCH replacing a title left the stamp at its old value,
  // still unchanged two minutes later. So a sweep that compared stamps alone was blind to
  // every rename, and the sweep is the only pass that looks at a page whose stamp has not
  // moved — the incremental never lists one, because its stamp is below the section
  // watermark.
  //
  // The reachable route to this state needs nothing unmeasured: `update_page_title` calls
  // `markPageStale`, which leaves `lastModifiedDateTime` alone, the PATCH renames the page
  // without moving it, and a `resyncPage` that hits a transient 429 is documented as
  // non-fatal. The mirror is then left holding the old title under a stamp that agrees
  // with Graph, and `find_page_by_name`, `search_pages` and every listing answer by it.
  const h = harness(
    {
      summaries: {
        'sec-1': [{ id: 'p1', title: 'Renamed', lastModifiedDateTime: '2026-08-19T11:30:00Z' }],
      },
    },
    {
      sections: [section()],
      pages: new Map([['p1', storedPage()]]),
      digestsBySection: new Map([['sec-1', [digest('p1', '2026-08-19T11:30:00Z')]]]),
    },
  );

  const report = await runFullSweep(h.deps, BUDGET);

  assert.equal(report.graphRequests, 3, 'tree, enumeration, and the content fetch the title forced');
  assert.equal(h.storeCalls.puts[0]?.page.title, 'Renamed');
  assert.equal(h.storeCalls.puts[0]?.page.titleLower, 'renamed', 'which is what by-name matching reads');
  assert.deepEqual(h.storeCalls.deletes, [], 'a rename is not a deletion');
});

test('a swept rename converges: the second sweep fetches nothing', async () => {
  // The other half. A comparison that fired every run would put a content request per
  // renamed page into every sweep for ever, against 400 an hour.
  const h = harness(
    {
      summaries: {
        'sec-1': [{ id: 'p1', title: 'Renamed', lastModifiedDateTime: '2026-08-19T11:30:00Z' }],
      },
    },
    {
      sections: [section()],
      pages: new Map([['p1', storedPage()]]),
      digestsBySection: new Map([['sec-1', [digest('p1', '2026-08-19T11:30:00Z')]]]),
    },
  );

  const first = await runFullSweep(h.deps, BUDGET);
  assert.equal(first.graphRequests, 3);

  const second = await runFullSweep(h.deps, BUDGET);
  assert.equal(second.graphRequests, 2, 'the titles now agree, so no page is fetched');
});

test('stamps that agree cost neither a fetch nor a write', async () => {
  // The other half of the bargain. Every page in every swept section reads its stamp back,
  // so a comparison that fired on agreement would put a content request per mirrored page
  // inside a run sized against 400 requests an hour.
  const h = harness(
    { summaries: { 'sec-1': [summary('p1', '2026-08-19T13:00:00Z')] } },
    {
      sections: [section()],
      digestsBySection: new Map([['sec-1', [digest('p1', '2026-08-19T13:00:00Z')]]]),
    },
  );

  const report = await runFullSweep(h.deps, BUDGET);

  assert.equal(report.graphRequests, 2, 'the tree and the section enumeration, and nothing else');
  assert.deepEqual(h.storeCalls.puts, []);
  assert.deepEqual(h.storeCalls.metadata, []);
  assert.equal(report.pagesUpdated, 0);
});

test('a re-fetch whose content is unchanged writes metadata and is not an update', async () => {
  // This is what makes the re-fetch converge. Without the metadata write the stored stamp
  // stays where it was, the next sweep disagrees with Graph again, and the same page is
  // fetched on every run for ever.
  const h = harness(
    { summaries: { 'sec-1': [summary('p1', '2026-08-19T13:00:00Z')] } },
    {
      sections: [section()],
      pages: new Map([['p1', storedPage({ lastModifiedDateTime: '2026-08-19T11:00:00Z' })]]),
      digestsBySection: new Map([['sec-1', [digest('p1', '2026-08-19T11:00:00Z')]]]),
    },
  );

  const report = await runFullSweep(h.deps, BUDGET);

  assert.deepEqual(h.storeCalls.puts, [], 'the content is identical, so no page write');
  assert.deepEqual(h.storeCalls.metadata, [
    { id: 'p1', lastModifiedDateTime: '2026-08-19T13:00:00Z' },
  ]);
  assert.equal(report.pagesUpdated, 0, 'nothing changed, so nothing is reported as changed');
});

test('a second sweep after a re-fetch disagrees about nothing', async () => {
  // Convergence, through the real transition rather than a hand-seeded document: the
  // metadata write above has to land where `listPageDigestsInSection` reads it back.
  const h = harness(
    { summaries: { 'sec-1': [summary('p1', '2026-08-19T13:00:00Z')] } },
    {
      sections: [section()],
      pages: new Map([['p1', storedPage({ lastModifiedDateTime: '2026-08-19T11:00:00Z' })]]),
      digestsBySection: new Map([['sec-1', [digest('p1', '2026-08-19T11:00:00Z')]]]),
    },
  );

  const first = await runFullSweep(h.deps, BUDGET);
  assert.equal(first.graphRequests, 3, 'tree, enumeration, and the one content fetch');

  const second = await runFullSweep(h.deps, BUDGET);
  assert.equal(second.graphRequests, 2, 'the second run fetches no page at all');
  assert.equal(h.storeCalls.metadata.length, 1);
});

test('a re-fetch whose content changed writes the page', async () => {
  const h = harness(
    { summaries: { 'sec-1': [summary('p1', '2026-08-19T13:00:00Z')] } },
    {
      sections: [section()],
      pages: new Map([
        ['p1', storedPage({ lastModifiedDateTime: '2026-08-19T11:00:00Z', contentHash: 'old' })],
      ]),
      digestsBySection: new Map([['sec-1', [digest('p1', '2026-08-19T11:00:00Z')]]]),
    },
  );

  const report = await runFullSweep(h.deps, BUDGET);

  assert.deepEqual(h.storeCalls.puts.map((p) => p.page.id), ['p1']);
  assert.deepEqual(h.storeCalls.metadata, [], 'a full write is not also a metadata write');
  assert.equal(report.pagesUpdated, 1);
});

test('a sweep whose enumeration failed fetches nothing and deletes nothing', async () => {
  const h = harness(
    {
      summaries: {
        'sec-1': () => {
          throw graphError(500);
        },
      },
    },
    {
      sections: [section()],
      digestsBySection: new Map([['sec-1', [digest('p1', '2026-08-19T11:00:00Z')]]]),
    },
  );

  await runFullSweep(h.deps, BUDGET);

  assert.deepEqual(h.storeCalls.deletes, [], 'a failed enumeration deletes nothing');
  assert.deepEqual(h.storeCalls.puts, [], 'and re-fetches nothing either');
});

test('a Firestore refusal on one page does not cost the next section its sweep', async () => {
  // Not the write path's "a failed invalidation never fails a write" — there is no write
  // here to protect. The sweep is the only pass that reconciles deletions, and one
  // section's Firestore failure must not stop every section after it from getting one.
  const h = harness(
    {
      summaries: {
        'sec-1': [summary('p1', '2026-08-19T13:00:00Z')],
        'sec-2': [summary('p2', '2026-08-19T13:00:00Z')],
      },
    },
    {
      sections: [section({ id: 'sec-1' }), section({ id: 'sec-2', path: '2026 / Other' })],
      digestsBySection: new Map([
        ['sec-1', [digest('p1', '2026-08-19T11:00:00Z')]],
        ['sec-2', [digest('p2', '2026-08-19T11:00:00Z')]],
      ]),
    },
  );

  const passes = h.deps.store.putPage.bind(h.deps.store);
  h.deps.store.putPage = (page, content) =>
    page.id === 'p1' ? Promise.reject(new Error('firestore-down')) : passes(page, content);

  const report = await runFullSweep(h.deps, BUDGET);

  assert.deepEqual(h.graphCalls.pageSummaries, ['sec-1', 'sec-2'], 'both sections are swept');
  assert.deepEqual(h.storeCalls.puts.map((p) => p.page.id), ['p2']);
  assert.equal(report.pagesFailed, 1);
});

test('a sweep out of budget stops before it deletes anything else', async () => {
  // The reconciliation loop spends Firestore writes — two per deletion — and the first
  // sweep after a deploy is when they are most widespread. Without the check the loop runs
  // to the end of the section whatever the budget says.
  const h = harness(
    { summaries: { 'sec-1': [] } },
    {
      sections: [section()],
      digestsBySection: new Map([['sec-1', [digest('p-gone')]]]),
    },
  );

  // tree + sec-1's enumeration = 2, leaving nothing for the reconciliation.
  const report = await runSweep(h.deps, { requestBudget: 2 });

  assert.deepEqual(h.storeCalls.deletes, []);
  assert.equal(report.done, false);
  assert.equal(report.outcome, 'budget-exhausted');
  assert.deepEqual(h.storeCalls.sweepResults, [], 'nor is the section recorded as swept');
});

test('a sweep out of budget stops before it re-fetches a disagreeing page', async () => {
  // The check matters more than it did. The reconciliation loop used to spend Firestore
  // writes alone; a disagreement now costs a Graph request, so a loop that ignored the
  // budget could spend a content fetch per mirrored page in the section — against 400 an
  // hour, shared with every interactive tool call.
  const h = harness(
    { summaries: { 'sec-1': [summary('p1', '2026-08-19T13:00:00Z')] } },
    {
      sections: [section()],
      digestsBySection: new Map([['sec-1', [digest('p1', '2026-08-19T11:00:00Z')]]]),
    },
  );

  // tree + sec-1's enumeration = 2, leaving nothing for the re-fetch.
  const report = await runSweep(h.deps, { requestBudget: 2 });

  assert.deepEqual(h.storeCalls.puts, [], 'the disagreeing page was not fetched');
  assert.equal(report.done, false, 'and the run says it left work behind');
  assert.equal(report.outcome, 'budget-exhausted');
  assert.deepEqual(h.storeCalls.sweepResults, [], 'nor is the section recorded as swept');
});

test('a sweep out of budget inside the last section still reports work outstanding', async () => {
  // The exhausted section being the *last* one is what makes this observable: `sweepPass`
  // never runs its loop-top check again, so it clears `sweepCursorSectionId` and the run
  // would report `complete` and `done: true` while pages it discovered were never fetched.
  const h = harness({ summaries: { 'sec-1': [summary('p-new')] } }, { sections: [section()] });

  const report = await runSweep(h.deps, { requestBudget: 2 });

  assert.deepEqual(h.storeCalls.puts, [], 'the discovered page was not fetched');
  assert.equal(report.done, false);
  assert.equal(report.outcome, 'budget-exhausted');
});

test('a page the old sweep stored unstamped is re-fetched by the same branch', async () => {
  // Before 2026-08-21 a discovered page was written with `title: ''` and the epoch. Both
  // reach the model, and a stale mark repaired neither: it dropped the content document
  // and left the wrong title answering `list_pages`. There is no separate predicate for
  // these any more — the epoch disagrees with anything Graph sends.
  const h = harness(
    { summaries: { 'sec-1': [summary('p1', '2026-08-19T13:00:00Z')] } },
    {
      sections: [section()],
      digestsBySection: new Map([
        ['sec-1', [{ ...digest('p1'), title: '', lastModifiedDateTime: '1970-01-01T00:00:00.000Z' }]],
      ]),
    },
  );

  const report = await runFullSweep(h.deps, BUDGET);

  const written = h.storeCalls.puts.find((put) => put.page.id === 'p1');
  assert.ok(written, 'the repair is a re-fetch');
  assert.equal(written.page.title, 'Page p1');
  assert.equal(written.page.lastModifiedDateTime, '2026-08-19T13:00:00Z');
  assert.equal(report.pagesUpdated, 1);
});

test('the sweep does not advance a section watermark', async () => {
  // Anything it could not fetch inside the budget has to be picked up by the next
  // incremental, which only happens if the watermark stays where it is.
  const h = harness({ tree: TREE, summaries: { 'sec-1': [summary('p-new')] } });

  await runSweep(h.deps, BUDGET);

  assert.deepEqual(h.storeCalls.watermarks, []);
  assert.deepEqual(h.storeCalls.sweepResults, ['sec-1']);
});

test('a budget-exhausted sweep records where to resume', async () => {
  const h = harness({ tree: TREE, summaries: { 'sec-1': [], 'sec-2': [] } });
  h.data.sections = [section({ id: 'sec-1' }), section({ id: 'sec-2', path: '2026 / Other' })];

  // tree + sec-1's enumeration = 2, leaving nothing for sec-2.
  const report = await runSweep(h.deps, { requestBudget: 2 });

  assert.equal(report.done, false);
  assert.ok(h.storeCalls.patches.some((p) => p.sweepCursorSectionId === 'sec-2'));
});

test('a scoped sweep visits only moved sections; a full sweep visits all of them', async () => {
  // No scripted tree, so the harness derives one naming `still` at the same 2020 timestamp
  // the store holds. That is what makes the scoped assertion mean something: the value the
  // filter reads is the overlaid live one, and a section the tree does not name is a
  // section `putStructure` would have deleted, which cannot happen in production.
  const stale = section({ id: 'still', graphLastModifiedDateTime: '2020-01-01T00:00:00Z' });
  const state = { ...initialSyncState(), sectionsScannedThrough: '2026-08-19T11:00:00.000Z' };

  const scoped = harness({ summaries: { still: [] } }, { sections: [stale], state });
  await runSweep(scoped.deps, BUDGET);
  assert.deepEqual(scoped.graphCalls.pageSummaries, [], 'nothing moved, so nothing is swept');

  const full = harness({ summaries: { still: [] } }, { sections: [stale], state });
  await runFullSweep(full.deps, BUDGET);
  assert.deepEqual(full.graphCalls.pageSummaries, ['still'], 'the weekly backstop visits everything');
});

test('a scoped sweep visits a section whose timestamp moved only in the live tree', async () => {
  // The other half of the pair, and the one the overlay is for: the stored copy still says
  // 2020 because no structure rewrite has happened since, and only the tree read knows the
  // section moved. Without the overlay this section is skipped and a deletion made in the
  // OneNote client is never noticed.
  const stale = section({ id: 'still', graphLastModifiedDateTime: '2020-01-01T00:00:00Z' });
  const state = { ...initialSyncState(), sectionsScannedThrough: '2026-08-19T11:00:00.000Z' };

  const moved: ExpandedNotebook[] = [
    {
      id: NB,
      displayName: '2026',
      sections: [{ id: 'still', displayName: 'Daily', lastModifiedDateTime: '2026-08-19T11:30:00Z' }],
      sectionGroups: [],
    },
  ];

  // The stored hash matches the *shape*, which the timestamps are excluded from, so the
  // run reads the tree and writes nothing — exactly the steady state the bug lived in.
  const h = harness(
    { tree: moved, summaries: { still: [] } },
    {
      sections: [stale],
      state: { ...state, structureHash: structureHashOf(buildStructure(moved, sel([NB]))) },
    },
  );

  await runSweep(h.deps, BUDGET);

  assert.equal(h.storeCalls.structures, 0, 'no rewrite, so nothing refreshed the stored copy');
  assert.deepEqual(h.graphCalls.pageSummaries, ['still'], 'and the section is swept anyway');
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
        putPageMetadata: store.putPageMetadata,
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

test('the timestamp alone never forces a rewrite, but it is written back', async () => {
  // lastModifiedDateTime moves on every write, so comparing it in the short-circuit would
  // rewrite every page the watermark overlap re-read and defeat the point of it. The
  // stamp still has to be *corrected*, through a metadata write that touches no content:
  // `resyncPage` stores this process's clock, that string is printed in every tool result,
  // and the sweep re-fetches any page whose stored stamp disagrees with Graph's. Without
  // the correction the local value is permanent and every later sweep fetches the page
  // again.
  const first = harness({ tree: TREE, changed: { 'sec-1': [summary('p-1', '2026-08-19T11:00:00Z')] } });
  await runIncremental(first.deps, BUDGET);
  const stored = first.data.pages.get('p-1');
  assert.ok(stored);

  const later = harness(
    { tree: TREE, changed: { 'sec-1': [summary('p-1', '2026-08-19T11:59:00Z')] } },
    { pages: new Map([['p-1', stored]]) },
  );
  const report = await runIncremental(later.deps, BUDGET);

  assert.deepEqual(later.storeCalls.puts, [], 'no content write');
  assert.deepEqual(later.storeCalls.metadata, [
    { id: 'p-1', lastModifiedDateTime: '2026-08-19T11:59:00Z' },
  ]);
  assert.equal(report.pagesUpdated, 0, 'and correcting a stamp is not a page update');
});

test('a short-circuit whose stamp agrees and whose copy is settled writes nothing at all', async () => {
  // The metadata write is a Firestore write per page, and the watermark overlap re-reads
  // every page edited in the last hour on every run. Firing it unconditionally would spend
  // one on each of them for no change.
  //
  // Driven through `writePageFromRaw` rather than through `runIncremental`, because
  // `storedPageIsCurrent` skips a page in exactly this state before any fetch happens, so
  // the short-circuit is not reachable from the incremental pass with a stamp that agrees
  // *and* a settled copy. `resyncPage` reaches it — it passes no stored document and does
  // not run the pre-check — and so would any later caller of this function. The property
  // belongs to the function, so it is asserted on the function.
  const { store, calls } = fakeStore({
    pages: new Map([['p1', settledPage({ lastModifiedDateTime: '2026-08-19T11:30:00Z' })]]),
  });
  const { blobs, calls: blobCalls } = fakeBlobs();

  const written = await writePageFromRaw(
    { store, blobs },
    { sectionId: 'sec-1', notebookId: NB, sectionPath: '2026 / Daily' },
    summary('p1', '2026-08-19T11:30:00Z'),
    rawHtml(TYPED_HTML),
  );

  assert.equal(written, false);
  assert.deepEqual(calls.puts, []);
  assert.deepEqual(calls.metadata, []);
  assert.deepEqual(blobCalls.ink, [], 'and no render, which is the expensive half');
});

test('a short-circuit whose copy the settle guard refuses writes metadata even so', async () => {
  // The two halves of the guard, at the level the branch lives on. Both leave the stored
  // copy unsettleable, and `putPageMetadata` is the only thing that refreshes
  // `contentSyncedAt`, so skipping the write here costs one Graph content request on every
  // later run — see the harness-driven pair above for the loop itself.
  const cases: [string, Partial<MirrorPage>][] = [
    ['a document written before contentSyncedAt existed', { contentSyncedAt: undefined }],
    ['a copy taken inside the settle window', { contentSyncedAt: '2026-08-19T11:30:05Z' }],
  ];

  for (const [why, overrides] of cases) {
    const { store, calls } = fakeStore({
      pages: new Map([
        ['p1', settledPage({ lastModifiedDateTime: '2026-08-19T11:30:00Z', ...overrides })],
      ]),
    });
    const { blobs } = fakeBlobs();

    const written = await writePageFromRaw(
      { store, blobs },
      { sectionId: 'sec-1', notebookId: NB, sectionPath: '2026 / Daily' },
      summary('p1', '2026-08-19T11:30:00Z'),
      rawHtml(TYPED_HTML),
    );

    assert.equal(written, false, `${why}: the content is still identical`);
    assert.deepEqual(calls.puts, [], why);
    assert.deepEqual(
      calls.metadata,
      [{ id: 'p1', lastModifiedDateTime: '2026-08-19T11:30:00Z' }],
      why,
    );
  }
});

test('an empty live stamp never overwrites a stored one', async () => {
  // `''` is `toPageSummary`'s fallback for a field Graph did not send, not a timestamp.
  // Storing it would put an empty string in every tool result for that page and leave the
  // real value unrecoverable without a full re-read.
  const h = harness(
    { tree: TREE, changed: { 'sec-1': [summary('p1', '')] } },
    { pages: new Map([['p1', storedPage({ lastModifiedDateTime: '2026-08-19T11:30:00Z' })]]) },
  );

  await runIncremental(h.deps, BUDGET);

  assert.deepEqual(h.storeCalls.metadata, []);
  assert.equal(h.data.pages.get('p1')?.lastModifiedDateTime, '2026-08-19T11:30:00Z');
});
