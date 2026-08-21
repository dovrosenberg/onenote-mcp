// The sync: what fills the Firestore mirror, and what keeps it honest.
//
// Two modes, both bounded and both resumable.
//
// **Incremental** reads the expanded tree (1 request), reconciles structure, picks the
// sections whose `lastModifiedDateTime` moved, lists each one's changed pages (1 request
// each), and fetches content for each changed page (1 request each). On the account this
// was sized against — 5 notebooks, 40 sections, 2000 pages — a quiet day is 1 request and
// a busy one is about 11.
//
// **Sweep** enumerates each section's pages — id, title and stamp, for one request — and
// reconciles them against the mirror. It does two jobs. It is the only way a deletion is
// ever noticed: Graph has no /delta on any OneNote resource and no tombstone for a
// deleted page, and the account-wide page list that would enumerate everything cheaply is
// the banned one, error 20266. It is also the only thing that notices a page edited in
// the OneNote client that the incremental pass missed, and the only thing that finds a
// page *moved into* a section — measured 2026-08-21, a move keeps the page's stamp, so it
// is below every later watermark. Both come from one comparison: a stored stamp that is
// not the same string as Graph's sends the page back through `syncPage`, and the content
// hash decides whether anything is written.
//
// Everything below follows from four things measured against the real account:
//
// 1. A page create, edit **and delete** each move the parent section's
//    `lastModifiedDateTime`, and nothing else does (2026-08-19, api-overview.md). That is
//    what makes tier 1 work, and it is what lets the nightly sweep visit only the
//    sections that moved. The value compared is the one *this run's* tree read observed,
//    overlaid onto the stored sections before the filter: the stored copy only moves on a
//    structure rewrite, so filtering on it goes blind an hour after the last one.
// 2. `$expand` on `/notebooks` has multi-minute outages — twice observed on 2026-08-19 —
//    while un-expanded calls on the same collection answer 200 throughout. So a failed
//    tree read is not fatal: the run keeps the structure already in Firestore and carries
//    on with the page pass. Structure is the slowest-changing thing in the account, and
//    skipping a whole poll cycle over it would be the more expensive mistake.
// 3. Cloud Run throttles CPU outside a request and times a request out at 300s, and the
//    service runs --max-instances=1. So a run does a bounded slice inside one request and
//    returns; there is no background work and nothing is deferred past the response.
// 4. OneNote allows 400 requests an hour. A run that spent them all would break the
//    interactive tools for the rest of the hour, so the budget is a hard stop and
//    exhausting it is a normal outcome reported as `budget-exhausted`, not a failure.
//
// The watermark rule is the one thing here that must not be "simplified". It is stored
// **per section**, and it is the time the pass *started* rather than the newest page it
// saw. Per section because a budget-bounded run does not finish, and a global watermark
// advanced past sections it never visited loses every edit in them permanently.
// Pass-start because Graph's clock and this service's clock are not the same clock; the
// hour of overlap `overlapFrom` reaches back covers the difference, and a page it
// re-lists that has not changed costs nothing at all — `storedPageIsCurrent` skips it
// without a request. The *section* scan reaches back a shorter distance for the opposite
// reason: every section it surfaces costs a listing request. The two windows are
// `WATERMARK_OVERLAP_MS` and `SECTION_SCAN_OVERLAP_MS` in ./mirror-schema.ts.

import { createHash } from 'node:crypto';

import { GraphRequestError } from './graph-structure.ts';
import type {
  ContainerChildren,
  ContainerKind,
  ExpandedNotebook,
  PageSummary,
} from './graph-structure.ts';
import { fitInkToByteBudget } from './ink.ts';
import { logEvent } from './logging.ts';
import type { MirrorBlobWriter } from './mirror-blobs.ts';
import {
  groupIdentity,
  htmlPlacement,
  htmlObjectName,
  inkObjectName,
  isActive,
  notebooksNeedingWideScan,
  selectionMatchesSeen,
  inkmlObjectName,
  notebookIdentity,
  overlapFrom,
  overlapSaveAgeMs,
  contentCopyIsSettled,
  pageListingDiffers,
  sectionIdentity,
  storedPageIsCurrent,
  SECTION_SCAN_OVERLAP_MS,
  utf8Bytes,
  type MirrorPage,
  type MirrorPageContent,
  type MirrorSection,
  type MirrorSectionGroup,
  type MirrorSyncState,
  type MirrorTombstone,
  type NotebookSelection,
  type NotebookStructureWrite,
  type NotebookTreeFields,
  type PageStamp,
  type SectionGroupStructureWrite,
  type SectionGroupTreeFields,
  type SectionStructureWrite,
  type SectionTreeFields,
  type SyncMode,
  type SyncOutcome,
} from './mirror-schema.ts';
import { MirrorLeaseHeldError } from './mirror-store.ts';
import { pageHtml, renderPageInkWithSource, type RawPageContent } from './page-content.ts';

/**
 * Wall clock one run may spend, well inside Cloud Run's 300s request timeout.
 *
 * Checked before each page is fetched rather than only before each section, because the
 * thing most likely to overrun is 120 resvg renders on a 1-CPU instance rather than the
 * requests themselves.
 */
export const SYNC_TIME_BUDGET_MS = 240_000;

/** How many sections are listed at once. The shared gate caps this again at 4. */
export const SYNC_CONCURRENCY = 4;

export interface SyncGraph {
  getExpandedTree(): Promise<ExpandedNotebook[]>;
  listContainerChildren(kind: ContainerKind, containerId: string): Promise<ContainerChildren>;
  listPagesChangedSince(sectionId: string, sinceIso: string): Promise<PageSummary[]>;
  listPagesInSection(sectionId: string, top?: number): Promise<PageSummary[]>;
  listPageSummaries(sectionId: string): Promise<PageSummary[]>;
}

export interface SyncContent {
  fetchRaw(pageId: string): Promise<RawPageContent>;
}

/** The slice of MirrorStore the sync writes through. MirrorStore satisfies it. */
export interface SyncStore {
  getSelection(): Promise<NotebookSelection>;
  getSyncState(): Promise<MirrorSyncState>;
  patchSyncState(patch: Partial<MirrorSyncState>): Promise<void>;
  acquireLease(mode: SyncMode, nowIso: string): Promise<void>;
  releaseLease(heldSince: string): Promise<void>;
  putStructure(structure: {
    notebooks: readonly NotebookStructureWrite[];
    sectionGroups: readonly SectionGroupStructureWrite[];
    sections: readonly SectionStructureWrite[];
  }): Promise<void>;
  listSectionsToSync(): Promise<MirrorSection[]>;
  listAllSectionGroups(): Promise<MirrorSectionGroup[]>;
  setSectionWatermark(sectionId: string, watermarkIso: string): Promise<void>;
  setSectionSweepResult(
    sectionId: string,
    fields: { pageCount: number; lastSweptAt: string },
  ): Promise<void>;
  setChildGroupsKnown(groupId: string, known: boolean): Promise<void>;
  getPage(pageId: string): Promise<MirrorPage | null>;
  putPage(page: MirrorPage, content: MirrorPageContent | null): Promise<void>;
  putPageMetadata(page: PageMetadataWrite): Promise<void>;
  deletePage(tombstone: MirrorTombstone): Promise<void>;
  listPageDigestsInSection(sectionId: string): Promise<PageStamp[]>;
}

/** The fields a short-circuited page write corrects without touching stored content. */
export type PageMetadataWrite = Pick<
  MirrorPage,
  'id' | 'title' | 'titleLower' | 'sectionId' | 'notebookId' | 'sectionPath' | 'lastModifiedDateTime'
>;

export interface SyncDeps {
  readonly graph: SyncGraph;
  readonly content: SyncContent;
  readonly store: SyncStore;
  readonly blobs: MirrorBlobWriter;
  /** Injected so a test does not wait in real time. */
  readonly now?: () => number;
}

export interface SyncOptions {
  readonly requestBudget: number;
  readonly timeBudgetMs?: number;
}

export interface SyncReport {
  readonly mode: SyncMode;
  readonly outcome: SyncOutcome;
  /** False when the budget stopped the run with work outstanding. */
  readonly done: boolean;
  readonly graphRequests: number;
  readonly sectionsVisited: number;
  readonly pagesUpdated: number;
  readonly pagesDeleted: number;
  readonly pagesFailed: number;
  /**
   * Pages the listing named whose stored copy Graph already describes exactly.
   *
   * The saving, reported rather than hidden: each one is a content request not spent
   * against the hourly 400. A run whose `pagesSkipped` collapses to zero while
   * `sectionsVisited` holds steady is the signature of the skip having stopped working —
   * a changed field on the page document, or a `contentSyncedAt` that stopped being
   * written — which nothing else in the report would show.
   */
  readonly pagesSkipped: number;
  /** Selected notebook ids matching no notebook. A mistyped id is silent otherwise. */
  readonly unknownNotebookIds: number;
  /** Active notebook ids matching no notebook. */
  readonly unknownActiveNotebookIds: number;
  /** Sections this run declined to visit because their notebook is not active. */
  readonly sectionsSkippedInactive: number;
  /** False when the expanded-tree read failed and stored structure was used instead. */
  readonly treeRead: boolean;
  readonly durationMs: number;
}

/**
 * The two hard stops on one run.
 *
 * Both are checked before spending rather than after, so an overrun costs nothing. The
 * request budget is the one that matters — a run that spent the hourly 400 would break
 * every interactive tool for the rest of the hour.
 */
class SyncBudget {
  #spent = 0;
  readonly #requests: number;
  readonly #deadline: number;
  readonly #now: () => number;

  constructor(requests: number, timeBudgetMs: number, now: () => number) {
    this.#requests = requests;
    this.#deadline = now() + timeBudgetMs;
    this.#now = now;
  }

  get spent(): number {
    return this.#spent;
  }

  get exhausted(): boolean {
    return this.#spent >= this.#requests || this.#now() >= this.#deadline;
  }

  /** Take one request's worth of budget. Callers check `exhausted` first. */
  take(): void {
    this.#spent += 1;
  }
}

// ---------------------------------------------------------------------------
// Structure
// ---------------------------------------------------------------------------

export interface BuiltStructure {
  readonly notebooks: NotebookStructureWrite[];
  readonly sectionGroups: SectionGroupStructureWrite[];
  readonly sections: SectionStructureWrite[];
  /** Selected notebook ids that matched no notebook in the tree. */
  readonly unknownNotebookIds: string[];
  /**
   * Active notebook ids that name no *mirrored* notebook.
   *
   * Two ways to get there and both are silent otherwise: a typo, which matches nothing in
   * the tree, and a real notebook id that was never added to `notebookIds`, whose pages
   * are not mirrored so marking it active reaches nothing.
   */
  readonly unknownActiveNotebookIds: string[];
}

/**
 * Flatten the expanded tree into the three structure collections.
 *
 * Pure, and exported so it is asserted directly rather than through a fake store.
 *
 * **Every notebook and section in the account is stored, not just the selected ones.**
 * The tree read returns them all for the same one request, and `mirrored` records which
 * are in the selection. Without this, `list_notebooks` — which takes no arguments — and
 * an unscoped `search_pages` would answer confidently and partially from a mirror holding
 * 3 of 55 notebooks, and a partial answer that cannot be detected as partial is the
 * failure CLAUDE.md already names about truncated searches.
 *
 * **It emits tree-owned fields only.** No `pagesSyncedThrough`, no `pageCount`, no
 * `childGroupsKnown`, no `pendingWrites` — the tree read knows none of them, and a
 * structure write that carried a default for one would reset hours of Graph requests.
 * `NEW_SECTION_DEFAULTS` and `NEW_GROUP_DEFAULTS` in ./mirror-schema.ts are what a
 * document gets the first time it appears, applied by the store on a create and never on
 * an update. `childGroupsKnown: false` is the creation default for the reason its own
 * docstring gives: `$expand` reaches one level of section group, so a new group's nested
 * groups are absent from the response rather than known to be empty, and the sweep sets it
 * true.
 *
 * Each document carries an `identity` built from those tree fields alone, which is what
 * lets the store skip one that did not change.
 */
export function buildStructure(
  tree: readonly ExpandedNotebook[],
  selection: NotebookSelection,
): BuiltStructure {
  const selected = new Set(selection.notebookIds);
  const seen = new Set<string>();

  const notebooks: NotebookStructureWrite[] = [];
  const sectionGroups: SectionGroupStructureWrite[] = [];
  const sections: SectionStructureWrite[] = [];

  for (const notebook of tree) {
    seen.add(notebook.id);
    const mirrored = selected.has(notebook.id);

    notebooks.push(withNotebookIdentity({
      id: notebook.id,
      displayName: notebook.displayName,
      mirrored,
      sectionCount: notebook.sections.length,
      sectionGroupCount: notebook.sectionGroups.length,
      graphLastModifiedDateTime: notebook.lastModifiedDateTime ?? null,
    }));

    for (const section of notebook.sections) {
      sections.push(withSectionIdentity({
        id: section.id,
        displayName: section.displayName,
        notebookId: notebook.id,
        parentId: notebook.id,
        parentKind: 'notebook',
        path: `${notebook.displayName} / ${section.displayName}`,
        mirrored,
        graphLastModifiedDateTime: section.lastModifiedDateTime ?? null,
      }));
    }

    for (const group of notebook.sectionGroups) {
      const groupPath = `${notebook.displayName} / ${group.displayName}`;
      sectionGroups.push(withGroupIdentity({
        id: group.id,
        displayName: group.displayName,
        notebookId: notebook.id,
        parentId: notebook.id,
        parentKind: 'notebook',
        mirrored,
        path: groupPath,
      }));

      for (const section of group.sections) {
        sections.push(withSectionIdentity({
          id: section.id,
          displayName: section.displayName,
          notebookId: notebook.id,
          parentId: group.id,
          parentKind: 'sectionGroup',
          path: `${groupPath} / ${section.displayName}`,
          mirrored,
          graphLastModifiedDateTime: section.lastModifiedDateTime ?? null,
        }));
      }
    }
  }

  const mirroredIds = new Set(notebooks.filter((n) => n.mirrored).map((n) => n.id));

  return {
    notebooks,
    sectionGroups,
    sections,
    unknownNotebookIds: selection.notebookIds.filter((id) => !seen.has(id)),
    unknownActiveNotebookIds: (selection.activeNotebookIds ?? []).filter(
      (id) => !mirroredIds.has(id),
    ),
  };
}

// The three wrappers below exist so a field added to a tree-field type is a type error
// here rather than a field silently missing from the identity — the object literal is
// checked against the parameter type, and the identity is computed from the same value.

function withNotebookIdentity(fields: NotebookTreeFields): NotebookStructureWrite {
  return { ...fields, identity: notebookIdentity(fields) };
}

function withGroupIdentity(fields: SectionGroupTreeFields): SectionGroupStructureWrite {
  return { ...fields, identity: groupIdentity(fields) };
}

function withSectionIdentity(fields: SectionTreeFields): SectionStructureWrite {
  return { ...fields, identity: sectionIdentity(fields) };
}

/**
 * A hash of everything about the tree the mirror stores.
 *
 * An unchanged hash skips `putStructure` entirely, so a run against an unchanged tree
 * makes no Firestore query against the structure collections at all. The account has 55
 * notebooks and 568 sections, and the tree changes when someone adds or renames one; a run
 * every fifteen minutes would otherwise read all three collections each time.
 *
 * It is built from the same `identity` strings the documents themselves carry, so "the
 * hash moved" and "this document changed" cannot disagree. The two answer different
 * questions: the hash is the cheap check that decides whether to call `putStructure` at
 * all, and the per-document identity decides what that call writes.
 *
 * That makes the hash **wider** than the one it replaced, which covered `id`,
 * `displayName`, `parentId` and `mirrored` alone. `path`, `parentKind` and a notebook's
 * two counts now move it too. More edits therefore trigger a structure pass — and that is
 * a fix rather than a cost, because a pass now writes only the documents whose identity
 * moved. Under the old wholesale write, widening the hash would have meant re-backfilling
 * the selection over a moved section; a section renamed inside a group changes the `path`
 * of nothing but itself.
 *
 * The timestamps are deliberately **excluded**, because they move constantly and would
 * rewrite every document on every run. `reconcileStructure` therefore returns them
 * separately, and `withLiveMtimes` overlays them onto the stored sections before
 * `pickCandidates` reads them — without that overlay the stored copies freeze here and the
 * sync goes blind.
 */
export function structureHashOf(built: BuiltStructure): string {
  const hash = createHash('sha256');
  for (const notebook of built.notebooks) hash.update(`${notebook.identity}\n`);
  for (const group of built.sectionGroups) hash.update(`${group.identity}\n`);
  for (const section of built.sections) hash.update(`${section.identity}\n`);
  return hash.digest('hex');
}

// ---------------------------------------------------------------------------
// Incremental
// ---------------------------------------------------------------------------

export async function runIncremental(
  deps: SyncDeps,
  options: SyncOptions,
): Promise<SyncReport> {
  return runMode('incremental', deps, options, incrementalPass);
}

export async function runSweep(deps: SyncDeps, options: SyncOptions): Promise<SyncReport> {
  return runMode('sweep', deps, options, (ctx) => sweepPass(ctx, false, true));
}

export async function runFullSweep(deps: SyncDeps, options: SyncOptions): Promise<SyncReport> {
  return runMode('sweep-full', deps, options, (ctx) => sweepPass(ctx, true, true));
}

/**
 * Every section of every mirrored notebook, active or not.
 *
 * The one mode with no activity filter. An inactive notebook is not re-checked by any
 * other pass, so this is the only thing that notices a page deleted in the OneNote client
 * from a notebook the operator froze. Run it by hand, or rarely.
 */
export async function runSweepAll(deps: SyncDeps, options: SyncOptions): Promise<SyncReport> {
  return runMode('sweep-all', deps, options, (ctx) => sweepPass(ctx, true, false));
}

interface PassContext {
  readonly deps: SyncDeps;
  readonly budget: SyncBudget;
  readonly startedAtIso: string;
  readonly state: MirrorSyncState;
  readonly selection: NotebookSelection;
  readonly tally: Tally;
}

/**
 * What one run accumulates: counts, and the two facts that change mid-run.
 *
 * `filterSupported` lives here rather than being re-read from `ctx.state`, which is a
 * snapshot taken once at the start. Without it, flipping the stored flag on the first
 * 400 would not change this run's behaviour and every remaining section would retry the
 * same failing filter — one wasted request each, on the run that has just discovered the
 * service does not support it.
 */
interface Tally {
  filterSupported: boolean;
  sectionsVisited: number;
  pagesUpdated: number;
  pagesDeleted: number;
  pagesFailed: number;
  pagesSkipped: number;
  unknownNotebookIds: number;
  unknownActiveNotebookIds: number;
  sectionsSkippedInactive: number;
  treeRead: boolean;
  done: boolean;
  /**
   * Did any candidate section's page listing fail?
   *
   * Separate from `done` because the two say different things in the report: `done` false
   * is what makes `runMode` answer `budget-exhausted`, and a 429 on one listing is not
   * that. What they share is the consequence — this run did not scan every candidate, so
   * `sectionsScannedThrough` must not advance past the section it missed.
   */
  sectionListingFailed: boolean;
}

/**
 * Take the lease, run the pass, release the lease, write the report.
 *
 * The lease is released in a `finally`, and `releaseLease` is itself written never to
 * throw. A run whose work is committed must not be reported as a failure just because
 * the lease could not be cleared — the scheduler would retry it, and the lease expires
 * on age anyway.
 */
async function runMode(
  mode: SyncMode,
  deps: SyncDeps,
  options: SyncOptions,
  pass: (ctx: PassContext) => Promise<void>,
): Promise<SyncReport> {
  const now = deps.now ?? Date.now;
  const startedAt = now();
  const startedAtIso = new Date(startedAt).toISOString();
  const budget = new SyncBudget(options.requestBudget, options.timeBudgetMs ?? SYNC_TIME_BUDGET_MS, now);

  const tally: Tally = {
    filterSupported: true,
    sectionsVisited: 0,
    pagesUpdated: 0,
    pagesDeleted: 0,
    pagesFailed: 0,
    pagesSkipped: 0,
    unknownNotebookIds: 0,
    unknownActiveNotebookIds: 0,
    sectionsSkippedInactive: 0,
    treeRead: false,
    done: true,
    sectionListingFailed: false,
  };

  // Outside the try: a held lease must not be followed by a release that clears the
  // other run's lease.
  await deps.store.acquireLease(mode, startedAtIso);

  let outcome: SyncOutcome = 'complete';
  try {
    const [state, selection] = await Promise.all([
      deps.store.getSyncState(),
      deps.store.getSelection(),
    ]);

    tally.filterSupported = state.datetimeFilterSupported;
    await pass({ deps, budget, startedAtIso, state, selection, tally });
    if (!tally.done) outcome = 'budget-exhausted';
  } catch (err) {
    outcome = 'failed';
    tally.done = false;
    logEvent('sync-failed', { mode, reason: reasonOf(err) });
    await deps.store.releaseLease(startedAtIso);
    throw err;
  } finally {
    if (outcome !== 'failed') await deps.store.releaseLease(startedAtIso);
  }

  const report: SyncReport = {
    mode,
    outcome,
    done: tally.done,
    graphRequests: budget.spent,
    sectionsVisited: tally.sectionsVisited,
    pagesUpdated: tally.pagesUpdated,
    pagesDeleted: tally.pagesDeleted,
    pagesFailed: tally.pagesFailed,
    pagesSkipped: tally.pagesSkipped,
    unknownNotebookIds: tally.unknownNotebookIds,
    unknownActiveNotebookIds: tally.unknownActiveNotebookIds,
    sectionsSkippedInactive: tally.sectionsSkippedInactive,
    treeRead: tally.treeRead,
    durationMs: now() - startedAt,
  };

  await deps.store.patchSyncState({
    lastRunOutcome: outcome,
    lastRunGraphRequests: report.graphRequests,
    lastRunPagesUpdated: report.pagesUpdated,
    lastRunPagesDeleted: report.pagesDeleted,
    unknownNotebookIds: report.unknownNotebookIds,
    unknownActiveNotebookIds: report.unknownActiveNotebookIds,
  });

  logEvent('sync-completed', {
    mode,
    outcome,
    graphRequests: report.graphRequests,
    sectionsVisited: report.sectionsVisited,
    pagesUpdated: report.pagesUpdated,
    pagesDeleted: report.pagesDeleted,
    pagesFailed: report.pagesFailed,
    pagesSkipped: report.pagesSkipped,
    sectionsSkippedInactive: report.sectionsSkippedInactive,
    durationMs: report.durationMs,
  });

  return report;
}

async function incrementalPass(ctx: PassContext): Promise<void> {
  if (ctx.selection.notebookIds.length === 0) {
    // Nothing selected is a real state, not an error: the selection document is
    // hand-edited and may legitimately be empty. It costs no Graph request to say so.
    return;
  }

  const structure = await reconcileStructure(ctx);
  const wideScanNotebookIds = await reconcileSelection(ctx, structure.mirroredNotebookIds);

  const sections = await ctx.deps.store.listSectionsToSync();
  const { eligible, skippedInactive } = splitByActivity(sections, ctx.selection, true);
  ctx.tally.sectionsSkippedInactive = skippedInactive;

  const candidates = pickCandidates(eligible, structure.liveMtimes, {
    state: ctx.state,
    mayFilterByTimestamp: ctx.tally.treeRead,
    wideScanNotebookIds,
  });

  for (const section of candidates) {
    if (ctx.budget.exhausted) {
      ctx.tally.done = false;
      return;
    }
    await syncSection(ctx, section);
  }

  // Only when every candidate completed. `sectionsScannedThrough` is what tier 1 compares
  // against, so advancing it after a partial run would make the next run skip the
  // sections this one never reached — whether the run stopped early (`done`) or carried
  // on past a section whose listing failed (`sectionListingFailed`).
  if (ctx.tally.done && ctx.tally.treeRead && !ctx.tally.sectionListingFailed) {
    await ctx.deps.store.patchSyncState({
      sectionsScannedThrough: ctx.startedAtIso,
      // Cleared on exactly the condition that advances the cutoff, and for the same
      // reason. A run stopped by its budget has not visited the sections it was widened
      // for, and clearing the set there would put them back below a cutoff that only
      // moves forward.
      wideScanNotebookIds: [],
      backfillComplete: sections.every((section) => section.pagesSyncedThrough !== null),
    });
  }
}

/** Section `lastModifiedDateTime` as one tree read saw it, by section id. */
export type SectionMtimes = ReadonlyMap<string, string | null>;

/**
 * What one structure pass learned.
 *
 * Empty timestamps mean no tree read happened — the budget was exhausted, or the read
 * failed.
 */
interface StructureResult {
  readonly liveMtimes: SectionMtimes;
  /**
   * The selected notebooks the tree returned, which is not `selection.notebookIds`: an id
   * in the selection that no notebook in the tree carries is an `unknownNotebookIds` entry
   * and is absent here.
   */
  readonly mirroredNotebookIds: readonly string[];
}

/**
 * Read the tree and write the structure, or carry on without it.
 *
 * A failed tree read is logged and survived: `$expand` on `/notebooks` was measured
 * unavailable for minutes at a time on 2026-08-19 while un-expanded calls answered 200,
 * and refusing to sync pages because the structure read failed would skip a whole poll
 * cycle over the slowest-changing thing in the account.
 */
async function reconcileStructure(ctx: PassContext): Promise<StructureResult> {
  const none: StructureResult = {
    liveMtimes: new Map(),
    mirroredNotebookIds: [],
  };

  if (ctx.budget.exhausted) {
    ctx.tally.done = false;
    return none;
  }

  ctx.budget.take();
  let tree: ExpandedNotebook[];
  try {
    tree = await ctx.deps.graph.getExpandedTree();
  } catch (err) {
    if (!(err instanceof GraphRequestError)) throw err;
    logEvent('sync-tree-failed', { status: err.status, reason: reasonOf(err) });
    await ctx.deps.store.patchSyncState({ lastTreeFailureAt: ctx.startedAtIso });
    return none;
  }

  ctx.tally.treeRead = true;

  const built = buildStructure(tree, ctx.selection);

  // Taken whether or not the hash moved. This is the whole fix: `structureHashOf`
  // excludes timestamps on purpose, so when the hash matches nothing writes them and the
  // stored copies stay at whatever the last structure rewrite recorded.
  const liveMtimes: SectionMtimes = new Map(
    built.sections.map((section) => [section.id, section.graphLastModifiedDateTime]),
  );
  const mirroredNotebookIds = built.notebooks.filter((n) => n.mirrored).map((n) => n.id);

  ctx.tally.unknownNotebookIds = built.unknownNotebookIds.length;
  if (built.unknownNotebookIds.length > 0) {
    // Ids only in the count, never the ids themselves — a notebook id is opaque, but the
    // count is all an operator needs to know their selection has a typo in it.
    logEvent('mirror-selection-unknown', { count: built.unknownNotebookIds.length });
  }

  ctx.tally.unknownActiveNotebookIds = built.unknownActiveNotebookIds.length;
  if (built.unknownActiveNotebookIds.length > 0) {
    logEvent('mirror-selection-unknown-active', { count: built.unknownActiveNotebookIds.length });
  }

  const hash = structureHashOf(built);
  if (hash === ctx.state.structureHash) {
    return { liveMtimes, mirroredNotebookIds };
  }

  await ctx.deps.store.putStructure(built);
  await ctx.deps.store.patchSyncState({ structureHash: hash });
  return { liveMtimes, mirroredNotebookIds };
}

/**
 * Record the selection lists, and name the notebooks a change to them has to widen.
 *
 * Returns the notebooks whose sections bypass tier 1 of `pickCandidates` this run: the
 * set carried in the state document, plus whatever this run's diff added. Without it a
 * notebook just mirrored or just activated would never be re-checked — tier 1 skips a
 * section whose `graphLastModifiedDateTime` is older than
 * `overlapFrom(state.sectionsScannedThrough, SECTION_SCAN_OVERLAP_MS)`, and that cutoff
 * advances on every completed run, so a section last edited three months ago while its
 * notebook was frozen is older than the cutoff for ever.
 *
 * `sectionsScannedThrough` is deliberately not touched. Nulling it is what this replaced,
 * and it is a global value: one activation made every mirrored active section a
 * candidate, which is one `listPagesChangedSince` per section of the whole selection —
 * about 70 requests on this account, against an hourly budget of 400, for a change that
 * concerned one notebook.
 *
 * Only candidacy is widened. Each named notebook's sections still list against their own
 * `pagesSyncedThrough`, which nothing here touches, so no page is re-fetched that has not
 * changed.
 *
 * A state document written before these fields existed carries
 * `mirroredNotebookIdsSeen: null`, and both schema functions read that as "never
 * recorded": `notebooksNeedingWideScan` answers `[]` and `selectionMatchesSeen` answers
 * false, so the first run after the deploy records the lists and widens nothing. That is
 * the general path rather than a branch of its own.
 *
 * A run that read no tree records nothing, and that is not tidiness. `mirroredNotebookIds`
 * is empty on such a run — the budget ran out before the read, or `getExpandedTree`
 * answered the 500 measured on 2026-08-19 — so the `activeNotebookIds === null` case would
 * resolve "every mirrored notebook is active" to the empty set, find nothing newly active,
 * and still write `activeNotebookIdsSeen: null`. The next healthy run would then diff
 * `null` against `null`, find nothing, and the notebook the operator just unfroze would
 * stay below the cutoff for ever. Recording is never urgent; the next run with a tree
 * diffs correctly.
 */
async function reconcileSelection(
  ctx: PassContext,
  mirroredNotebookIds: readonly string[],
): Promise<ReadonlySet<string>> {
  const carried = new Set(ctx.state.wideScanNotebookIds);
  if (!ctx.tally.treeRead) return carried;

  // Recorded from the selection document verbatim, while the `null` active list resolves
  // through the tree. The two domains differ: an id naming no notebook is in the first and
  // not the second, so a diff against it can only ever under-report, never widen something
  // that does not exist.
  const recorded = {
    mirroredNotebookIdsSeen: [...ctx.selection.notebookIds],
    activeNotebookIdsSeen:
      ctx.selection.activeNotebookIds === null ? null : [...ctx.selection.activeNotebookIds],
  };

  const seen = {
    mirrored: ctx.state.mirroredNotebookIdsSeen,
    active: ctx.state.activeNotebookIdsSeen,
  };
  const added = notebooksNeedingWideScan(seen, ctx.selection, mirroredNotebookIds);
  // Both halves, because a removal widens nothing and still has to be recorded: a run that
  // wrote only when it had something to widen would leave the deactivated notebook in
  // `activeNotebookIdsSeen`, and re-activating it later would diff against a list it is
  // already in. On a steady account this is the branch every run takes, and it costs no
  // write.
  if (added.length === 0 && selectionMatchesSeen(seen, ctx.selection)) return carried;

  const widened = new Set([...carried, ...added]);
  await ctx.deps.store.patchSyncState({ ...recorded, wideScanNotebookIds: [...widened] });
  if (added.length > 0) {
    // A count rather than the ids: a notebook id is opaque and this line reaches a public
    // log, and the count is what says whether one edit widened one notebook or fifty. A
    // removal reaches here to be recorded and logs nothing, because it widened nothing.
    logEvent('mirror-selection-widened', { count: added.length });
  }
  return widened;
}

/**
 * Split sections into the ones this run may visit and a count of what it declined.
 *
 * A separate exported function rather than a parameter on `pickCandidates`, for two
 * reasons. The count of what was declined has to reach the report, and a filter folded
 * into `pickCandidates` would have to return it alongside the list. And `pickCandidates`
 * has an early return for `!state.sectionRollUpTrusted || !timestampsAreFresh` that a
 * folded-in filter would have to be applied on both sides of — the kind of duplication
 * that ends with one side wrong.
 *
 * `includeBackfill` is what makes an inactive notebook fill up exactly once: a section
 * that has never been synced is eligible whatever its notebook's activity, and from then
 * on no incremental run lists it again. A sweep never backfills, so it passes false.
 */
export function splitByActivity(
  sections: readonly MirrorSection[],
  selection: NotebookSelection,
  includeBackfill: boolean,
): { eligible: MirrorSection[]; skippedInactive: number } {
  const eligible: MirrorSection[] = [];
  let skippedInactive = 0;

  for (const section of sections) {
    if (isActive(selection, section.notebookId)) {
      eligible.push(section);
    } else if (includeBackfill && section.pagesSyncedThrough === null) {
      eligible.push(section);
    } else {
      skippedInactive += 1;
    }
  }

  return { eligible, skippedInactive };
}

/**
 * Stored sections carrying this run's observed timestamps.
 *
 * A section absent from `live` keeps its stored value rather than losing it. That case is
 * only reachable when no tree read happened — the budget was exhausted, or the read
 * failed — and `mayFilterByTimestamp` is false then, so `pickCandidates` returns
 * everything regardless.
 *
 * A live entry of `null` is Graph reporting no timestamp on the section, and it overwrites
 * the stored value like any other observation. Falling back to the stored one there would
 * be the frozen-timestamp bug reached through a narrower door.
 */
export function withLiveMtimes(
  sections: readonly MirrorSection[],
  live: SectionMtimes,
): MirrorSection[] {
  return sections.map((section) => {
    const observed = live.get(section.id);
    return observed === undefined ? section : { ...section, graphLastModifiedDateTime: observed };
  });
}

/**
 * The three decisions `pickCandidates` makes its filter from, named at the call site.
 *
 * An options object rather than three more positional parameters: a `ReadonlySet<string>`
 * beside a `boolean` and a state object has no self-evident order. `sections` and `live`
 * stay positional because they are the data being filtered and the overlay applied to it,
 * and their types say which is which.
 */
export interface CandidateOptions {
  readonly state: MirrorSyncState;
  readonly mayFilterByTimestamp: boolean;
  /** Notebooks whose sections skip the timestamp cutoff; see `reconcileSelection`. */
  readonly wideScanNotebookIds: ReadonlySet<string>;
}

/**
 * The sections worth listing this run.
 *
 * The timestamps compared are `live` — what this run's tree read saw — overlaid onto the
 * stored sections here rather than by the caller, because a call site that forgot the
 * overlay would filter on stored values that only move when the structure is rewritten,
 * and the sync would go blind an hour after the last rewrite with nothing to say so.
 *
 * `mayFilterByTimestamp` means the timestamps compared are this run's rather than a stored
 * copy that has frozen, so both call sites pass `ctx.tally.treeRead` and nothing else.
 *
 * It used to be `treeRead && !structure.rewritten` in the incremental pass, and that was
 * the bug this replaced: `notebookIdentity` carries `mirrored`, so editing `notebookIds`
 * moves the structure hash, and a moved hash returned every observed section here — before
 * the wide-scan clause was reached at all. One notebook added to the selection therefore
 * listed every mirrored active section in the account, which is what the per-notebook set
 * exists to stop.
 *
 * Each reason a structure change used to need a wide pass is now covered by a narrower
 * rule. A section the tree just gained is created with `pagesSyncedThrough: null` from
 * `NEW_SECTION_DEFAULTS`, so the null-watermark clause below takes it. A section whose
 * notebook just became mirrored or active is the wide-scan set's job. A renamed section
 * keeps a watermark no structure write touches, and a rename changes no page.
 *
 * A section moved between notebooks brings no watermark with it. Measured 2026-08-21 and
 * recorded in `api-overview.md`, "Moving a section reissues its id across notebooks and
 * keeps it within one": the OneNote client reissues a section's Graph id when it changes
 * notebook, and keeps it when the section is only reparented under a section group in the
 * same notebook. So the old id is absent from the next tree read, its document is deleted
 * by absence, and the new id creates a document with `pagesSyncedThrough: null` — the
 * null-watermark clause below, which no cutoff and no activity filter can decline. The
 * class this docstring used to record as knowingly uncovered, a section carrying arrears
 * into a mirrored active notebook, is unreachable.
 *
 * The cost that replaces it: a cross-notebook move re-backfills that section in full, every
 * page fetched again under its new id and every page document under the old id deleted.
 * Correct and self-healing, and a section's worth of requests against the hourly budget.
 *
 * Two sections are always candidates whatever the clock says: one never synced
 * (`pagesSyncedThrough === null`), and one Graph reports no timestamp for — "the field is
 * absent" and "the roll-up cannot be trusted" have to be the same branch, or a service
 * that quietly stopped returning it would silently stop the mirror updating.
 */
export function pickCandidates(
  sections: readonly MirrorSection[],
  live: SectionMtimes,
  options: CandidateOptions,
): MirrorSection[] {
  const { state, mayFilterByTimestamp, wideScanNotebookIds } = options;
  const observed = withLiveMtimes(sections, live);
  if (!state.sectionRollUpTrusted || !mayFilterByTimestamp) return observed;

  const since = overlapFrom(state.sectionsScannedThrough, SECTION_SCAN_OVERLAP_MS);
  return observed.filter(
    (section) =>
      // First, because it is the only clause that can be true of a section every other
      // one declines. A notebook just mirrored or just activated holds sections whose
      // timestamps predate the cutoff by months, and the cutoff only advances, so nothing
      // else would ever make them candidates again.
      wideScanNotebookIds.has(section.notebookId) ||
      section.pagesSyncedThrough === null ||
      section.graphLastModifiedDateTime === null ||
      section.graphLastModifiedDateTime >= since,
  );
}

/**
 * One section: list what changed, fetch the pages the mirror is behind on, then advance
 * the watermark.
 *
 * Not every listed page is fetched. The listing carries each page's stamp and title, and
 * `storedPageIsCurrent` decides from those whether the stored copy already matches — which
 * is what the watermark overlap costs and what this saves, because the overlap re-lists
 * every page edited in the last hour on every run.
 *
 * The watermark moves only after the whole changed set landed. A failure part-way leaves
 * it where it was, so the next run retries the section rather than skipping the pages it
 * never reached.
 */
async function syncSection(ctx: PassContext, section: MirrorSection): Promise<void> {
  const since = overlapFrom(section.pagesSyncedThrough);

  let changed: PageSummary[];
  try {
    changed = await listChanged(ctx, section.id, since);
  } catch (err) {
    logEvent('sync-section-failed', { sectionId: section.id, reason: reasonOf(err) });
    ctx.tally.pagesFailed += 1;
    // Recorded, because otherwise the section-scan cutoff advances past a section this
    // run never listed. Tier 1 of `pickCandidates` then has `SECTION_SCAN_OVERLAP_MS` to
    // catch it, and a section whose Graph timestamp is older than that window drops out
    // of the candidate set for good — its watermark stuck behind its real edits until the
    // nightly full sweep. Holding the cutoff costs the next run a re-list of the sections
    // inside the window, which is the price of one transient 429.
    ctx.tally.sectionListingFailed = true;
    return;
  }

  ctx.tally.sectionsVisited += 1;

  for (const summary of changed) {
    // Checked before each page rather than only per section: 120 resvg renders on a
    // 1-CPU instance is what is most likely to overrun the wall clock. It is checked
    // ahead of the skip too, because the budget bounds wall clock as well as requests and
    // a section left half-visited has to report `done: false` so its watermark stays put.
    if (ctx.budget.exhausted) {
      ctx.tally.done = false;
      return;
    }

    // The saving this pass exists for. The listing already carried the stamp and the
    // title, so a page Graph describes exactly as the mirror already holds it needs no
    // content request — and the content request is the resource under the hourly 400,
    // where the write, the blob and the resvg render that `writePageFromRaw`'s
    // short-circuit already saves are not. With an hour of watermark overlap, every page
    // edited in the last hour was otherwise re-fetched by every run for the next hour.
    //
    // The document read here is handed to `writePageFromRaw`, so a page that is *not*
    // skipped is still read from Firestore exactly once.
    const stored = await ctx.deps.store.getPage(summary.id);
    if (stored !== null && storedPageIsCurrent(stored, summary, section.id)) {
      ctx.tally.pagesSkipped += 1;
      continue;
    }

    const written = await syncPage(ctx, section, summary, stored);

    // Only a page that was *written* counts. One the overlap re-read and found unchanged
    // is what the window costs, not what it saves, and logging it here would produce
    // evidence arguing to keep a window that is catching nothing.
    if (written) {
      const ageMs = overlapSaveAgeMs(section.pagesSyncedThrough, summary.lastModifiedDateTime);
      if (ageMs !== null) logEvent('sync-overlap-save', { ageMs });
    }
  }

  await ctx.deps.store.setSectionWatermark(section.id, ctx.startedAtIso);
}

/**
 * The changed-page list, with the unfiltered fallback.
 *
 * `$filter=lastModifiedDateTime ge …` was measured accepted on 2026-08-19. If the
 * service ever stops accepting it the answer is a 400, and the fallback costs the same
 * one request — the documented default sort for a section's pages is already
 * `lastModifiedTime desc`, confirmed in the same run, so a client-side cutoff on the
 * first page of results is equivalent. `datetimeFilterSupported` is flipped so every
 * later section in this and subsequent runs skips straight to the fallback.
 */
async function listChanged(
  ctx: PassContext,
  sectionId: string,
  since: string,
): Promise<PageSummary[]> {
  if (!ctx.tally.filterSupported) {
    ctx.budget.take();
    return cutOffAt(await ctx.deps.graph.listPagesInSection(sectionId, 100), since);
  }

  ctx.budget.take();
  try {
    return await ctx.deps.graph.listPagesChangedSince(sectionId, since);
  } catch (err) {
    if (!(err instanceof GraphRequestError) || err.status !== 400) throw err;

    logEvent('sync-filter-unsupported', { status: err.status });
    ctx.tally.filterSupported = false;
    await ctx.deps.store.patchSyncState({ datetimeFilterSupported: false });

    ctx.budget.take();
    return cutOffAt(await ctx.deps.graph.listPagesInSection(sectionId, 100), since);
  }
}

/** Pages at or after `since`, from a list Graph already sorted newest-first. */
function cutOffAt(pages: readonly PageSummary[], since: string): PageSummary[] {
  const kept: PageSummary[] = [];
  for (const page of pages) {
    if (page.lastModifiedDateTime < since) break;
    kept.push(page);
  }
  return kept;
}

/**
 * Fetch and store one page: raw HTML, and the ink rendered once.
 *
 * A 404 is not an error. The page was deleted between the listing and the fetch, which
 * is the issue's "lazy tombstoning" and comes free on this path.
 *
 * Returns true only when a page document was written. A deletion, a failure, and a copy
 * the short-circuit found already current are all false. `syncSection` reads it to decide
 * whether the watermark overlap earned its cost on this page.
 */
async function syncPage(
  ctx: PassContext,
  section: MirrorSection,
  summary: PageSummary,
  stored?: MirrorPage | null,
): Promise<boolean> {
  ctx.budget.take();

  let raw: RawPageContent;
  try {
    raw = await ctx.deps.content.fetchRaw(summary.id);
  } catch (err) {
    if (err instanceof GraphRequestError && err.status === 404) {
      await deletePage(ctx, summary.id, section, 'not-found');
      return false;
    }
    logEvent('mirror-page-failed', { pageId: summary.id, reason: reasonOf(err) });
    ctx.tally.pagesFailed += 1;
    return false;
  }

  try {
    // Only counted when something was actually written. A page re-read because it fell
    // inside the watermark overlap and found unchanged is not an update, and reporting
    // it as one would make every run look busy and hide a sync that had stopped working.
    if (!(await storePage(ctx, section, summary, raw, stored))) return false;
    ctx.tally.pagesUpdated += 1;
    return true;
  } catch (err) {
    logEvent('mirror-page-failed', { pageId: summary.id, reason: reasonOf(err) });
    ctx.tally.pagesFailed += 1;
    return false;
  }
}

/** True when the page was written; false when the stored copy was already current. */
async function storePage(
  ctx: PassContext,
  section: MirrorSection,
  summary: PageSummary,
  raw: RawPageContent,
  stored?: MirrorPage | null,
): Promise<boolean> {
  const written = await writePageFromRaw(
    { store: ctx.deps.store, blobs: ctx.deps.blobs },
    placementOf(section),
    summary,
    raw,
    () => {
      ctx.tally.pagesFailed += 1;
    },
    stored,
  );
  return written;
}

/** Where a page sits, denormalised onto its document so no query needs a join. */
export interface PagePlacement {
  readonly sectionId: string;
  readonly notebookId: string;
  readonly sectionPath: string;
}

function placementOf(section: MirrorSection): PagePlacement {
  return {
    sectionId: section.id,
    notebookId: section.notebookId,
    sectionPath: section.path,
  };
}

/** The narrow slice both the sync and a post-write resync write through. */
export interface PageWriteDeps {
  readonly store: Pick<SyncStore, 'getPage' | 'putPage' | 'putPageMetadata'>;
  readonly blobs: MirrorBlobWriter;
}

/**
 * Store one page from a raw content response. The one place a page document is built.
 *
 * Shared by the sync and by the resync a write tool triggers, deliberately: a second copy
 * that skipped the ink render, or spilled at a different threshold, would produce mirror
 * documents that differ depending on which path last touched them — and the difference
 * would only show up as a wrong answer to a model days later.
 *
 * `onInkFailure` exists because the sync counts a failed render into its report and a
 * resync has no report to count into. Neither treats it as fatal: most pages are typed,
 * and a page with unrenderable ink is still worth mirroring for its text.
 */
export async function writePageFromRaw(
  deps: PageWriteDeps,
  placement: PagePlacement,
  summary: PageSummary,
  raw: RawPageContent,
  onInkFailure: () => void = () => {},
  /**
   * The stored document, when the caller has already read it. `undefined` means "read it
   * here"; `null` means "there is none". `syncSection`'s pre-check has always just read
   * one, and a second read would double the Firestore cost of every page it does not skip.
   */
  known?: MirrorPage | null,
): Promise<boolean> {
  const html = pageHtml(raw) ?? '';
  const hash = createHash('sha256').update(html).digest('hex');

  // The short-circuit that makes the watermark overlap nearly free: a page re-read
  // because it fell inside the window, and unchanged, costs one request and no write.
  //
  // It compares the title and the section as well as the content hash, and both are
  // load-bearing rather than defensive. `update_page_title` changes a page's title and
  // nothing else, so a content-hash-only comparison short-circuited every rename and the
  // mirror kept serving the old title — which `find_page_by_name` and `search_pages` then
  // matched against. A page moved between sections is the same shape of miss.
  //
  // `lastModifiedDateTime` is deliberately NOT compared. It moves on every write, so
  // including it would rewrite every page the overlap re-read and defeat the whole point.
  const stored = known === undefined ? await deps.store.getPage(summary.id) : known;
  if (
    stored !== null &&
    stored.contentState === 'present' &&
    stored.contentHash === hash &&
    stored.title === summary.title &&
    stored.sectionId === placement.sectionId
  ) {
    // The content is right. Two other things may not be, and both have to be corrected
    // here or nowhere.
    //
    // The stamp: `resyncPage` writes this process's clock into `lastModifiedDateTime`,
    // nothing above compares that field, and the string is printed in every tool result.
    // The sweep re-fetches any page whose stored stamp disagrees with Graph's, so without
    // this write the local value is permanent and that page is fetched again on every
    // sweep, for ever.
    //
    // `contentSyncedAt`: the settle guard in `storedPageIsCurrent` refuses a copy that
    // cannot be shown to have been taken well after Graph's stamp, and `putPageMetadata`
    // is the only thing that refreshes that field on a page nothing rewrote. Without this
    // clause a refused copy was fetched, found identical, written nowhere, and refused
    // again on the next run — one Graph content request per run per page for the whole
    // hour-wide listing window, on the freshest pages in the account, invisible in a run
    // report because `pagesUpdated` and `pagesSkipped` both stay at zero. Every page
    // document written before `contentSyncedAt` existed has that shape.
    //
    // An empty live stamp is `toPageSummary`'s fallback for a field Graph did not send and
    // must not overwrite a good stored one. Such a page fails the settle guard for ever —
    // `Date.parse('')` is NaN — and is re-read once per run; that is the documented price
    // of never hiding an edit, and no page Graph stamps ever pays it.
    if (
      summary.lastModifiedDateTime !== '' &&
      (summary.lastModifiedDateTime !== stored.lastModifiedDateTime ||
        !contentCopyIsSettled(stored.contentSyncedAt, summary.lastModifiedDateTime))
    ) {
      await deps.store.putPageMetadata({
        id: summary.id,
        title: summary.title,
        titleLower: summary.title.toLowerCase(),
        sectionId: placement.sectionId,
        notebookId: placement.notebookId,
        sectionPath: placement.sectionPath,
        lastModifiedDateTime: summary.lastModifiedDateTime,
      });
    }
    return false;
  }

  const ink = renderInkOrNull(summary.id, raw, onInkFailure);
  const bytes = utf8Bytes(html);
  const location = htmlPlacement(bytes);

  if (location === 'gcs') {
    logEvent('mirror-html-spilled', { pageId: summary.id, bytes });
    await deps.blobs.putHtml(summary.id, html);
  }

  if (ink !== null) {
    await deps.blobs.putInk(summary.id, ink.image.png);
    await deps.blobs.putInkml(summary.id, ink.inkml);
  }

  const page: MirrorPage = {
    id: summary.id,
    title: summary.title,
    titleLower: summary.title.toLowerCase(),
    sectionId: placement.sectionId,
    notebookId: placement.notebookId,
    sectionPath: placement.sectionPath,
    lastModifiedDateTime: summary.lastModifiedDateTime,
    contentState: 'present',
    contentHash: hash,
    htmlLocation: location,
    htmlObject: location === 'gcs' ? htmlObjectName(summary.id) : null,
    htmlBytes: bytes,
    ink:
      ink === null
        ? null
        : {
            objectName: inkObjectName(summary.id),
            inkmlObjectName: inkmlObjectName(summary.id),
            width: ink.image.width,
            height: ink.image.height,
            strokeCount: ink.image.strokeCount,
            bytes: ink.image.png.byteLength,
          },
  };

  const content: MirrorPageContent | null =
    location === 'firestore' ? { pageId: summary.id, html, bytes, contentHash: hash } : null;

  await deps.store.putPage(page, content);
  return true;
}

/**
 * Render the ink, or answer null.
 *
 * A page that resvg or the InkML parser rejects must not fail the page. Most pages are
 * typed and have no ink at all; one that has unrenderable ink is still worth mirroring
 * for its text, and the alternative — failing the page — would mean the mirror silently
 * never holds it.
 */
function renderInkOrNull(
  pageId: string,
  raw: RawPageContent,
  onFailure: () => void,
): { image: ReturnType<typeof fitInkToByteBudget>; inkml: string } | null {
  try {
    const rendered = renderPageInkWithSource(raw);
    if (rendered === null) return null;
    return { image: fitInkToByteBudget(rendered.image), inkml: rendered.inkml };
  } catch (err) {
    logEvent('mirror-ink-failed', { pageId, reason: reasonOf(err) });
    onFailure();
    return null;
  }
}

async function deletePage(
  ctx: PassContext,
  pageId: string,
  section: { id: string; notebookId: string },
  reason: MirrorTombstone['reason'],
): Promise<void> {
  await ctx.deps.blobs.deleteForPage(pageId);
  await ctx.deps.store.deletePage({
    id: pageId,
    sectionId: section.id,
    notebookId: section.notebookId,
    reason,
  });
  ctx.tally.pagesDeleted += 1;
}

// ---------------------------------------------------------------------------
// Sweep
// ---------------------------------------------------------------------------

/**
 * Enumerate each section's pages and reconcile them against the mirror: deletions,
 * stamp disagreements, and discoveries.
 *
 * `unscoped` false sweeps only sections whose timestamp moved since the last sweep,
 * which the probe supports — a delete moves it. `unscoped` true sweeps every mirrored
 * section, and is the weekly backstop against the timestamp missing something.
 *
 * The sweep also queues pages Graph has that the mirror lacks. Measured 2026-08-21, a page
 * moved between sections keeps the `lastModifiedDateTime` it already had, so it sits below
 * every later section watermark and `listPagesChangedSince` never returns it. Without this
 * loop it would be invisible until someone next edited it.
 */
async function sweepPass(
  ctx: PassContext,
  unscoped: boolean,
  activeOnly: boolean,
): Promise<void> {
  if (ctx.selection.notebookIds.length === 0) return;

  const structure = await reconcileStructure(ctx);
  await learnNestedGroups(ctx);

  const stored = await ctx.deps.store.listSectionsToSync();

  // `includeBackfill` is false: a sweep reconciles a section it has already filled, and a
  // never-synced section in an inactive notebook is the incremental pass's job. `sweep-all` skips the filter entirely, which is what it is
  // for.
  let all = stored;
  if (activeOnly) {
    const split = splitByActivity(stored, ctx.selection, false);
    all = split.eligible;
    ctx.tally.sectionsSkippedInactive = split.skippedInactive;
  }

  const sections = unscoped
    ? all
    : pickCandidates(all, structure.liveMtimes, {
        state: ctx.state,
        mayFilterByTimestamp: ctx.tally.treeRead,
        // Read, never cleared: a sweep reconciles a section's pages against Graph rather
        // than resuming a watermark, so it is not what the widening is waiting for.
        wideScanNotebookIds: new Set(ctx.state.wideScanNotebookIds),
      });

  const resumeAt = ctx.state.sweepCursorSectionId;
  const start = resumeAt === null ? 0 : Math.max(0, sections.findIndex((s) => s.id === resumeAt));

  for (let index = start; index < sections.length; index += 1) {
    const section = sections[index] as MirrorSection;

    // Nothing started here, so the cursor names this section and the next run begins with
    // it.
    if (ctx.budget.exhausted) {
      ctx.tally.done = false;
      await ctx.deps.store.patchSyncState({ sweepCursorSectionId: section.id });
      return;
    }

    // Started and cut short part-way through its pages, so the cursor has to name *this*
    // section too. Writing the next one — which is what recording the cursor only at the
    // top of the loop did — leaves the interrupted section half reconciled, and nothing
    // returns to it until the cursor reaches the end of the list and resets, several runs
    // later. It is the section that most needed the visit: the budget ran out inside it
    // because it had the most to do. The first `/sync/sweep/full` after a deploy is when
    // that is likeliest, because every page `resyncPage` ever wrote carries a local
    // `…000Z` stamp that disagrees with Graph's spelling and so costs a content fetch.
    if (!(await sweepSection(ctx, section))) {
      await ctx.deps.store.patchSyncState({ sweepCursorSectionId: section.id });
      return;
    }
  }

  await ctx.deps.store.patchSyncState({ sweepCursorSectionId: null });
}

/**
 * Fill in the section groups `$expand` could not reach.
 *
 * `$expand` nests two levels, so a group inside a group is absent from the tree rather
 * than known to be empty. Until this runs, `childGroupsKnown` is false on every group and
 * the read path treats such a group as a mirror miss — which is correct but sends every
 * `list_sections` on it to Graph.
 */
async function learnNestedGroups(ctx: PassContext): Promise<void> {
  const groups = (await ctx.deps.store.listAllSectionGroups()).filter(
    (group) => group.mirrored && !group.childGroupsKnown,
  );

  for (const group of groups) {
    if (ctx.budget.exhausted) {
      ctx.tally.done = false;
      return;
    }
    ctx.budget.take();
    try {
      await ctx.deps.graph.listContainerChildren('sectionGroups', group.id);
      await ctx.deps.store.setChildGroupsKnown(group.id, true);
    } catch (err) {
      logEvent('sync-section-failed', { sectionId: group.id, reason: reasonOf(err) });
    }
  }
}

/**
 * One section's page reconciliation: deletions, stamp disagreements, and discoveries.
 *
 * **A failed enumeration deletes nothing and fetches nothing.** An auth failure or a 500
 * would otherwise empty the mirror one section at a time, and nothing about a deleted
 * page is recoverable from here. That is the single most important line in this file.
 *
 * Returns false only when the budget cut the section short part-way through its pages,
 * which is what `sweepPass` needs to write a resume cursor that names this section rather
 * than the next one. A failed enumeration returns true: it is logged, it reconciled
 * nothing, and the sweep carries on to the sections after it exactly as before.
 */
async function sweepSection(ctx: PassContext, section: MirrorSection): Promise<boolean> {
  ctx.budget.take();

  let live: PageSummary[];
  try {
    live = await ctx.deps.graph.listPageSummaries(section.id);
  } catch (err) {
    logEvent('sync-section-failed', { sectionId: section.id, reason: reasonOf(err) });
    ctx.tally.pagesFailed += 1;
    return true;
  }

  ctx.tally.sectionsVisited += 1;

  const liveById = new Map(live.map((page) => [page.id, page]));
  const mirrored = await ctx.deps.store.listPageDigestsInSection(section.id);

  for (const stored of mirrored) {
    // Bounded, like the discovery loop below it. A deletion is two Firestore writes the
    // request budget does not otherwise see, and a stamp disagreement now costs a Graph
    // request as well — so without this check one section could spend a content fetch per
    // mirrored page in it, against 400 an hour shared with every interactive tool call.
    // The first sweep after a deploy is when both are most widespread. An early return
    // skips `setSectionSweepResult`, so this section is not recorded as swept, and
    // `done: false` is what makes the run report `budget-exhausted` rather than a clean
    // pass.
    if (ctx.budget.exhausted) {
      ctx.tally.done = false;
      return false;
    }

    const match = liveById.get(stored.id);

    if (match === undefined) {
      await deletePage(ctx, stored.id, section, 'sweep');
      continue;
    }

    // Matched, so it is not a discovery. Removing it leaves `liveById` holding exactly
    // the pages Graph has that the mirror lacks, which is what the loop below wants.
    liveById.delete(stored.id);

    // The sweep's second job: the listing disagrees with the stored copy, so re-read the
    // page. `pageListingDiffers` compares the stamp *and* the title, and the title half is
    // what makes the sweep able to see a rename at all: measured 2026-08-21
    // (api-overview.md), a rename moves no `lastModifiedDateTime`, and the incremental
    // never lists a page whose stamp is below the section watermark — so a stamp-only
    // comparison here left a page renamed outside this server carrying its old title in the
    // mirror for ever, which is the field every listing and every by-name lookup matches
    // on. The reachable route needs nothing unmeasured: `update_page_title` marks the page
    // stale without touching its stamp, the PATCH renames it without moving the stamp, and
    // a `resyncPage` that hit a transient failure is documented as non-fatal.
    //
    // The disagreement is a hint and nothing more — `writePageFromRaw` compares the content
    // hash, the title and the section and writes only if something actually changed, and
    // corrects the stored stamp when nothing did, which is what stops the next sweep asking
    // again.
    //
    // **This must not become a stale mark.** `markPageStale` deletes the page-content
    // document, and nothing re-fetches a stale page: the incremental will not list a page
    // whose Graph stamp is behind the section watermark, no read path writes to the
    // mirror, and a mark is therefore permanent. `resyncPage` stamps this process's clock,
    // so a page written through this server is stored *ahead* of Graph — one sweep could
    // delete the mirrored content of every page the server has ever written. A re-fetch
    // repairs that case instead of destroying it, which is why the comparison needs no
    // direction test and no tolerance, and why no edit is too small for it to notice.
    //
    // It also subsumes the pre-2026-08-21 discovered pages, which carry `title: ''` and
    // `new Date(0)`: the epoch disagrees with anything Graph sends, and the re-fetch writes
    // both fields from Graph's listing.
    if (pageListingDiffers(stored, match)) await syncPage(ctx, section, match);
  }

  // Pages Graph has that the mirror lacks — new, or moved in. Fetched now, budget
  // permitting; anything left is picked up by the next incremental, because this
  // section's watermark is deliberately not advanced here.
  for (const page of liveById.values()) {
    if (ctx.budget.exhausted) {
      ctx.tally.done = false;
      return false;
    }
    await syncPage(ctx, section, page);
  }

  await ctx.deps.store.setSectionSweepResult(section.id, {
    pageCount: live.length,
    lastSweptAt: ctx.startedAtIso,
  });
  return true;
}

/** A reason string for a log line. Never a message, which can carry a request body. */
function reasonOf(err: unknown): string {
  if (err instanceof GraphRequestError) return `graph-${err.status}`;
  if (err instanceof MirrorLeaseHeldError) return 'lease-held';
  return err instanceof Error ? err.name : 'unknown';
}

// ---------------------------------------------------------------------------
// Resync after a write
// ---------------------------------------------------------------------------

/** What a write tool knows about the page it just changed, without re-reading metadata. */
export interface ResyncHint {
  /** Set by create_page and update_page_title; absent on an append, which cannot change it. */
  readonly title?: string;
  /** Set by the tools that name a section. Absent means "use the placement already stored". */
  readonly sectionId?: string;
}

export interface ResyncDeps extends PageWriteDeps {
  readonly content: SyncContent;
  readonly store: Pick<SyncStore, 'getPage' | 'putPage' | 'putPageMetadata' | 'deletePage'> & {
    getSection(sectionId: string): Promise<MirrorSection | null>;
  };
  readonly now?: () => number;
}

export type ResyncOutcome = 'updated' | 'unchanged' | 'not-mirrored' | 'deleted';

/**
 * Re-read one page from Graph and store it, right after a write changed it.
 *
 * One Graph request, on top of the write itself. What it buys is that a read immediately
 * after a write answers from the mirror with the new content, instead of falling through
 * to Graph until the next sync run — which on a fifteen-minute schedule is a long window
 * in the middle of a conversation.
 *
 * **It re-reads content and nothing else.** Measured 2026-08-19 and recorded in
 * `api-overview.md`: a PATCH is visible to the next content read, at 3.7 seconds
 * including both round trips — but page *metadata* is weaker, and
 * `GET /pages/{id}?$select=title` returned `""` for pages created seconds earlier. So the
 * title comes from the caller, which just set it or knows it did not change, and
 * `lastModifiedDateTime` is stamped locally. Reading either back from Graph here would
 * trade a correct value for an unreliable one.
 *
 * A page in a notebook the mirror does not hold is `not-mirrored` and costs no request:
 * the write tools reach the whole account, and only the selection is mirrored.
 */
export async function resyncPage(
  deps: ResyncDeps,
  pageId: string,
  hint: ResyncHint = {},
): Promise<ResyncOutcome> {
  const stored = await deps.store.getPage(pageId);

  const sectionId = hint.sectionId ?? stored?.sectionId;
  if (sectionId === undefined) return 'not-mirrored';

  const section = await deps.store.getSection(sectionId);
  if (section === null || !section.mirrored) return 'not-mirrored';

  let raw: RawPageContent;
  try {
    raw = await deps.content.fetchRaw(pageId);
  } catch (err) {
    if (err instanceof GraphRequestError && err.status === 404) {
      // Written and then deleted, or the id was wrong. Either way the mirror must not
      // keep a copy of a page that is gone.
      await deps.blobs.deleteForPage(pageId);
      await deps.store.deletePage({
        id: pageId,
        sectionId: section.id,
        notebookId: section.notebookId,
        reason: 'not-found',
      });
      return 'deleted';
    }
    throw err;
  }

  const now = deps.now ?? Date.now;
  const written = await writePageFromRaw(
    deps,
    placementOf(section),
    {
      id: pageId,
      title: hint.title ?? stored?.title ?? '',
      // Stamped locally rather than read back, for the reason above. It is within a
      // second or two of what Graph recorded, and the next sync corrects it either way.
      lastModifiedDateTime: new Date(now()).toISOString(),
    },
    raw,
  );

  return written ? 'updated' : 'unchanged';
}
