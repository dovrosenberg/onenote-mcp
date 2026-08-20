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
// **Sweep** enumerates page ids per section and reconciles them against the mirror, which
// is the only way a deletion is ever noticed: Graph has no /delta on any OneNote resource
// and no tombstone for a deleted page, and the account-wide page list that would
// enumerate everything cheaply is the banned one, error 20266.
//
// Everything below follows from four things measured against the real account:
//
// 1. A page create, edit **and delete** each move the parent section's
//    `lastModifiedDateTime`, and nothing else does (2026-08-19, api-overview.md). That is
//    what makes tier 1 work, and it is what lets the nightly sweep visit only the
//    sections that moved.
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
// hour of overlap in `overlapFrom` covers the difference, and a re-fetch of an unchanged
// page is nearly free because `contentHash` short-circuits it.

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
  htmlPlacement,
  htmlObjectName,
  inkObjectName,
  inkmlObjectName,
  overlapFrom,
  utf8Bytes,
  type MirrorNotebook,
  type MirrorPage,
  type MirrorPageContent,
  type MirrorSection,
  type MirrorSectionGroup,
  type MirrorSyncState,
  type MirrorTombstone,
  type NotebookSelection,
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
  listPageIds(sectionId: string): Promise<string[]>;
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
  releaseLease(): Promise<void>;
  putStructure(structure: {
    notebooks: readonly MirrorNotebook[];
    sectionGroups: readonly MirrorSectionGroup[];
    sections: readonly MirrorSection[];
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
  deletePage(tombstone: MirrorTombstone): Promise<void>;
  listPageIdsInSection(sectionId: string): Promise<string[]>;
}

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
  /** Selected notebook ids matching no notebook. A mistyped id is silent otherwise. */
  readonly unknownNotebookIds: number;
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
  readonly notebooks: MirrorNotebook[];
  readonly sectionGroups: MirrorSectionGroup[];
  readonly sections: MirrorSection[];
  /** Selected notebook ids that matched no notebook in the tree. */
  readonly unknownNotebookIds: string[];
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
 * `childGroupsKnown` is false on every group here, because `$expand` reaches one level of
 * section group and a first-level group's nested groups are *absent from the response*
 * rather than known to be empty. The sweep sets it true.
 */
export function buildStructure(
  tree: readonly ExpandedNotebook[],
  selection: NotebookSelection,
): BuiltStructure {
  const selected = new Set(selection.notebookIds);
  const seen = new Set<string>();

  const notebooks: MirrorNotebook[] = [];
  const sectionGroups: MirrorSectionGroup[] = [];
  const sections: MirrorSection[] = [];

  for (const notebook of tree) {
    seen.add(notebook.id);
    const mirrored = selected.has(notebook.id);

    notebooks.push({
      id: notebook.id,
      displayName: notebook.displayName,
      mirrored,
      sectionCount: notebook.sections.length,
      sectionGroupCount: notebook.sectionGroups.length,
      graphLastModifiedDateTime: notebook.lastModifiedDateTime ?? null,
    });

    for (const section of notebook.sections) {
      sections.push({
        id: section.id,
        displayName: section.displayName,
        notebookId: notebook.id,
        parentId: notebook.id,
        parentKind: 'notebook',
        path: `${notebook.displayName} / ${section.displayName}`,
        mirrored,
        graphLastModifiedDateTime: section.lastModifiedDateTime ?? null,
        pagesSyncedThrough: null,
        pageCount: 0,
      });
    }

    for (const group of notebook.sectionGroups) {
      const groupPath = `${notebook.displayName} / ${group.displayName}`;
      sectionGroups.push({
        id: group.id,
        displayName: group.displayName,
        notebookId: notebook.id,
        parentId: notebook.id,
        parentKind: 'notebook',
        mirrored,
        path: groupPath,
        childGroupsKnown: false,
      });

      for (const section of group.sections) {
        sections.push({
          id: section.id,
          displayName: section.displayName,
          notebookId: notebook.id,
          parentId: group.id,
          parentKind: 'sectionGroup',
          path: `${groupPath} / ${section.displayName}`,
          mirrored,
          graphLastModifiedDateTime: section.lastModifiedDateTime ?? null,
          pagesSyncedThrough: null,
          pageCount: 0,
        });
      }
    }
  }

  return {
    notebooks,
    sectionGroups,
    sections,
    unknownNotebookIds: selection.notebookIds.filter((id) => !seen.has(id)),
  };
}

/**
 * A hash of everything about the tree the mirror stores.
 *
 * An unchanged hash skips every structure write. The account has 55 notebooks and 568
 * sections; rewriting all of them every fifteen minutes would be millions of writes a
 * month for a tree that changes when someone adds a notebook. The timestamps are
 * deliberately **excluded** — they move constantly and are read from the live tree, not
 * from the stored copy, so including them would defeat the whole point.
 */
export function structureHashOf(built: BuiltStructure): string {
  const hash = createHash('sha256');
  for (const notebook of built.notebooks) {
    hash.update(`n ${notebook.id} ${notebook.displayName} ${notebook.mirrored}\n`);
  }
  for (const group of built.sectionGroups) {
    hash.update(`g ${group.id} ${group.displayName} ${group.parentId} ${group.mirrored}\n`);
  }
  for (const section of built.sections) {
    hash.update(`s ${section.id} ${section.displayName} ${section.parentId} ${section.mirrored}\n`);
  }
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
  return runMode('sweep', deps, options, (ctx) => sweepPass(ctx, false));
}

export async function runFullSweep(deps: SyncDeps, options: SyncOptions): Promise<SyncReport> {
  return runMode('sweep-full', deps, options, (ctx) => sweepPass(ctx, true));
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
  unknownNotebookIds: number;
  treeRead: boolean;
  done: boolean;
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
    unknownNotebookIds: 0,
    treeRead: false,
    done: true,
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
    await deps.store.releaseLease();
    throw err;
  } finally {
    if (outcome !== 'failed') await deps.store.releaseLease();
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
    unknownNotebookIds: tally.unknownNotebookIds,
    treeRead: tally.treeRead,
    durationMs: now() - startedAt,
  };

  await deps.store.patchSyncState({
    lastRunOutcome: outcome,
    lastRunGraphRequests: report.graphRequests,
    lastRunPagesUpdated: report.pagesUpdated,
    lastRunPagesDeleted: report.pagesDeleted,
    unknownNotebookIds: report.unknownNotebookIds,
  });

  logEvent('sync-completed', {
    mode,
    outcome,
    graphRequests: report.graphRequests,
    sectionsVisited: report.sectionsVisited,
    pagesUpdated: report.pagesUpdated,
    pagesDeleted: report.pagesDeleted,
    pagesFailed: report.pagesFailed,
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

  const structureChanged = await reconcileStructure(ctx);

  const sections = await ctx.deps.store.listSectionsToSync();
  const candidates = pickCandidates(sections, ctx.state, ctx.tally.treeRead && !structureChanged);

  for (const section of candidates) {
    if (ctx.budget.exhausted) {
      ctx.tally.done = false;
      return;
    }
    await syncSection(ctx, section);
  }

  // Only when every candidate completed. `sectionsScannedThrough` is what tier 1 compares
  // against, so advancing it after a partial run would make the next run skip the
  // sections this one never reached.
  if (ctx.tally.done && ctx.tally.treeRead) {
    await ctx.deps.store.patchSyncState({
      sectionsScannedThrough: ctx.startedAtIso,
      backfillComplete: sections.every((section) => section.pagesSyncedThrough !== null),
    });
  }
}

/**
 * Read the tree and write the structure, or carry on without it.
 *
 * Returns true when the structure was rewritten. A failed tree read is logged and
 * survived: `$expand` on `/notebooks` was measured unavailable for minutes at a time on
 * 2026-08-19 while un-expanded calls answered 200, and refusing to sync pages because
 * the structure read failed would skip a whole poll cycle over the slowest-changing
 * thing in the account.
 */
async function reconcileStructure(ctx: PassContext): Promise<boolean> {
  if (ctx.budget.exhausted) {
    ctx.tally.done = false;
    return false;
  }

  ctx.budget.take();
  let tree: ExpandedNotebook[];
  try {
    tree = await ctx.deps.graph.getExpandedTree();
  } catch (err) {
    if (!(err instanceof GraphRequestError)) throw err;
    logEvent('sync-tree-failed', { status: err.status, reason: reasonOf(err) });
    await ctx.deps.store.patchSyncState({ lastTreeFailureAt: ctx.startedAtIso });
    return false;
  }

  ctx.tally.treeRead = true;

  const built = buildStructure(tree, ctx.selection);
  ctx.tally.unknownNotebookIds = built.unknownNotebookIds.length;
  if (built.unknownNotebookIds.length > 0) {
    // Ids only in the count, never the ids themselves — a notebook id is opaque, but the
    // count is all an operator needs to know their selection has a typo in it.
    logEvent('mirror-selection-unknown', { count: built.unknownNotebookIds.length });
  }

  const hash = structureHashOf(built);
  if (hash === ctx.state.structureHash) return false;

  await ctx.deps.store.putStructure(built);
  await ctx.deps.store.patchSyncState({ structureHash: hash });
  return true;
}

/**
 * Which sections this run visits.
 *
 * With `sectionRollUpTrusted` — the default, on the strength of the 2026-08-19 probe —
 * only sections whose timestamp moved past the overlap window, plus any never synced.
 * Without it, or when the tree read failed so the stored timestamps are stale, every
 * mirrored section. `listSectionsToSync` already returns them oldest-watermark-first, so
 * a budget-bounded run round-robins rather than starving the tail.
 *
 * A section with no `graphLastModifiedDateTime` is always a candidate. "The field is
 * absent" and "the timestamp cannot be relied on" have to behave identically, or a
 * service that quietly stopped returning it would silently stop the mirror updating.
 */
export function pickCandidates(
  sections: readonly MirrorSection[],
  state: MirrorSyncState,
  timestampsAreFresh: boolean,
): MirrorSection[] {
  if (!state.sectionRollUpTrusted || !timestampsAreFresh) return [...sections];

  const since = overlapFrom(state.sectionsScannedThrough);
  return sections.filter(
    (section) =>
      section.pagesSyncedThrough === null ||
      section.graphLastModifiedDateTime === null ||
      section.graphLastModifiedDateTime >= since,
  );
}

/**
 * One section: list what changed, fetch each changed page, then advance the watermark.
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
    return;
  }

  ctx.tally.sectionsVisited += 1;

  for (const summary of changed) {
    // Checked before each page rather than only per section: 120 resvg renders on a
    // 1-CPU instance is what is most likely to overrun the wall clock.
    if (ctx.budget.exhausted) {
      ctx.tally.done = false;
      return;
    }
    await syncPage(ctx, section, summary);
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
 */
async function syncPage(
  ctx: PassContext,
  section: MirrorSection,
  summary: PageSummary,
): Promise<void> {
  ctx.budget.take();

  let raw: RawPageContent;
  try {
    raw = await ctx.deps.content.fetchRaw(summary.id);
  } catch (err) {
    if (err instanceof GraphRequestError && err.status === 404) {
      await deletePage(ctx, summary.id, section, 'not-found');
      return;
    }
    logEvent('mirror-page-failed', { pageId: summary.id, reason: reasonOf(err) });
    ctx.tally.pagesFailed += 1;
    return;
  }

  try {
    // Only counted when something was actually written. A page re-read because it fell
    // inside the watermark overlap and found unchanged is not an update, and reporting
    // it as one would make every run look busy and hide a sync that had stopped working.
    if (await storePage(ctx, section, summary, raw)) ctx.tally.pagesUpdated += 1;
  } catch (err) {
    logEvent('mirror-page-failed', { pageId: summary.id, reason: reasonOf(err) });
    ctx.tally.pagesFailed += 1;
  }
}

/** True when the page was written; false when the stored copy was already current. */
async function storePage(
  ctx: PassContext,
  section: MirrorSection,
  summary: PageSummary,
  raw: RawPageContent,
): Promise<boolean> {
  const written = await writePageFromRaw(
    { store: ctx.deps.store, blobs: ctx.deps.blobs },
    placementOf(section),
    summary,
    raw,
    () => {
      ctx.tally.pagesFailed += 1;
    },
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
  readonly store: Pick<SyncStore, 'getPage' | 'putPage'>;
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
): Promise<boolean> {
  const html = pageHtml(raw) ?? '';
  const hash = createHash('sha256').update(html).digest('hex');

  const stored = await deps.store.getPage(summary.id);
  if (stored !== null && stored.contentHash === hash && stored.contentState === 'present') {
    // The page's timestamp moved but its content did not — an ink stroke edited and
    // undone, a title change, or the watermark overlap re-reading a page already held.
    // Writing nothing here is what makes the overlap nearly free.
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
 * Enumerate page ids per section and reconcile them against the mirror.
 *
 * `unscoped` false sweeps only sections whose timestamp moved since the last sweep,
 * which the probe supports — a delete moves it. `unscoped` true sweeps every mirrored
 * section, and is the weekly backstop against the timestamp missing something.
 *
 * The sweep also queues ids Graph has that the mirror lacks. A page *moved into* a
 * mirrored section may not have its own `lastModifiedDateTime` bumped by the move — the
 * same class of unknown as the section roll-up — so without this it would be invisible
 * until someone next edited it.
 */
async function sweepPass(ctx: PassContext, unscoped: boolean): Promise<void> {
  if (ctx.selection.notebookIds.length === 0) return;

  await reconcileStructure(ctx);
  await learnNestedGroups(ctx);

  const all = await ctx.deps.store.listSectionsToSync();
  const sections = unscoped ? all : pickCandidates(all, ctx.state, ctx.tally.treeRead);

  const resumeAt = ctx.state.sweepCursorSectionId;
  const start = resumeAt === null ? 0 : Math.max(0, sections.findIndex((s) => s.id === resumeAt));

  for (let index = start; index < sections.length; index += 1) {
    const section = sections[index] as MirrorSection;
    if (ctx.budget.exhausted) {
      ctx.tally.done = false;
      await ctx.deps.store.patchSyncState({ sweepCursorSectionId: section.id });
      return;
    }
    await sweepSection(ctx, section);
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
 * One section's id reconciliation.
 *
 * **A failed enumeration deletes nothing.** An auth failure or a 500 would otherwise
 * empty the mirror one section at a time, and nothing about a deleted page is
 * recoverable from here. That is the single most important line in this file.
 */
async function sweepSection(ctx: PassContext, section: MirrorSection): Promise<void> {
  ctx.budget.take();

  let live: string[];
  try {
    live = await ctx.deps.graph.listPageIds(section.id);
  } catch (err) {
    logEvent('sync-section-failed', { sectionId: section.id, reason: reasonOf(err) });
    ctx.tally.pagesFailed += 1;
    return;
  }

  ctx.tally.sectionsVisited += 1;

  const liveIds = new Set(live);
  const mirrored = await ctx.deps.store.listPageIdsInSection(section.id);

  for (const pageId of mirrored) {
    if (!liveIds.has(pageId)) await deletePage(ctx, pageId, section, 'sweep');
  }

  // Ids Graph has that the mirror lacks — new, or moved in. Fetched now, budget
  // permitting; anything left is picked up by the next incremental, because this
  // section's watermark is deliberately not advanced here.
  const mirroredIds = new Set(mirrored);
  for (const pageId of live) {
    if (mirroredIds.has(pageId)) continue;
    if (ctx.budget.exhausted) {
      ctx.tally.done = false;
      return;
    }
    await syncPage(ctx, section, {
      id: pageId,
      title: '',
      lastModifiedDateTime: new Date(0).toISOString(),
    });
  }

  await ctx.deps.store.setSectionSweepResult(section.id, {
    pageCount: live.length,
    lastSweptAt: ctx.startedAtIso,
  });
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
  readonly store: Pick<SyncStore, 'getPage' | 'putPage' | 'deletePage'> & {
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
