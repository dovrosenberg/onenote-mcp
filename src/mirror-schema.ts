// The mirror's document shapes, and every decision about them that could be wrong.
//
// Nothing here touches Firestore, Cloud Storage, or the network. That is the point:
// ./mirror-store.ts and ./mirror-blobs.ts cannot be tested on this machine — there is no
// Firestore emulator here and CLAUDE.md rules out an in-memory fake, because what is at
// stake in those files is transaction behaviour and FieldValue.serverTimestamp() and a
// fake would assert the fake. So everything that is a *decision* rather than a call
// lives here, where it runs in a plain unit test: document ids, the field shapes, the
// oversize threshold, the watermark arithmetic, and how a hand-edited document is read.
//
// Two layout choices are worth stating up front, because both look arbitrary and neither
// is:
//
// **`pages` and `pageContent` are two collections keyed identically.** Every query the
// read tools make reads page metadata; none of them reads the HTML, which is fetched by
// document key. Keeping a megabyte of markup out of the queried collection is what makes
// an unscoped title scan cost 2000 small document reads instead of gigabytes of
// transfer, and it is where a vector field would go later — Firestore's FindNearest
// operates on the queried collection.
//
// **A deleted page is hard-deleted, with a tombstone in its own collection.** The
// alternative — a `deleted` boolean on the page — puts an equality filter on the front
// of every read query, turns each into a three-field composite index, and makes an
// omitted filter a silent correctness bug that returns deleted pages as live ones. A
// separate collection makes the wrong query impossible to write.

/** Raised when a value cannot be turned into, or read out of, a mirror document. */
export class MirrorSchemaError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'MirrorSchemaError';
  }
}

// ---------------------------------------------------------------------------
// Collection names
//
// Subcollections of the root document named by MIRROR_ROOT_DOC. They are fixed strings
// rather than derived from the root path because a Firestore composite index is keyed by
// collection id alone, and scripts/gcp-bootstrap.sh creates those indexes by name.
// ---------------------------------------------------------------------------

export const NOTEBOOKS_COLLECTION = 'notebooks';
export const SECTION_GROUPS_COLLECTION = 'sectionGroups';
export const SECTIONS_COLLECTION = 'sections';
export const PAGES_COLLECTION = 'pages';
export const PAGE_CONTENT_COLLECTION = 'pageContent';
export const TOMBSTONES_COLLECTION = 'tombstones';

/** The single document holding all machine-written sync state. */
export const SYNC_STATE_PATH = ['sync', 'state'] as const;

/** Bumped when a stored shape changes in a way a running service would misread. */
export const SCHEMA_VERSION = 1;

// ---------------------------------------------------------------------------
// Document ids
// ---------------------------------------------------------------------------

/** Firestore's own limit on a document id, in UTF-8 bytes. */
const MAX_DOCUMENT_ID_BYTES = 1500;

/**
 * A Graph id as a Firestore document id.
 *
 * Firestore forbids `/` in a document id, and a OneNote id can contain one. It also
 * forbids `.`, `..`, and anything matching `__…__`. `encodeURIComponent` handles the
 * slash and leaves everything else in a real OneNote id untouched; the other three are
 * checked rather than encoded, because they cannot arise from a Graph id and a value
 * that produced one would mean something upstream is badly wrong.
 *
 * The Graph id is also stored verbatim in the document's `id` field, so nothing ever
 * decodes one to read it. There is deliberately no `decodeMirrorId`.
 */
export function encodeMirrorId(graphId: string): string {
  if (graphId === '') {
    throw new MirrorSchemaError('A Graph id cannot be empty.');
  }

  const encoded = encodeURIComponent(graphId);

  if (encoded === '.' || encoded === '..') {
    throw new MirrorSchemaError('A Graph id cannot encode to "." or "..".');
  }
  if (/^__.*__$/.test(encoded)) {
    throw new MirrorSchemaError('A Graph id cannot encode to a reserved "__…__" name.');
  }
  if (Buffer.byteLength(encoded, 'utf8') > MAX_DOCUMENT_ID_BYTES) {
    throw new MirrorSchemaError(
      `A Graph id encodes to ${Buffer.byteLength(encoded, 'utf8')} bytes, over Firestore's ${MAX_DOCUMENT_ID_BYTES}-byte document id limit.`,
    );
  }

  return encoded;
}

// ---------------------------------------------------------------------------
// The hand-edited selection
// ---------------------------------------------------------------------------

/** What a human writes in the root document. The service never writes this. */
export interface NotebookSelection {
  /** Graph notebook ids whose *pages* are mirrored. */
  readonly notebookIds: readonly string[];
  /**
   * Graph notebook ids a sync re-checks after the first backfill.
   *
   * `null` means "every selected notebook is active", which is what a document that has
   * never heard of this field reads as. An empty array is a different thing — a
   * deliberate "freeze everything" — and the two cannot be told apart by a
   * `readonly string[]` that happened to be empty, which is why this is nullable.
   */
  readonly activeNotebookIds: readonly string[] | null;
}

/**
 * Is this notebook one a sync re-checks?
 *
 * True when the operator named no active set at all, or named this notebook in it. It
 * deliberately does not consult `notebookIds`: a notebook whose pages are not mirrored
 * never reaches a code path that asks.
 */
export function isActive(selection: NotebookSelection, notebookId: string): boolean {
  return selection.activeNotebookIds === null || selection.activeNotebookIds.includes(notebookId);
}

/**
 * The two selection lists a previous run recorded, as `MirrorSyncState` holds them.
 *
 * `mirrored` is null only when nothing was ever recorded; `active` is null both for that
 * and for "every mirrored notebook is active". `notebooksNeedingWideScan` is given the
 * pair so the caller does not have to decide which null it is looking at.
 */
export interface SeenSelection {
  readonly mirrored: readonly string[] | null;
  readonly active: readonly string[] | null;
}

/**
 * Which notebooks a selection change has to widen the next scan for.
 *
 * A notebook that has just been added to `notebookIds`, or has just become active, holds
 * sections whose `graphLastModifiedDateTime` may be months old. Tier 1 of
 * `pickCandidates` in ./mirror-sync.ts skips a section older than the cutoff derived from
 * `sectionsScannedThrough`, and that cutoff advances on every completed run, so nothing
 * else would ever make those sections candidates again.
 *
 * Removal widens nothing. A notebook dropped from either list is one this run does less
 * work for, and there is no backlog to catch up on.
 *
 * Only *candidacy* is widened. Each named notebook's sections still list against their own
 * `pagesSyncedThrough`, which nothing here touches, so a widened section costs one
 * `listPagesChangedSince` and re-fetches only the pages that actually changed.
 *
 * `previous.mirrored === null` means no run has ever recorded these lists — a state
 * document written before the fields existed. Nothing was skipped for inactivity under it
 * that a recorded list would have caught, so the answer is empty and the caller's job is
 * to record.
 *
 * The active set of a `null` list is every mirrored notebook, so a list becoming null
 * names everything that was not already in it. The result is deduped and sorted, because
 * it is stored and compared and the order a diff produced it in is not information.
 */
export function notebooksNeedingWideScan(
  previous: SeenSelection,
  current: NotebookSelection,
  mirroredNotebookIds: readonly string[],
): string[] {
  if (previous.mirrored === null) return [];

  const wasMirrored = new Set(previous.mirrored);
  const newlyMirrored = current.notebookIds.filter((id) => !wasMirrored.has(id));

  const wasActive = new Set(previous.active ?? previous.mirrored);
  const nowActive = current.activeNotebookIds ?? mirroredNotebookIds;
  const newlyActive = nowActive.filter((id) => !wasActive.has(id));

  return [...new Set([...newlyMirrored, ...newlyActive])].sort();
}

/**
 * Do the recorded lists already say what the selection document says?
 *
 * `notebooksNeedingWideScan` answering `[]` is not the same question, and using it as one
 * loses a re-activation. Deactivating a notebook widens nothing, so a run that only wrote
 * when it had something to widen would leave `activeNotebookIdsSeen` naming the notebook
 * that was just frozen — and re-activating it later would then diff against a list it is
 * already in and widen nothing, which is the case the widening exists for.
 *
 * Order is not a change: `readSelection` preserves the order an operator typed, and
 * retyping the same ids in another order means nothing here. A `null` active list is
 * compared as a value rather than as a set, because `null` and a list naming every
 * mirrored notebook are the same active set today and different ones the moment a
 * notebook is added.
 */
export function selectionMatchesSeen(
  previous: SeenSelection,
  current: NotebookSelection,
): boolean {
  if (previous.mirrored === null) return false;
  if (!sameIds(previous.mirrored, current.notebookIds)) return false;
  if (previous.active === null || current.activeNotebookIds === null) {
    return previous.active === null && current.activeNotebookIds === null;
  }
  return sameIds(previous.active, current.activeNotebookIds);
}

/**
 * Do two id lists hold the same ids, order aside?
 *
 * Precondition: neither side holds duplicates. Length plus one-way containment is
 * set-equality only under that — `sameIds(['a', 'y'], ['a', 'a'])` answers true. Both
 * callers' inputs come through `readIdList`, which collapses duplicates, so the
 * precondition holds today; a third caller reading a list from somewhere else has to
 * check it.
 */
function sameIds(a: readonly string[], b: readonly string[]): boolean {
  if (a.length !== b.length) return false;
  const set = new Set(a);
  return b.every((id) => set.has(id));
}

/**
 * Read the selection document, tolerating anything a person might leave in it.
 *
 * A missing document, an absent `notebookIds`, a non-array, or an array holding
 * non-strings all resolve to "mirror nothing" rather than throwing. A human edits this
 * by hand in the Firestore console, and a half-finished edit must not take the sync down
 * — it should mean no pages are mirrored this run, which is visible and recoverable.
 * Non-string and empty entries are dropped individually, so one bad row does not discard
 * the good ones beside it.
 *
 * Duplicates are collapsed. Order is preserved, because the sync reports unmatched ids
 * back and matching that against what was typed is easier in the original order.
 *
 * `activeNotebookIds` is read by the same rules with one difference: a field that is
 * absent or not an array becomes `null` rather than `[]`. See `NotebookSelection`.
 */
export function readSelection(data: Record<string, unknown> | undefined): NotebookSelection {
  const active = readIdList(data?.['activeNotebookIds']);
  const raw = data?.['notebookIds'];
  if (!Array.isArray(raw)) return { notebookIds: [], activeNotebookIds: active };

  return { notebookIds: readIdList(raw) ?? [], activeNotebookIds: active };
}

/**
 * A hand-edited array of ids, or null when the field is not an array at all.
 *
 * The null is what `activeNotebookIds` needs and `notebookIds` discards. A malformed
 * active list must read as "every selected notebook is active" rather than as "none":
 * failing open costs Graph requests, and failing closed would freeze the mirror with
 * nothing saying so.
 */
function readIdList(raw: unknown): string[] | null {
  if (!Array.isArray(raw)) return null;

  const seen = new Set<string>();
  const ids: string[] = [];
  for (const entry of raw) {
    if (typeof entry !== 'string') continue;
    const trimmed = entry.trim();
    if (trimmed === '' || seen.has(trimmed)) continue;
    seen.add(trimmed);
    ids.push(trimmed);
  }

  return ids;
}

// ---------------------------------------------------------------------------
// Watermarks
// ---------------------------------------------------------------------------

/**
 * How far back a *section's page listing* reaches beyond that section's watermark.
 *
 * Graph's clock and this service's clock are not the same clock, and the watermark is
 * the time a pass *started* rather than the newest page it saw. A page this window pulls
 * into a listing costs nothing on its own: the listing already carries the page's stamp
 * and title, and `storedPageIsCurrent` skips a stored copy that matches without a Graph
 * request. So the hour buys margin and spends bytes.
 *
 * Not the window that decides which *sections* are listed at all. That one costs one
 * request per section per run and is `SECTION_SCAN_OVERLAP_MS`, below.
 */
export const WATERMARK_OVERLAP_MS = 3_600_000;

/**
 * How far back the *section scan* reaches beyond `sectionsScannedThrough`.
 *
 * Separate from `WATERMARK_OVERLAP_MS` because the two windows cost different things.
 * That one pulls unchanged pages into a listing that is already being made. This one
 * keeps an edited section a candidate, and every candidate costs one
 * `listPagesChangedSince` on every run for as long as the window lasts — so an hour here
 * re-listed a section edited once on every run for the next hour, against a budget of 400
 * requests an hour shared with the interactive tools.
 *
 * Fifteen minutes. Both windows cover the same two things, and both are now measured
 * rather than guessed at:
 *
 * - **Clock skew.** `api-overview.md` bounds it at seconds rather than minutes — a page
 *   created inside a 15-second local bracket was stamped by Graph inside that bracket.
 *   `recordClockSkew` in ./graph-throttle.ts reports it continuously from the running
 *   service, and `graph-clock-skew` in the logs is the number to read.
 * - **Propagation.** A section's `lastModifiedDateTime` rolls up from page creates, edits
 *   and deletes, measured 2026-08-19. The largest movement ever seen in a stamp that had
 *   nothing done to it is one second, and that shift was not path-dependent.
 *
 * What the window has to cover is the sum of those two, and nothing else. The cutoff is
 * derived from the start of the last *completed* run, so an edit made between two runs
 * carries a stamp after that instant and is a candidate under a window of any width. The
 * case the window exists for is an edit made *before* a run started that the run's tree
 * read did not yet reflect: the next cutoff has to still be behind it, which it is
 * whenever the lag is under the window. Fifteen minutes is about sixty times the largest
 * effect either measurement above shows. The lag itself has never been measured directly,
 * which is why the log line below is the thing to watch rather than an argument.
 *
 * **What would show fifteen minutes is too narrow:** an `ageMs` above 900000 on a
 * `sync-overlap-save` line. That event fires for a page whose stamp predates its
 * section's watermark *and* that turned out to have really changed — a page the window is
 * the only reason the sync saw. The largest `ageMs` over a long run is the narrowest
 * window that would still have caught everything, and a value above this constant is a
 * page a 15-minute scan window would have reached late. `README.md` carries the query
 * under **Proving it works**.
 *
 * Being wrong here costs a night rather than an edit: the nightly `/sync/sweep/full`
 * compares stamps and titles against every stored page, so a section this window declined
 * is reconciled then. A notebook just mirrored or just activated does not depend on this
 * window at all — `notebooksNeedingWideScan` covers that case.
 */
export const SECTION_SCAN_OVERLAP_MS = 900_000;

/**
 * The instant a re-query should start from, given a watermark.
 *
 * A null watermark means the section has never been synced, which asks for everything:
 * the epoch. Overlap is subtracted rather than added, and the result is clamped to never
 * exceed the watermark itself, so a negative or absurd overlap cannot skip a window.
 *
 * The default is the page-listing window, because that is the call with a watermark of
 * its own to reach back from. The section scan passes `SECTION_SCAN_OVERLAP_MS`
 * explicitly.
 */
export function overlapFrom(
  watermarkIso: string | null,
  overlapMs: number = WATERMARK_OVERLAP_MS,
): string {
  if (watermarkIso === null) return new Date(0).toISOString();

  const watermark = Date.parse(watermarkIso);
  if (Number.isNaN(watermark)) {
    // An unparseable watermark is treated as never-synced. Refusing to sync would be the
    // worse failure: the value is written by this service, so a bad one means something
    // already went wrong and the recovery is a full re-read, not a halt.
    return new Date(0).toISOString();
  }

  return new Date(Math.min(watermark, watermark - overlapMs)).toISOString();
}

/**
 * How far a page's stamp predates the watermark, when the overlap is why it was listed.
 *
 * Both overlap windows are margins against Graph's propagation lag and the gap between
 * its clock and this process's. A page whose stamp is older than the watermark is one
 * that only `overlapFrom`'s reach put in front of the sync; the largest age ever seen for
 * a page that turned out to have really changed is the narrowest window that would still
 * have caught everything, and it is the evidence `SECTION_SCAN_OVERLAP_MS` names for
 * whether fifteen minutes is too narrow.
 *
 * Returns null when the page is not one of those: a stamp at or after the watermark would
 * have been listed by a window of any width, a null watermark means the section asked for
 * everything from the epoch, and an unparseable value on either side is a measurement
 * that cannot be taken rather than one to guess at.
 *
 * The answer is a number of milliseconds. It names no page and no section, which is what
 * makes it safe in a log line.
 */
export function overlapSaveAgeMs(watermarkIso: string | null, modifiedIso: string): number | null {
  if (watermarkIso === null) return null;

  const watermark = Date.parse(watermarkIso);
  const modified = Date.parse(modifiedIso);
  if (Number.isNaN(watermark) || Number.isNaN(modified)) return null;
  if (modified >= watermark) return null;

  return watermark - modified;
}

// ---------------------------------------------------------------------------
// Page HTML placement
// ---------------------------------------------------------------------------

/**
 * The largest raw page HTML that goes inside a Firestore document, in UTF-8 bytes.
 *
 * Firestore's document limit is 1 MiB and index entries count against it. 700 KB rather
 * than something nearer the cap because the byte length of the HTML is not the size of
 * the document — there are other fields, and UTF-8 bytes are not JS string length. The
 * margin is cheap; being wrong means a failed write on the largest and most valuable
 * pages.
 */
export const HTML_INLINE_LIMIT_BYTES = 700_000;

export type HtmlLocation = 'firestore' | 'gcs';

/**
 * Where this page's HTML goes.
 *
 * Over the limit it spills to Cloud Storage. It is never truncated and never skipped.
 * Truncation would violate the rule ./page-html.ts is built on — no character of the
 * user's writing is lost — and would hand a model half a page with no way to know.
 * Skipping would make exactly the pages that are most expensive to fetch from Graph the
 * ones that always miss, which inverts the point of the mirror.
 */
export function htmlPlacement(bytes: number): HtmlLocation {
  return bytes > HTML_INLINE_LIMIT_BYTES ? 'gcs' : 'firestore';
}

/** UTF-8 byte length, which is what every Firestore and GCS limit is measured in. */
export function utf8Bytes(text: string): number {
  return Buffer.byteLength(text, 'utf8');
}

// ---------------------------------------------------------------------------
// Object names in the mirror's bucket
//
// Prefixed by kind so a human listing the bucket can tell the three apart, and keyed by
// the encoded page id so `deleteForPage` can build all three without a lookup.
// ---------------------------------------------------------------------------

export function inkObjectName(pageId: string): string {
  return `ink/${encodeMirrorId(pageId)}.png`;
}

export function inkmlObjectName(pageId: string): string {
  return `inkml/${encodeMirrorId(pageId)}.xml`;
}

export function htmlObjectName(pageId: string): string {
  return `html/${encodeMirrorId(pageId)}.html`;
}

// ---------------------------------------------------------------------------
// Stored shapes
//
// Every one carries the Graph id verbatim in `id`, so nothing decodes a document id.
// Timestamps written by Graph stay ISO strings exactly as Graph spelled them, because
// every tool result already prints that string; the parallel `lastModified` field exists
// only so Firestore can order and range-filter on it.
// ---------------------------------------------------------------------------

export interface MirrorNotebook {
  readonly id: string;
  readonly displayName: string;
  /** Is this notebook in the hand-edited selection — that is, are its pages mirrored. */
  readonly mirrored: boolean;
  readonly sectionCount: number;
  readonly sectionGroupCount: number;
  readonly graphLastModifiedDateTime: string | null;
}

export type ContainerParentKind = 'notebook' | 'sectionGroup';

export interface MirrorSectionGroup {
  readonly id: string;
  readonly displayName: string;
  readonly notebookId: string;
  readonly parentId: string;
  readonly parentKind: ContainerParentKind;
  readonly mirrored: boolean;
  /** `2026 / 062 - February` — the containers above it, then itself. */
  readonly path: string;
  /**
   * Have this group's own nested section groups been enumerated?
   *
   * Load-bearing. `$expand` reaches one level of section group, so a first-level group's
   * nested groups are *absent from the response* rather than known to be empty. A
   * `list_sections` answered from a document where this is false would omit them and
   * look complete, so the read path treats false as a mirror miss and goes to Graph. The
   * sweep sets it true for mirrored notebooks.
   */
  readonly childGroupsKnown: boolean;
}

export interface MirrorSection {
  readonly id: string;
  readonly displayName: string;
  readonly notebookId: string;
  readonly parentId: string;
  readonly parentKind: ContainerParentKind;
  /** `2026 / 062 - February / Daily log` — what `SectionRef.path` already means. */
  readonly path: string;
  readonly mirrored: boolean;
  readonly graphLastModifiedDateTime: string | null;
  /**
   * The per-section watermark: the time the last completed pass over this section
   * *started*, or null when it has never been synced.
   *
   * Per section rather than global because a budget-bounded run does not finish, and a
   * global watermark advanced past sections it never visited loses every edit in them
   * permanently. Per-section makes a partial run trivially correct and needs no cursor.
   */
  readonly pagesSyncedThrough: string | null;
  readonly pageCount: number;
  /**
   * How many writes are in flight against this section's *page listing*.
   *
   * A page listing served from the mirror reports which pages a section holds and what
   * they are called. Two writes change that and cannot be repaired by marking a page
   * stale: `create_page` adds a page the mirror has no document for, and
   * `update_page_title` changes the title every listing and by-name lookup matches on.
   * Both raise this before they call Graph and lower it after the resync, so the whole
   * window answers `list_pages`, `list_pages_by_name`, `find_page_by_name` and
   * `search_pages` from Graph rather than from a copy that is provably behind.
   *
   * A count rather than a flag because two writes against one section can overlap: the
   * first to finish must not clear the second's hold.
   *
   * An append does not raise it. It changes the page's content, which
   * `contentState: 'stale'` already covers, and its `lastModifiedDateTime`, which only
   * moves the page within an ordering — no listing becomes wrong about what exists or
   * what it is called.
   *
   * A structure write does not clear it. `putStructure` sends the tree-owned fields and
   * merges, and this is not one of them — see `SectionTreeFields`. It used to, which was
   * a rare race with an in-flight write; the expiry below is what still covers a hold
   * nothing lowers.
   */
  readonly pendingWrites?: number;
  /** When the most recent hold was taken, so a hold a dead process left expires. */
  readonly pendingWritesSince?: string;
}

// ---------------------------------------------------------------------------
// Tree-owned fields, and the identity that says whether one changed
//
// A structure document has two owners. `buildStructure` in ./mirror-sync.ts knows what
// the expanded tree said — a name, a parent, a path, whether the notebook is in the
// selection. The sync and the write tools own everything else — a section's watermark and
// page count, a group's `childGroupsKnown`, a section's listing hold. Those are learned
// over hours of Graph requests and the tree read knows nothing about them.
//
// Before this split a structure write sent whole documents built from the tree alone, so
// every one of the sync-owned fields was reset to its creation default whenever the tree
// hash moved — for a section renamed in the OneNote client, or a notebook created in a
// part of the account that is not even mirrored. A reset watermark is a full re-backfill
// of the selection; a reset `childGroupsKnown` sends every `list_sections` on a section
// group to Graph until the next sweep.
// ---------------------------------------------------------------------------

export type NotebookTreeFields = Pick<
  MirrorNotebook,
  | 'id'
  | 'displayName'
  | 'mirrored'
  | 'sectionCount'
  | 'sectionGroupCount'
  | 'graphLastModifiedDateTime'
>;

export type SectionGroupTreeFields = Pick<
  MirrorSectionGroup,
  'id' | 'displayName' | 'notebookId' | 'parentId' | 'parentKind' | 'path' | 'mirrored'
>;

export type SectionTreeFields = Pick<
  MirrorSection,
  | 'id'
  | 'displayName'
  | 'notebookId'
  | 'parentId'
  | 'parentKind'
  | 'path'
  | 'mirrored'
  | 'graphLastModifiedDateTime'
>;

/** The `identity` field every structure document carries, so the store can skip a write. */
export interface StructureIdentity {
  readonly identity: string;
}

export type NotebookStructureWrite = NotebookTreeFields & StructureIdentity;
export type SectionGroupStructureWrite = SectionGroupTreeFields & StructureIdentity;
export type SectionStructureWrite = SectionTreeFields & StructureIdentity;

/**
 * What a structure document is created with the first time it appears, and never again.
 *
 * `pagesSyncedThrough: null` is what makes a brand new section a backfill candidate —
 * `pickCandidates` in ./mirror-sync.ts treats it as "never synced" and reads it whatever
 * its timestamp says. Re-sending it to a section that already exists is the reset this
 * split removes, so the store applies these only on a document it just created.
 *
 * The field is not optional, and the reason is a Firestore rule rather than a convention.
 * `listSectionsToSync` orders by `pagesSyncedThrough`, and an `orderBy` excludes every
 * document that does not carry the field — so a section created without it is absent from
 * the sync's own listing for ever, and nothing reports a section it never enumerated.
 * `buildStructure` used to guarantee the field by sending it on every write; now the only
 * thing that guarantees it is this default being applied on a create.
 */
export const NEW_SECTION_DEFAULTS = { pagesSyncedThrough: null, pageCount: 0 } as const satisfies
  Pick<MirrorSection, 'pagesSyncedThrough' | 'pageCount'>;

/**
 * `childGroupsKnown: false` is what a group is created with, for the reason the field's
 * own docstring gives: `$expand` reaches one level of section group, so a new group's
 * nested groups are absent from the response rather than known to be empty.
 */
export const NEW_GROUP_DEFAULTS = { childGroupsKnown: false } as const satisfies
  Pick<MirrorSectionGroup, 'childGroupsKnown'>;

/**
 * The separator between fields in an identity string.
 *
 * A NUL, because a OneNote display name can hold every other plausible separator. With a
 * space, `{displayName: 'a b', notebookId: 'c'}` and `{displayName: 'a', notebookId: 'b c'}`
 * join to the same string, so a rename that moved a space across a field boundary would
 * produce an unchanged identity and the document would never be rewritten.
 */
const IDENTITY_SEPARATOR = '\0';

/**
 * What the tree said about this notebook, as one string.
 *
 * Two documents with the same identity are the same as far as a structure write is
 * concerned, so the store skips the second one. `structureHashOf` in ./mirror-sync.ts
 * hashes these same strings, which is what stops "the hash moved" and "this document
 * changed" from ever disagreeing.
 *
 * `graphLastModifiedDateTime` is deliberately absent from all three identities. It moves
 * whenever anyone edits a page, so including it would rewrite every document in the
 * account on nearly every run — and nothing reads the stored copy for freshness anyway:
 * `withLiveMtimes` overlays what this run's tree read saw before `pickCandidates` filters
 * on it.
 */
export function notebookIdentity(fields: NotebookTreeFields): string {
  const {
    id,
    displayName,
    mirrored,
    sectionCount,
    sectionGroupCount,
    graphLastModifiedDateTime,
    ...rest
  } = fields;
  everyFieldConsidered(rest, graphLastModifiedDateTime);

  return [
    'n',
    id,
    displayName,
    String(mirrored),
    String(sectionCount),
    String(sectionGroupCount),
  ].join(IDENTITY_SEPARATOR);
}

/** See `notebookIdentity`. `childGroupsKnown` is sync-owned and is not a tree field. */
export function groupIdentity(fields: SectionGroupTreeFields): string {
  const { id, displayName, notebookId, parentId, parentKind, path, mirrored, ...rest } = fields;
  everyFieldConsidered(rest);

  return ['g', id, displayName, notebookId, parentId, parentKind, path, String(mirrored)].join(
    IDENTITY_SEPARATOR,
  );
}

/** See `notebookIdentity`. The watermark and the page count are sync-owned. */
export function sectionIdentity(fields: SectionTreeFields): string {
  const {
    id,
    displayName,
    notebookId,
    parentId,
    parentKind,
    path,
    mirrored,
    graphLastModifiedDateTime,
    ...rest
  } = fields;
  everyFieldConsidered(rest, graphLastModifiedDateTime);

  return ['s', id, displayName, notebookId, parentId, parentKind, path, String(mirrored)].join(
    IDENTITY_SEPARATOR,
  );
}

/**
 * A compile-time check that an identity function has looked at every field of its input.
 *
 * Each of the three above destructures its parameter completely — the fields that go into
 * the string, then the ones deliberately left out, then a rest binding — and passes the
 * rest here. A field added to one of the tree-field types and not named in the
 * destructuring lands in `rest`, and `Record<string, never>` refuses it.
 *
 * Without this, adding a field to `SectionTreeFields` is a type error in `buildStructure`,
 * where the object literal is checked, and *not* in `sectionIdentity`, which would keep
 * compiling while ignoring it. A section whose only change was that field would then have
 * an unchanged identity and would never be written — the silent-skip failure this whole
 * design is built to avoid, arriving through the one door the types left open.
 *
 * `omitted` takes the deliberate exclusions so a destructured-but-unused binding is a use.
 * Nothing here runs for any other reason; the parameter types are the whole point.
 */
function everyFieldConsidered(_rest: Record<string, never>, ..._omitted: unknown[]): void {}

// ---------------------------------------------------------------------------
// What a structure write does to one collection
// ---------------------------------------------------------------------------

/** One document to write, already keyed by its Firestore document id. */
export interface PlannedStructureWrite {
  readonly documentId: string;
  /** Ready to `set(…, { merge: true })`. Creation defaults are already folded in. */
  readonly fields: Record<string, unknown>;
  /**
   * True when no document with this id was there. Only a create carries the defaults.
   *
   * Nothing in `src/` reads it — `#reconcileCollection` sends every write the same way. It
   * is here for the tests, which cannot otherwise tell a create from a merge on a notebook:
   * `createDefaults` is `{}` there, so the written fields are identical either way. Do not
   * delete it as dead, and do not log from it expecting the store to have branched on it.
   */
  readonly created: boolean;
}

export interface StructureWritePlan {
  readonly writes: readonly PlannedStructureWrite[];
  /** Document ids the tree no longer holds. */
  readonly deletes: readonly string[];
}

/**
 * Decide what one structure collection needs, given what it holds and what the tree says.
 *
 * This is the whole of `#replaceCollection`'s judgement, and it lives here because
 * ./mirror-store.ts cannot be tested on this machine — the rule stated at the top of this
 * file. That module is left holding the Firestore query, the batching and the calls; every
 * branch below is reachable from a plain unit test.
 *
 * Four outcomes, and each one exists because its absence is a specific failure:
 *
 * - **Skip** a document whose stored `identity` equals the incoming one. Without it,
 *   renaming a single section rewrites all 568 section documents in the account.
 * - **Merge** a document whose identity differs, by sending the tree fields and nothing
 *   else. The fields the tree does not know — `pagesSyncedThrough`, `pageCount`,
 *   `lastSweptAt`, `childGroupsKnown`, `pendingWrites`, `pendingWritesSince` — are absent
 *   from `fields`, so a merging write leaves them where they are. A cleared watermark is a
 *   full re-backfill of that section.
 * - **Create** with `createDefaults` a document that was not there. See
 *   `NEW_SECTION_DEFAULTS` for what a missing `pagesSyncedThrough` costs.
 * - **Delete** a document the incoming set does not name. By absence rather than by a
 *   `lastSeenAt` sweep, because the caller has just read the complete tree in one request
 *   — so absence is a fact rather than an inference, and it must not be reached by the
 *   skip above.
 *
 * One consequence of merging that the old wholesale write did not have: a field dropped
 * from `buildStructure` is no longer removed from the stored documents. `set(…, { merge:
 * true })` leaves a field it does not mention exactly where it is, so a retired field
 * lingers until a `FieldValue.delete()` or a one-off pass clears it. Retiring a field is
 * therefore a migration now, not an edit.
 *
 * `storedIdentities` maps document id to whatever the stored `identity` field held. A
 * present key with an `undefined` value is a document written before that field existed:
 * it exists, so it is merged rather than created, and it does not match, so it is written
 * once. That is the one-off pass every deployment of this change pays.
 */
export function planStructureWrite(
  storedIdentities: ReadonlyMap<string, unknown>,
  documents: readonly (StructureIdentity & { id: string })[],
  createDefaults: object,
): StructureWritePlan {
  const unseen = new Set(storedIdentities.keys());
  const writes: PlannedStructureWrite[] = [];

  for (const document of documents) {
    const documentId = encodeMirrorId(document.id);
    const existed = storedIdentities.has(documentId);
    unseen.delete(documentId);

    if (existed && storedIdentities.get(documentId) === document.identity) continue;

    writes.push({
      documentId,
      fields: { ...(existed ? {} : createDefaults), ...document },
      created: !existed,
    });
  }

  return { writes, deletes: [...unseen] };
}

export type ContentState = 'present' | 'stale' | 'missing';

export interface MirrorPageInk {
  readonly objectName: string;
  readonly inkmlObjectName: string;
  readonly width: number;
  readonly height: number;
  readonly strokeCount: number;
  readonly bytes: number;
}

export interface MirrorPage {
  readonly id: string;
  readonly title: string;
  /** Lowercased, for a future prefix query. Nothing reads it in this pass. */
  readonly titleLower: string;
  readonly sectionId: string;
  readonly notebookId: string;
  readonly sectionPath: string;
  /** Exactly as Graph spelled it; every tool result already prints this string. */
  readonly lastModifiedDateTime: string;
  readonly contentState: ContentState;
  /** sha256 of the raw HTML. An unchanged hash skips the write and the ink render. */
  readonly contentHash: string | null;
  readonly htmlLocation: HtmlLocation;
  readonly htmlObject: string | null;
  readonly htmlBytes: number;
  readonly ink: MirrorPageInk | null;
  /**
   * When the sync last wrote this page's content.
   *
   * Typed `unknown` because it is written with `FieldValue.serverTimestamp()` and comes
   * back as a Firestore `Timestamp`, not a string. `timestampToIso` is what reads it;
   * nothing else should touch it.
   */
  readonly contentSyncedAt?: unknown;
}

/**
 * A Firestore timestamp as an ISO string, or null for anything unrecognised.
 *
 * Three shapes reach this: a `Timestamp` with `toDate()`, the plain
 * `{_seconds, _nanoseconds}` a serialised one decodes to, and an absent field on a
 * document written before the field existed. None of them is worth throwing over — the
 * value's only job is to tell a model how stale an answer is, and "unknown" is a
 * survivable answer where a failed tool call is not.
 */
export function timestampToIso(value: unknown): string | null {
  if (value === null || value === undefined) return null;

  if (typeof value === 'object' && 'toDate' in value) {
    const date = (value as { toDate: () => Date }).toDate();
    return Number.isNaN(date.getTime()) ? null : date.toISOString();
  }

  if (typeof value === 'object' && '_seconds' in value) {
    const seconds = (value as { _seconds: unknown })._seconds;
    if (typeof seconds !== 'number' || !Number.isFinite(seconds)) return null;
    return new Date(seconds * 1000).toISOString();
  }

  if (typeof value === 'string') {
    const parsed = Date.parse(value);
    return Number.isNaN(parsed) ? null : new Date(parsed).toISOString();
  }

  return null;
}

/**
 * The three fields a page listing carries. `PageSummary` from ./graph-structure.ts
 * satisfies it.
 *
 * Declared structurally rather than imported, so this module keeps its property of having
 * no imports at all.
 */
export interface PageStamp {
  readonly id: string;
  readonly title: string;
  readonly lastModifiedDateTime: string;
}

/**
 * Do the mirror's stamp and Graph's stamp for one page disagree?
 *
 * A hint, not a verdict. It says the two strings are not the same string; it does not
 * claim the page changed, and the name says so — the sweep's answer to a disagreement is
 * to re-fetch the page, and `writePageFromRaw`'s content-hash comparison is what decides
 * whether anything is written. That is why there is no margin, no direction test and no
 * `Date.parse` here:
 *
 * - **No margin.** A tolerance buys jitter-safety by discarding every real edit inside the
 *   window, and an edit the sweep never notices is served as current for ever. A false
 *   positive costs one Graph request.
 * - **No direction.** Stored-ahead-of-live is the signature of a page written through this
 *   server: `resyncPage` stamps `new Date().toISOString()` from this process's clock,
 *   because Graph's page *metadata* read is measured-unreliable (`api-overview.md`:
 *   `GET /pages/{id}?$select=title` answers `""` for a page created seconds earlier). Under
 *   the old response — `markPageStale`, which deletes the page-content document — firing on
 *   that direction destroyed the mirrored copy of exactly the most-used pages in the
 *   account, and nothing put it back, because no read path ever writes to the mirror. A
 *   re-fetch does the opposite: it replaces the local stamp with Graph's own. **Do not put
 *   a mark back in that branch.**
 * - **No parsing.** Two spellings of one instant (`…:00Z` and `…:00.000Z`) disagree, and
 *   that is wanted: the second is what `resyncPage` writes, `lastModifiedDateTime` is
 *   printed in every tool result, and a re-fetch is what replaces it with Graph's spelling.
 *   A parsed compare would leave the local spelling stored for ever.
 *
 * A page the pre-2026-08-21 sweep discovered carries `title: ''` and `new Date(0)`, and
 * needs no predicate of its own: the epoch disagrees with anything Graph sends, so the same
 * branch re-fetches it and the repair writes both fields from Graph's listing.
 *
 * An empty live stamp is `toPageSummary`'s fallback for a field Graph did not send. It
 * disagrees like any other value, so such a page is re-fetched once per sweep and the
 * short-circuit declines to store the empty string over a good one. That is the price of
 * never hiding an edit, and no page Graph stamps ever pays it.
 *
 * `storedPageIsCurrent` delegates its stamp clause here rather than restating the
 * comparison. Both callers answer a disagreement by fetching the page, so one rule serves
 * both, and a change made to this one — a margin, a parse, a direction test — cannot reach
 * only one of them.
 */
export function pageStampDiffers(stored: PageStamp, live: PageStamp): boolean {
  return stored.lastModifiedDateTime !== live.lastModifiedDateTime;
}

/**
 * How long a stored copy must have been in hand before its stamp is trusted.
 *
 * Graph stamps `lastModifiedDateTime` to the whole second — measured 2026-08-21 and
 * recorded in api-overview.md under "Page timestamps have whole-second resolution". So a
 * content fetch landing in the same second as a concurrent edit stores the older content
 * under a stamp that never moves, and `storedPageIsCurrent` would then skip that page on
 * every later run. Requiring the stored copy to have been fetched well after the stamp
 * closes that window.
 *
 * Unlike the stamp equality, this comparison **does** cross clocks: Firestore's
 * `serverTimestamp()` on one side, Graph's stamp on the other. That is why it is 30
 * seconds rather than 2. The size is bounded by measurement rather than chosen blind —
 * api-overview.md records the Graph-to-service skew as seconds rather than minutes, and
 * `recordClockSkew` in ./graph-throttle.ts measures it continuously from the deployed
 * service, so a figure that turns out to be wrong shows up in the logs rather than as a
 * silent miss.
 *
 * Failing the guard only ever forces an extra fetch. It never defers work into a later
 * window, so it cannot interact with the watermark overlap.
 */
export const TIMESTAMP_SETTLE_MS = 30_000;

/**
 * Does the mirror already hold this page exactly as Graph describes it?
 *
 * True means the content fetch can be skipped, which is the point: the changed-page
 * listing already carried the stamp and the title, so re-reading the content would spend
 * one of the hourly 400 Graph requests to learn what the listing already said.
 *
 * Every clause is a way that answering true carelessly would serve superseded content as
 * current:
 *
 * - `contentState` — a write already marked this stale, or its content was never stored.
 * - the stamp — exact string equality, delegated to `pageStampDiffers` so the sweep and
 *   the skip cannot drift into two different rules for one comparison. An empty or
 *   unparseable live stamp needs no clause of its own: `Date.parse` answers NaN for it and
 *   every comparison against NaN is false, so the settle guard below refuses the page.
 * - the title — a rename moves no timestamp, so a stamp-only check goes permanently blind
 *   to one. Measured 2026-08-21 and recorded in api-overview.md: a `PATCH
 *   /pages/{id}/content` replacing the title left `lastModifiedDateTime` at its old value,
 *   still unchanged two minutes later. The title is the field `find_page_by_name` and
 *   `search_pages` match on, so a mirror that missed a rename answers by the wrong name
 *   indefinitely. Do not drop this comparison as redundant with the stamp.
 * - the section — page ids survive a rename of their placement, so only the placement
 *   changed and the stored `sectionId` is what every listing query filters on.
 * - the settle guard — see `TIMESTAMP_SETTLE_MS`.
 *
 * **There is no tolerance on the stamp, and a future reader must not add one.** The
 * tempting justification is api-overview.md's record of Graph reporting four unchanged
 * pages exactly one second later across two reads — a window of a second or two would
 * absorb that. It would also discard every genuine edit made inside the window, and the
 * incremental sync advances the section watermark past an edit it did not act on, so such
 * a page is never listed again. The failure directions are not symmetric: a stamp that
 * disagrees for no reason costs one Graph request and then agrees, because the re-fetch
 * writes Graph's own string back.
 */
export function storedPageIsCurrent(
  stored: MirrorPage,
  live: PageStamp,
  sectionId: string,
  settleMs: number = TIMESTAMP_SETTLE_MS,
): boolean {
  if (stored.contentState !== 'present') return false;
  if (pageStampDiffers(stored, live)) return false;
  if (stored.title !== live.title) return false;
  if (stored.sectionId !== sectionId) return false;

  // Null for an absent field, which is what a document written before `contentSyncedAt`
  // existed carries. It cannot prove anything about when the copy was taken.
  const syncedAt = timestampToIso(stored.contentSyncedAt);
  if (syncedAt === null) return false;

  // Not `Math.abs`. A copy fetched *before* Graph stamped the edit is not settled by any
  // margin — it predates the edit outright — and an absolute value would call it current.
  //
  // A stamp neither side can parse — `''`, which is `toPageSummary`'s fallback for a field
  // Graph did not send — makes this NaN, and NaN fails every comparison, so the page is
  // fetched. That is the wanted answer and it is why there is no separate empty-stamp
  // clause above; a guard that could never change the result would be untestable code.
  return Date.parse(syncedAt) - Date.parse(live.lastModifiedDateTime) > settleMs;
}

export interface MirrorPageContent {
  readonly pageId: string;
  /** Raw and untrimmed. `trimPageHtml` runs at read time so the trimmer can change. */
  readonly html: string;
  readonly bytes: number;
  readonly contentHash: string;
}

export type TombstoneReason = 'sweep' | 'not-found';

export interface MirrorTombstone {
  readonly id: string;
  readonly sectionId: string;
  readonly notebookId: string;
  readonly reason: TombstoneReason;
}

export type SyncMode = 'incremental' | 'sweep' | 'sweep-full' | 'sweep-all';
export type SyncOutcome = 'complete' | 'budget-exhausted' | 'failed';

export interface MirrorSyncState {
  readonly schemaVersion: number;
  /** sha256 of the normalised tree; an unchanged hash skips every structure write. */
  readonly structureHash: string | null;
  /** Start of the last run that compared the whole tree. */
  readonly sectionsScannedThrough: string | null;
  /**
   * Does a page change move its section's `lastModifiedDateTime`?
   *
   * Measured true on 2026-08-19 (api-overview.md) — a create, an edit and a delete each
   * move it, and three reads 20 seconds apart with no write did not. It is stored rather
   * than hardcoded so an operator can turn it off without a deploy if the service changes
   * its behaviour; with it off, every mirrored section is visited every run.
   */
  readonly sectionRollUpTrusted: boolean;
  /** Set false by the automatic fallback when the datetime `$filter` answers 400. */
  readonly datetimeFilterSupported: boolean;
  readonly backfillComplete: boolean;
  readonly lastRunOutcome: SyncOutcome | null;
  readonly lastRunGraphRequests: number;
  readonly lastRunPagesUpdated: number;
  readonly lastRunPagesDeleted: number;
  /** Where a budget-exhausted sweep resumes. */
  readonly sweepCursorSectionId: string | null;
  /** When the expanded-tree read last 500'd, so a run on stale structure can say so. */
  readonly lastTreeFailureAt: string | null;
  /** The lease. A second run while one is held is a 409, not a double spend. */
  readonly runningMode: SyncMode | null;
  readonly runningSince: string | null;
  /** How many selected notebook ids matched no notebook. A mistyped id is silent otherwise. */
  readonly unknownNotebookIds: number;
  /**
   * The two selection lists this service last reconciled, so a change can be diffed.
   *
   * A hash could say *that* they changed and not *which* notebook, and which notebook is
   * the whole point: widening the scan for every notebook when one was activated costs a
   * listing request per section of the account. `activeNotebookIdsSeen` is `null` for
   * "every mirrored notebook is active", exactly as in the selection document;
   * `mirroredNotebookIdsSeen` null is the one that means "never recorded", and it is the
   * field both `notebooksNeedingWideScan` and `selectionMatchesSeen` key on. A recorded
   * empty list reads back as `[]` rather than as null, because `readIdList` returns an
   * array for any array, so the two cannot be confused.
   */
  readonly mirroredNotebookIdsSeen: readonly string[] | null;
  readonly activeNotebookIdsSeen: readonly string[] | null;
  /**
   * Notebooks whose sections bypass the tier-1 timestamp cutoff until a run completes.
   *
   * Stored rather than held for one run, because a run is budget-bounded and may stop
   * with sections outstanding. It is cleared in the same place `sectionsScannedThrough`
   * advances — only when every candidate was visited.
   *
   * `readSyncState` reads a malformed value here as `[]`, which is the opposite choice
   * from `activeNotebookIds` in `readIdList`, and deliberately so. Failing open there
   * means extra Graph requests and failing closed would freeze the mirror; here the
   * asymmetry runs the other way — failing closed costs one pending widening, which is
   * some sections not re-listed until their timestamps move, and failing open would mean
   * inventing a notebook set to widen for out of a value nothing could parse. This field
   * is written only by this service, so a malformed one is a bug rather than a
   * hand-edit.
   */
  readonly wideScanNotebookIds: readonly string[];
  /** How many *active* ids matched no notebook. Same failure as `unknownNotebookIds`. */
  readonly unknownActiveNotebookIds: number;
}

/** The state a document that has never been written reads as. */
export function initialSyncState(): MirrorSyncState {
  return {
    schemaVersion: SCHEMA_VERSION,
    structureHash: null,
    sectionsScannedThrough: null,
    sectionRollUpTrusted: true,
    datetimeFilterSupported: true,
    backfillComplete: false,
    lastRunOutcome: null,
    lastRunGraphRequests: 0,
    lastRunPagesUpdated: 0,
    lastRunPagesDeleted: 0,
    sweepCursorSectionId: null,
    lastTreeFailureAt: null,
    runningMode: null,
    runningSince: null,
    unknownNotebookIds: 0,
    mirroredNotebookIdsSeen: null,
    activeNotebookIdsSeen: null,
    wideScanNotebookIds: [],
    unknownActiveNotebookIds: 0,
  };
}

/**
 * Read a stored sync-state document, defaulting every field it does not carry.
 *
 * Tolerant in the same direction as `readSelection` and for a related reason: a document
 * written by an older schema version, or half-written by an interrupted run, must not
 * stop the next run. A field of the wrong type falls back to its initial value rather
 * than throwing, and the defaults are all "do the safe, more expensive thing" —
 * `structureHash` null forces a structure write, a null watermark forces a full read.
 */
export function readSyncState(data: Record<string, unknown> | undefined): MirrorSyncState {
  const initial = initialSyncState();
  if (data === undefined) return initial;

  return {
    schemaVersion: numberOr(data['schemaVersion'], initial.schemaVersion),
    structureHash: stringOrNull(data['structureHash']),
    sectionsScannedThrough: stringOrNull(data['sectionsScannedThrough']),
    sectionRollUpTrusted: booleanOr(data['sectionRollUpTrusted'], initial.sectionRollUpTrusted),
    datetimeFilterSupported: booleanOr(
      data['datetimeFilterSupported'],
      initial.datetimeFilterSupported,
    ),
    backfillComplete: booleanOr(data['backfillComplete'], initial.backfillComplete),
    lastRunOutcome: syncOutcomeOrNull(data['lastRunOutcome']),
    lastRunGraphRequests: numberOr(data['lastRunGraphRequests'], 0),
    lastRunPagesUpdated: numberOr(data['lastRunPagesUpdated'], 0),
    lastRunPagesDeleted: numberOr(data['lastRunPagesDeleted'], 0),
    sweepCursorSectionId: stringOrNull(data['sweepCursorSectionId']),
    lastTreeFailureAt: stringOrNull(data['lastTreeFailureAt']),
    runningMode: syncModeOrNull(data['runningMode']),
    runningSince: stringOrNull(data['runningSince']),
    unknownNotebookIds: numberOr(data['unknownNotebookIds'], 0),
    mirroredNotebookIdsSeen: readIdList(data['mirroredNotebookIdsSeen']),
    activeNotebookIdsSeen: readIdList(data['activeNotebookIdsSeen']),
    wideScanNotebookIds: readIdList(data['wideScanNotebookIds']) ?? [],
    unknownActiveNotebookIds: numberOr(data['unknownActiveNotebookIds'], 0),
  };
}

/**
 * Is a held lease stale enough to take?
 *
 * A run that died mid-slice — an instance replaced, a request cut at Cloud Run's
 * timeout — leaves `runningMode` set with nothing to clear it. The lease therefore
 * expires on age rather than being held until released, and the expiry is longer than
 * any run's wall-clock budget so it can never fire on a run that is still working.
 */
export const LEASE_EXPIRY_MS = 900_000;

export function leaseIsHeld(
  state: MirrorSyncState,
  now: number,
  expiryMs: number = LEASE_EXPIRY_MS,
): boolean {
  if (state.runningMode === null || state.runningSince === null) return false;

  const since = Date.parse(state.runningSince);
  // An unparseable timestamp means nothing can ever clear the lease by age, which would
  // wedge the sync permanently. Treat it as expired.
  if (Number.isNaN(since)) return false;

  return now - since < expiryMs;
}

/**
 * How long a section's page-listing hold survives without being released.
 *
 * `endWrite` in ./write-tools.ts runs in a `finally`, so a hold outlives its write only
 * when the process stops between the two — Cloud Run cutting the request at 300 seconds,
 * or the instance being reclaimed. Nothing runs after that to lower the count, and a hold
 * that nothing can clear would send every listing for that section to Graph forever: one
 * request for a `list_pages`, and up to 61 for an unscoped `search_pages`.
 *
 * So the hold expires on age, and the expiry is longer than any request may live. Ten
 * minutes against Cloud Run's 300-second ceiling.
 */
export const LISTING_HOLD_EXPIRY_MS = 600_000;

/**
 * Is this section's page listing held back from being answered by the mirror?
 *
 * A count without a timestamp is not a document this code writes — both fields go in one
 * operation — and it is read as *expired* rather than as held, for the reason
 * `leaseIsHeld` reads an unparseable `runningSince` as expired: the alternative is a
 * state nothing can ever clear.
 */
export function listingIsHeld(
  section: Pick<MirrorSection, 'pendingWrites' | 'pendingWritesSince'>,
  now: number,
  expiryMs: number = LISTING_HOLD_EXPIRY_MS,
): boolean {
  const pending = section.pendingWrites;
  if (typeof pending !== 'number' || !(pending > 0)) return false;

  const since = section.pendingWritesSince;
  if (typeof since !== 'string') return false;

  const at = Date.parse(since);
  if (Number.isNaN(at)) return false;

  return now - at < expiryMs;
}

function stringOrNull(value: unknown): string | null {
  return typeof value === 'string' && value !== '' ? value : null;
}

function numberOr(value: unknown, fallback: number): number {
  return typeof value === 'number' && Number.isFinite(value) ? value : fallback;
}

function booleanOr(value: unknown, fallback: boolean): boolean {
  return typeof value === 'boolean' ? value : fallback;
}

function syncModeOrNull(value: unknown): SyncMode | null {
  return value === 'incremental' ||
    value === 'sweep' ||
    value === 'sweep-full' ||
    value === 'sweep-all'
    ? value
    : null;
}

function syncOutcomeOrNull(value: unknown): SyncOutcome | null {
  return value === 'complete' || value === 'budget-exhausted' || value === 'failed' ? value : null;
}
