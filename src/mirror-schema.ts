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
 */
export function readSelection(data: Record<string, unknown> | undefined): NotebookSelection {
  const raw = data?.['notebookIds'];
  if (!Array.isArray(raw)) return { notebookIds: [] };

  const seen = new Set<string>();
  const notebookIds: string[] = [];
  for (const entry of raw) {
    if (typeof entry !== 'string') continue;
    const trimmed = entry.trim();
    if (trimmed === '' || seen.has(trimmed)) continue;
    seen.add(trimmed);
    notebookIds.push(trimmed);
  }

  return { notebookIds };
}

// ---------------------------------------------------------------------------
// Watermarks
// ---------------------------------------------------------------------------

/**
 * How far back a re-query reaches beyond the stored watermark.
 *
 * Graph's clock and this service's clock are not the same clock, and the watermark is
 * the time a pass *started* rather than the newest page it saw. A re-fetch of an
 * unchanged page is idempotent and nearly free — `contentHash` short-circuits it — and a
 * missed page is neither.
 */
export const WATERMARK_OVERLAP_MS = 3_600_000;

/**
 * The instant a re-query should start from, given a watermark.
 *
 * A null watermark means the section has never been synced, which asks for everything:
 * the epoch. Overlap is subtracted rather than added, and the result is clamped to never
 * exceed the watermark itself, so a negative or absurd overlap cannot skip a window.
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

export type SyncMode = 'incremental' | 'sweep' | 'sweep-full';
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
  return value === 'incremental' || value === 'sweep' || value === 'sweep-full' ? value : null;
}

function syncOutcomeOrNull(value: unknown): SyncOutcome | null {
  return value === 'complete' || value === 'budget-exhausted' || value === 'failed' ? value : null;
}
