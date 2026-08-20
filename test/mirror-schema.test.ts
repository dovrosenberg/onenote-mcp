// Every decision in the mirror that could be wrong, tested where it runs without a
// backend.
//
// This file exists because src/mirror-store.ts and src/mirror-blobs.ts cannot be tested
// on this machine at all — no Firestore emulator, and CLAUDE.md rules out an in-memory
// fake because the behaviour at stake there is transaction retry and
// FieldValue.serverTimestamp(). So the rule is that anything a person could get wrong
// lives in src/mirror-schema.ts and is asserted here, and the two I/O modules are left
// holding only calls.
//
// What no test here covers is whether Firestore accepts what these functions produce.
// The document-id rules, the 1500-byte id limit and the 1 MiB document limit are read
// off Google's documentation, not measured, and only a live write settles them.

import test from 'node:test';
import assert from 'node:assert/strict';

import {
  HTML_INLINE_LIMIT_BYTES,
  LEASE_EXPIRY_MS,
  MirrorSchemaError,
  SCHEMA_VERSION,
  WATERMARK_OVERLAP_MS,
  encodeMirrorId,
  htmlObjectName,
  htmlPlacement,
  initialSyncState,
  inkObjectName,
  inkmlObjectName,
  leaseIsHeld,
  listingIsHeld,
  LISTING_HOLD_EXPIRY_MS,
  overlapFrom,
  isActive,
  readSelection,
  readSyncState,
  utf8Bytes,
} from '../src/mirror-schema.ts';

// A real OneNote page id, shape-wise: two GUID-ish halves joined by an exclamation mark.
const REAL_ID = '0-a1b2c3d4e5f67890abcdef1234567890!1-A1B2C3D4E5F67890!123';

// ---------------------------------------------------------------------------
// Document ids
// ---------------------------------------------------------------------------

test('a real Graph id survives encoding unchanged except where it must not', () => {
  const encoded = encodeMirrorId(REAL_ID);
  assert.equal(encoded.includes('/'), false, 'Firestore forbids a slash in a document id');
  assert.equal(encoded, encodeURIComponent(REAL_ID));
});

test('a Graph id containing a slash produces a legal Firestore id', () => {
  // This is the whole reason for the encoding. A slash in a document id is not an error
  // in Firestore, it is a path separator, so it would silently write to the wrong place.
  const encoded = encodeMirrorId('0-abc/def!1-xyz');
  assert.equal(encoded.includes('/'), false);
  assert.equal(encoded, '0-abc%2Fdef!1-xyz');
});

test('the reserved Firestore ids are rejected rather than encoded around', () => {
  // None of these can come out of Graph. They are checked because a value that produced
  // one would mean something upstream is badly wrong, and writing it would either fail
  // opaquely or land somewhere unexpected.
  for (const id of ['.', '..', '__proto__', '__x__']) {
    assert.throws(() => encodeMirrorId(id), MirrorSchemaError, id);
  }
  assert.throws(() => encodeMirrorId(''), MirrorSchemaError);
});

test('an id over the Firestore document-id limit is refused, measured in UTF-8 bytes', () => {
  // Percent-encoding inflates, so the check is on the encoded byte length rather than
  // the input's character count.
  assert.doesNotThrow(() => encodeMirrorId('a'.repeat(1500)));
  assert.throws(() => encodeMirrorId('a'.repeat(1501)), MirrorSchemaError);
  // Each of these encodes to three characters, so 500 of them is exactly at the limit.
  assert.doesNotThrow(() => encodeMirrorId('/'.repeat(500)));
  assert.throws(() => encodeMirrorId('/'.repeat(501)), MirrorSchemaError);
});

test('the three object names are distinct and keyed by the encoded page id', () => {
  const encoded = encodeMirrorId(REAL_ID);

  assert.equal(inkObjectName(REAL_ID), `ink/${encoded}.png`);
  assert.equal(inkmlObjectName(REAL_ID), `inkml/${encoded}.xml`);
  assert.equal(htmlObjectName(REAL_ID), `html/${encoded}.html`);

  // deleteForPage builds all three from the page id with no lookup, so they must not
  // collide across kinds.
  const names = new Set([inkObjectName(REAL_ID), inkmlObjectName(REAL_ID), htmlObjectName(REAL_ID)]);
  assert.equal(names.size, 3);
});

// ---------------------------------------------------------------------------
// The hand-edited selection
// ---------------------------------------------------------------------------

test('a well-formed selection reads back in order, deduplicated', () => {
  assert.deepEqual(
    readSelection({ notebookIds: ['nb-1', 'nb-2', 'nb-1'] }),
    { notebookIds: ['nb-1', 'nb-2'], activeNotebookIds: null },
  );
});

test('anything a person could leave in the selection document means "mirror nothing"', () => {
  // A human edits this by hand in the Firestore console. A half-finished edit must not
  // take the sync down; it should mean no pages are mirrored this run, which is visible
  // in the run report and recoverable by fixing the document.
  for (const data of [
    undefined,
    {},
    { notebookIds: null },
    { notebookIds: 'nb-1' },
    { notebookIds: 42 },
    { notebookIds: {} },
    { notebookIds: [] },
    { note: 'still deciding' },
  ]) {
    assert.deepEqual(readSelection(data as Record<string, unknown> | undefined), {
      notebookIds: [],
      activeNotebookIds: null,
    });
  }
});

test('one bad entry does not discard the good ones beside it', () => {
  assert.deepEqual(
    readSelection({ notebookIds: ['nb-1', 42, null, '', '   ', { id: 'nb-x' }, 'nb-2'] }),
    { notebookIds: ['nb-1', 'nb-2'], activeNotebookIds: null },
  );
});

test('selection entries are trimmed, because a console paste carries whitespace', () => {
  assert.deepEqual(readSelection({ notebookIds: ['  nb-1  ', 'nb-1'] }), {
    notebookIds: ['nb-1'],
    activeNotebookIds: null,
  });
});

test('an absent or malformed active list means every selected notebook is active', () => {
  // Null rather than []: the two mean opposite things, and failing open is the safe
  // direction. A malformed value that read as "none active" would freeze the mirror with
  // nothing saying so; reading it as "all active" only costs Graph requests.
  for (const data of [
    { notebookIds: ['nb-1'] },
    { notebookIds: ['nb-1'], activeNotebookIds: null },
    { notebookIds: ['nb-1'], activeNotebookIds: 'nb-1' },
    { notebookIds: ['nb-1'], activeNotebookIds: 42 },
    { notebookIds: ['nb-1'], activeNotebookIds: {} },
  ]) {
    assert.equal(readSelection(data as Record<string, unknown>).activeNotebookIds, null);
  }
});

test('an empty active list is a deliberate edit, not a missing one', () => {
  // "Freeze everything" is a state an operator may legitimately want, and it has to be
  // distinguishable from a document that has never heard of the field.
  assert.deepEqual(readSelection({ notebookIds: ['nb-1'], activeNotebookIds: [] }), {
    notebookIds: ['nb-1'],
    activeNotebookIds: [],
  });
});

test('the active list is read by the same rules as the selection', () => {
  assert.deepEqual(
    readSelection({
      notebookIds: ['nb-1', 'nb-2'],
      activeNotebookIds: ['  nb-2  ', 'nb-2', 42, '', null, 'nb-1'],
    }),
    { notebookIds: ['nb-1', 'nb-2'], activeNotebookIds: ['nb-2', 'nb-1'] },
  );
});

test('isActive is true for everything when no active set was named', () => {
  const selection = readSelection({ notebookIds: ['nb-1', 'nb-2'] });
  assert.equal(isActive(selection, 'nb-1'), true);
  // It does not consult notebookIds: a notebook that is not mirrored never reaches a
  // code path that asks, and adding the check would make the answer depend on two fields.
  assert.equal(isActive(selection, 'nb-nowhere'), true);
});

test('isActive is true only for the listed notebooks when a set was named', () => {
  const selection = readSelection({
    notebookIds: ['nb-1', 'nb-2'],
    activeNotebookIds: ['nb-2'],
  });
  assert.equal(isActive(selection, 'nb-2'), true);
  assert.equal(isActive(selection, 'nb-1'), false);

  assert.equal(isActive(readSelection({ notebookIds: ['nb-1'], activeNotebookIds: [] }), 'nb-1'), false);
});

// ---------------------------------------------------------------------------
// Watermarks
// ---------------------------------------------------------------------------

test('overlapFrom subtracts the overlap and never returns a later instant', () => {
  assert.equal(
    overlapFrom('2026-08-19T12:00:00.000Z'),
    new Date(Date.parse('2026-08-19T12:00:00.000Z') - WATERMARK_OVERLAP_MS).toISOString(),
  );
  assert.equal(overlapFrom('2026-08-19T12:00:00.000Z', 0), '2026-08-19T12:00:00.000Z');

  // A negative overlap would otherwise skip a window, which loses edits permanently.
  assert.equal(overlapFrom('2026-08-19T12:00:00.000Z', -60_000), '2026-08-19T12:00:00.000Z');
});

test('a never-synced or unreadable watermark asks for everything', () => {
  const epoch = new Date(0).toISOString();

  assert.equal(overlapFrom(null), epoch);
  // The value is written by this service, so an unparseable one means something already
  // went wrong. A full re-read is the recovery; refusing to sync would be worse.
  assert.equal(overlapFrom('not a date'), epoch);
  assert.equal(overlapFrom(''), epoch);
});

test('the overlap is an hour, and it is subtracted rather than added', () => {
  assert.equal(WATERMARK_OVERLAP_MS, 3_600_000);
  assert.ok(
    Date.parse(overlapFrom('2026-08-19T12:00:00.000Z')) <
      Date.parse('2026-08-19T12:00:00.000Z'),
  );
});

// ---------------------------------------------------------------------------
// Page HTML placement
// ---------------------------------------------------------------------------

test('HTML spills to GCS strictly above the limit', () => {
  assert.equal(htmlPlacement(0), 'firestore');
  assert.equal(htmlPlacement(HTML_INLINE_LIMIT_BYTES - 1), 'firestore');
  assert.equal(htmlPlacement(HTML_INLINE_LIMIT_BYTES), 'firestore');
  assert.equal(htmlPlacement(HTML_INLINE_LIMIT_BYTES + 1), 'gcs');
});

test('the limit leaves real headroom under Firestore 1 MiB document cap', () => {
  // The margin is the point. The byte length of the HTML is not the size of the
  // document — there are other fields, and index entries count against the same limit.
  assert.ok(HTML_INLINE_LIMIT_BYTES < 1_048_576 * 0.7);
});

test('sizes are measured in UTF-8 bytes, not string length', () => {
  // A page of handwriting notes is full of characters that are more than one byte, and
  // every Firestore and GCS limit is in bytes. Measuring length would under-count and
  // spill too late.
  assert.equal(utf8Bytes('abc'), 3);
  assert.equal(utf8Bytes('é'), 2);
  assert.equal(utf8Bytes('—'), 3);
  assert.equal(utf8Bytes('🖊'), 4);
  assert.ok(utf8Bytes('🖊') > '🖊'.length);
});

// ---------------------------------------------------------------------------
// Sync state
// ---------------------------------------------------------------------------

test('an absent sync-state document reads as the initial state', () => {
  assert.deepEqual(readSyncState(undefined), initialSyncState());
});

test('the initial state defaults to what the probe measured, and to doing more work', () => {
  const initial = initialSyncState();

  // Measured 2026-08-19: a page create, edit and delete each move the section timestamp,
  // and it does not move otherwise. See api-overview.md.
  assert.equal(initial.sectionRollUpTrusted, true);
  assert.equal(initial.datetimeFilterSupported, true);

  // Every other default is the safe, more expensive branch: a null hash forces a
  // structure write, a null watermark forces a full read, an incomplete backfill keeps
  // the sync running.
  assert.equal(initial.structureHash, null);
  assert.equal(initial.sectionsScannedThrough, null);
  assert.equal(initial.backfillComplete, false);
  assert.equal(initial.runningMode, null);
  assert.equal(initial.schemaVersion, SCHEMA_VERSION);
});

test('a field of the wrong type falls back rather than throwing', () => {
  // A document half-written by an interrupted run, or written by an older schema
  // version, must not stop the next run.
  const state = readSyncState({
    schemaVersion: 'one',
    structureHash: 42,
    sectionRollUpTrusted: 'true',
    datetimeFilterSupported: null,
    backfillComplete: 1,
    lastRunOutcome: 'exploded',
    lastRunGraphRequests: Number.NaN,
    runningMode: 'sideways',
    unknownNotebookIds: [],
  });

  assert.deepEqual(state, initialSyncState());
});

test('a well-formed sync-state document reads back verbatim', () => {
  const stored = {
    schemaVersion: 1,
    structureHash: 'abc123',
    sectionsScannedThrough: '2026-08-19T12:00:00Z',
    sectionRollUpTrusted: false,
    datetimeFilterSupported: false,
    backfillComplete: true,
    lastRunOutcome: 'budget-exhausted',
    lastRunGraphRequests: 120,
    lastRunPagesUpdated: 14,
    lastRunPagesDeleted: 2,
    sweepCursorSectionId: 'sec-9',
    lastTreeFailureAt: '2026-08-19T11:00:00Z',
    runningMode: 'sweep',
    runningSince: '2026-08-19T12:05:00Z',
    unknownNotebookIds: 1,
    activeSelectionHash: 'def456',
    unknownActiveNotebookIds: 2,
  };

  assert.deepEqual(readSyncState(stored), stored);
});

test('an empty string reads as null, so a cleared field is not a cursor', () => {
  // Clearing a field in the Firestore console leaves an empty string. Reading that as a
  // cursor id would resume a sweep at a section that does not exist.
  const state = readSyncState({ sweepCursorSectionId: '', structureHash: '' });
  assert.equal(state.sweepCursorSectionId, null);
  assert.equal(state.structureHash, null);
});

// ---------------------------------------------------------------------------
// The lease
// ---------------------------------------------------------------------------

test('a lease is held only while it is younger than the expiry', () => {
  const state = {
    ...initialSyncState(),
    runningMode: 'incremental' as const,
    runningSince: '2026-08-19T12:00:00.000Z',
  };
  const started = Date.parse('2026-08-19T12:00:00.000Z');

  assert.equal(leaseIsHeld(state, started), true);
  assert.equal(leaseIsHeld(state, started + LEASE_EXPIRY_MS - 1), true);
  assert.equal(leaseIsHeld(state, started + LEASE_EXPIRY_MS), false);
});

test('a section listing is held while a write against it is unfinished', () => {
  const since = '2026-08-19T12:00:00.000Z';
  const at = Date.parse(since);

  assert.equal(listingIsHeld({ pendingWrites: 1, pendingWritesSince: since }, at), true);
  assert.equal(listingIsHeld({ pendingWrites: 2, pendingWritesSince: since }, at), true);
  // A count rather than a flag: two writes against one section can overlap, and the
  // first to finish must not clear the second's hold.
  assert.equal(listingIsHeld({ pendingWrites: 0, pendingWritesSince: since }, at), false);
  assert.equal(listingIsHeld({}, at), false, 'a section nothing is writing to');
});

test('a listing hold expires on age, so a dead process cannot wedge a section', () => {
  const since = '2026-08-19T12:00:00.000Z';
  const at = Date.parse(since);
  const held = { pendingWrites: 1, pendingWritesSince: since };

  assert.equal(listingIsHeld(held, at + LISTING_HOLD_EXPIRY_MS - 1), true);
  assert.equal(listingIsHeld(held, at + LISTING_HOLD_EXPIRY_MS), false);
  // Longer than any request may live: Cloud Run cuts one at 300 seconds, so a hold older
  // than this belongs to a process that is gone.
  assert.ok(LISTING_HOLD_EXPIRY_MS > 300_000);
});

test('a hold with no timestamp reads as expired, not as held forever', () => {
  // Both fields go in one write, so this is not a document this code produces. The
  // direction matters: reading it as held would leave a section nothing could ever clear,
  // sending every listing for it to Graph permanently.
  const at = Date.parse('2026-08-19T12:00:00.000Z');

  assert.equal(listingIsHeld({ pendingWrites: 1 }, at), false);
  assert.equal(listingIsHeld({ pendingWrites: 1, pendingWritesSince: 'not a date' }, at), false);
});

test('the lease expiry outlasts any run, so it cannot fire on one still working', () => {
  // SYNC_TIME_BUDGET_MS is 240s and Cloud Run's request timeout is 300s. A 15-minute
  // expiry is comfortably past both, so the only thing it ever clears is a run whose
  // instance died.
  assert.ok(LEASE_EXPIRY_MS > 300_000);
});

test('a lease with no mode, or an unreadable timestamp, is not held', () => {
  const now = Date.parse('2026-08-19T12:00:00.000Z');

  assert.equal(leaseIsHeld(initialSyncState(), now), false);
  assert.equal(
    leaseIsHeld({ ...initialSyncState(), runningMode: 'sweep', runningSince: null }, now),
    false,
  );
  // Nothing could ever clear this one by age, so treating it as held would wedge the
  // sync permanently.
  assert.equal(
    leaseIsHeld({ ...initialSyncState(), runningMode: 'sweep', runningSince: 'nonsense' }, now),
    false,
  );
});
