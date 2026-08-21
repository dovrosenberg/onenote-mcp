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
  groupIdentity,
  htmlObjectName,
  htmlPlacement,
  initialSyncState,
  inkObjectName,
  inkmlObjectName,
  leaseIsHeld,
  listingIsHeld,
  LISTING_HOLD_EXPIRY_MS,
  notebookIdentity,
  overlapFrom,
  pageHasDrifted,
  planStructureWrite,
  isActive,
  notebooksNeedingWideScan,
  readSelection,
  readSyncState,
  sectionIdentity,
  selectionMatchesSeen,
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
    mirroredNotebookIdsSeen: ['nb-1', 'nb-2'],
    activeNotebookIdsSeen: ['nb-1'],
    wideScanNotebookIds: ['nb-2'],
    unknownActiveNotebookIds: 2,
  };

  assert.deepEqual(readSyncState(stored), stored);
});

test('notebooksNeedingWideScan names only what became mirrored or active', () => {
  const mirroredIds = ['a', 'b', 'c'];

  // Adding one notebook to the selection widens that one.
  assert.deepEqual(
    notebooksNeedingWideScan(
      { mirrored: ['a', 'b'], active: ['a', 'b'] },
      { notebookIds: ['a', 'b', 'c'], activeNotebookIds: ['a', 'b'] },
      mirroredIds,
    ),
    ['c'],
  );

  // Activating one widens that one.
  assert.deepEqual(
    notebooksNeedingWideScan(
      { mirrored: ['a', 'b'], active: ['a'] },
      { notebookIds: ['a', 'b'], activeNotebookIds: ['a', 'b'] },
      ['a', 'b'],
    ),
    ['b'],
  );

  // Removing widens nothing: nothing has to be caught up on.
  assert.deepEqual(
    notebooksNeedingWideScan(
      { mirrored: ['a', 'b'], active: ['a', 'b'] },
      { notebookIds: ['a'], activeNotebookIds: ['a'] },
      ['a'],
    ),
    [],
  );

  // `null` means every mirrored notebook is active, so a list becoming null activates
  // everything that was not already in the list.
  assert.deepEqual(
    notebooksNeedingWideScan(
      { mirrored: ['a', 'b', 'c'], active: ['a'] },
      { notebookIds: ['a', 'b', 'c'], activeNotebookIds: null },
      mirroredIds,
    ),
    ['b', 'c'],
  );

  // ...and null becoming a subset activates nothing new.
  assert.deepEqual(
    notebooksNeedingWideScan(
      { mirrored: ['a', 'b'], active: null },
      { notebookIds: ['a', 'b'], activeNotebookIds: ['a'] },
      ['a', 'b'],
    ),
    [],
  );

  // Never recorded — a state document written before these fields existed. Nothing was
  // skipped under it, so there is nothing to widen; recording is the whole job.
  assert.deepEqual(
    notebooksNeedingWideScan(
      { mirrored: null, active: null },
      { notebookIds: ['a', 'b'], activeNotebookIds: null },
      ['a', 'b'],
    ),
    [],
  );
});

test('a null active list resolves through the tree, not through the selection', () => {
  // The third parameter is the notebooks the tree actually returned as mirrored, which is
  // not `current.notebookIds`: a selected id naming no notebook is in the second list and
  // not the first. `activeNotebookIds: null` means "every mirrored notebook is active", so
  // resolving it through the selection would widen for a notebook that does not exist.
  const previous = { mirrored: ['a', 'b', 'ghost'], active: ['a'] };
  const current = { notebookIds: ['a', 'b', 'ghost'], activeNotebookIds: null };

  assert.deepEqual(notebooksNeedingWideScan(previous, current, ['a', 'b']), ['b']);
  // What the selection-resolved answer would have been, for contrast.
  assert.deepEqual(notebooksNeedingWideScan(previous, current, current.notebookIds), [
    'b',
    'ghost',
  ]);
});

test('a state document written before the selection fields reads as never recorded', () => {
  // What is actually deployed: `activeSelectionHash` and none of the fields that replaced
  // it. `mirroredNotebookIdsSeen` null is what makes the first run after this deploy
  // record the lists and widen nothing — reading the absent lists as "the selection was
  // empty" would widen every mirrored notebook on that run instead.
  const state = readSyncState({
    schemaVersion: 1,
    structureHash: 'abc123',
    sectionsScannedThrough: '2026-08-19T12:00:00Z',
    activeSelectionHash: 'def456',
  });

  assert.equal(state.mirroredNotebookIdsSeen, null);
  assert.equal(state.activeNotebookIdsSeen, null);
  assert.deepEqual(state.wideScanNotebookIds, []);
  assert.deepEqual(
    notebooksNeedingWideScan(
      { mirrored: state.mirroredNotebookIdsSeen, active: state.activeNotebookIdsSeen },
      { notebookIds: ['a', 'b'], activeNotebookIds: null },
      ['a', 'b'],
    ),
    [],
  );
});

test('selectionMatchesSeen asks a different question from the wide-scan diff', () => {
  // A removal widens nothing and still has to be recorded, so "nothing to widen" cannot
  // be the test for "nothing to write".
  const removed = { mirrored: ['a', 'b'], active: ['a', 'b'] };
  const now = { notebookIds: ['a', 'b'], activeNotebookIds: ['a'] };
  assert.deepEqual(notebooksNeedingWideScan(removed, now, ['a', 'b']), []);
  assert.equal(selectionMatchesSeen(removed, now), false);

  // Order is not a change; nothing was typed that was not already there.
  assert.equal(
    selectionMatchesSeen(
      { mirrored: ['a', 'b'], active: ['b', 'a'] },
      { notebookIds: ['b', 'a'], activeNotebookIds: ['a', 'b'] },
    ),
    true,
  );

  // `null` and a list naming every mirrored notebook are the same active set today and
  // different ones the moment a notebook is added, so they are not the same record.
  assert.equal(
    selectionMatchesSeen(
      { mirrored: ['a', 'b'], active: null },
      { notebookIds: ['a', 'b'], activeNotebookIds: ['a', 'b'] },
    ),
    false,
  );

  // Never recorded is never a match, whatever the selection says.
  assert.equal(
    selectionMatchesSeen({ mirrored: null, active: null }, { notebookIds: [], activeNotebookIds: null }),
    false,
  );
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

// ---------------------------------------------------------------------------
// Structure document identities
// ---------------------------------------------------------------------------

const NOTEBOOK_FIELDS = {
  id: 'nb-1',
  displayName: '2026',
  mirrored: true,
  sectionCount: 3,
  sectionGroupCount: 1,
  graphLastModifiedDateTime: '2026-08-19T11:00:00Z',
};

const GROUP_FIELDS = {
  id: 'grp-1',
  displayName: '062 - February',
  notebookId: 'nb-1',
  parentId: 'nb-1',
  parentKind: 'notebook' as const,
  path: '2026 / 062 - February',
  mirrored: true,
};

test('a section identity covers the tree fields and excludes the timestamp', () => {
  const base = {
    id: 'sec-1',
    displayName: 'Daily',
    notebookId: 'nb-1',
    parentId: 'nb-1',
    parentKind: 'notebook' as const,
    path: '2026 / Daily',
    mirrored: true,
    graphLastModifiedDateTime: '2026-08-19T11:00:00Z',
  };

  // The timestamp moves whenever anyone edits a page. Including it would rewrite every
  // section document on every structure change, which is the cost this whole task removes.
  assert.equal(
    sectionIdentity(base),
    sectionIdentity({ ...base, graphLastModifiedDateTime: '2026-08-19T23:59:00Z' }),
  );

  // Each of these is a real structural change and must produce a different identity.
  assert.notEqual(sectionIdentity(base), sectionIdentity({ ...base, displayName: 'Weekly' }));
  assert.notEqual(sectionIdentity(base), sectionIdentity({ ...base, mirrored: false }));
  assert.notEqual(sectionIdentity(base), sectionIdentity({ ...base, parentId: 'grp-1' }));
  assert.notEqual(sectionIdentity(base), sectionIdentity({ ...base, path: '2026 / Other' }));
  assert.notEqual(
    sectionIdentity(base),
    sectionIdentity({ ...base, parentKind: 'sectionGroup' as const }),
  );
});

test('a section group identity covers the tree fields and not childGroupsKnown', () => {
  // `childGroupsKnown` is sync-owned: the sweep sets it true after enumerating the nested
  // groups, and a structure write that carried it would put it back to false and send
  // every `list_sections` on that group to Graph until the next sweep re-learned them.
  const base = groupIdentity(GROUP_FIELDS);
  assert.notEqual(base, groupIdentity({ ...GROUP_FIELDS, displayName: 'March' }));
  assert.notEqual(base, groupIdentity({ ...GROUP_FIELDS, mirrored: false }));
  assert.notEqual(base, groupIdentity({ ...GROUP_FIELDS, parentId: 'grp-0' }));
  assert.notEqual(base, groupIdentity({ ...GROUP_FIELDS, path: '2026 / Other' }));

  // A group carries no timestamp at all, so nothing here can be confused with one.
  assert.equal(base, groupIdentity({ ...GROUP_FIELDS }));
});

test('a notebook identity covers the counts and excludes the timestamp', () => {
  assert.equal(
    notebookIdentity(NOTEBOOK_FIELDS),
    notebookIdentity({ ...NOTEBOOK_FIELDS, graphLastModifiedDateTime: '2099-01-01T00:00:00Z' }),
  );

  assert.notEqual(
    notebookIdentity(NOTEBOOK_FIELDS),
    notebookIdentity({ ...NOTEBOOK_FIELDS, displayName: '2027' }),
  );
  assert.notEqual(
    notebookIdentity(NOTEBOOK_FIELDS),
    notebookIdentity({ ...NOTEBOOK_FIELDS, mirrored: false }),
  );
  // The counts are what a `list_notebooks` answer prints, so a section added to a notebook
  // has to reach the stored document even though nothing else about the notebook moved.
  assert.notEqual(
    notebookIdentity(NOTEBOOK_FIELDS),
    notebookIdentity({ ...NOTEBOOK_FIELDS, sectionCount: 4 }),
  );
  assert.notEqual(
    notebookIdentity(NOTEBOOK_FIELDS),
    notebookIdentity({ ...NOTEBOOK_FIELDS, sectionGroupCount: 0 }),
  );
});

const SEPARATOR_FIELDS = {
  id: 'sec-1',
  displayName: 'a b',
  notebookId: 'c',
  parentId: 'nb-1',
  parentKind: 'notebook' as const,
  path: '2026 / Daily',
  mirrored: true,
  graphLastModifiedDateTime: null,
};

test('a display name holding the separator cannot forge another split of the fields', () => {
  // The failure a space separator has: `{displayName: 'a b', notebookId: 'c'}` and
  // `{displayName: 'a', notebookId: 'b c'}` join to the same string, so a rename that
  // moved a space across a field boundary would be invisible and the document would never
  // be rewritten. A OneNote display name can hold a space; it cannot hold a NUL.
  const left = sectionIdentity({ ...SEPARATOR_FIELDS, displayName: 'a b', notebookId: 'c' });
  const right = sectionIdentity({ ...SEPARATOR_FIELDS, displayName: 'a', notebookId: 'b c' });
  assert.notEqual(left, right);
});

// ---------------------------------------------------------------------------
// planStructureWrite
//
// This is `#replaceCollection`'s whole judgement, extracted so a test can reach it:
// src/mirror-store.ts needs a Firestore backend and has none here, so a branch left in
// that file is a branch nothing checks. Every test below is a mutation that used to pass.
// ---------------------------------------------------------------------------

const SECTION_DEFAULTS = { pagesSyncedThrough: null, pageCount: 0 };

function incoming(id: string, identity: string): { id: string; identity: string } {
  return { id, identity };
}

test('a document whose stored identity matches is not written at all', () => {
  // The failure without it: renaming one section moves the tree hash, and every one of the
  // account's 568 section documents is rewritten with the tree's idea of it.
  const plan = planStructureWrite(
    new Map([
      ['sec-1', 'same'],
      ['sec-2', 'same'],
    ]),
    [incoming('sec-1', 'same'), incoming('sec-2', 'moved')],
    SECTION_DEFAULTS,
  );

  assert.deepEqual(plan.writes.map((w) => w.documentId), ['sec-2']);
  assert.deepEqual(plan.deletes, []);
});

test('a document that already existed is merged, and carries no creation default', () => {
  // The defaults are the reset this whole split removes. `pagesSyncedThrough` absent from
  // the written fields is what makes the stored watermark survive a `set(…, {merge:true})`.
  const plan = planStructureWrite(
    new Map([['sec-1', 'old']]),
    [incoming('sec-1', 'new')],
    SECTION_DEFAULTS,
  );

  const write = plan.writes[0];
  assert.equal(write?.created, false);
  assert.equal('pagesSyncedThrough' in (write?.fields ?? {}), false);
  assert.equal('pageCount' in (write?.fields ?? {}), false);
  assert.deepEqual(write?.fields, { id: 'sec-1', identity: 'new' });
});

test('a document that was not there is created with the defaults', () => {
  const plan = planStructureWrite(new Map(), [incoming('sec-1', 'new')], SECTION_DEFAULTS);

  const write = plan.writes[0];
  assert.equal(write?.created, true);
  // Not decoration: `listSectionsToSync` orders by this field, and an `orderBy` drops a
  // document that lacks it — so a section created without it is never synced and nothing
  // says so.
  // Strict, because `assert.equal` is `==`: a default of `'0'` for `pageCount`, or of `0`
  // for a field Firestore then orders on, would pass a loose comparison.
  assert.strictEqual(write?.fields['pagesSyncedThrough'], null);
  assert.strictEqual(write?.fields['pageCount'], 0);
});

test('a stored document with no identity is merged, not created', () => {
  // What every document looks like on the first run after this field was introduced. It
  // exists, so it must not take the create branch — that would null every watermark in the
  // account, which is the failure the identity exists to prevent — and it does not match,
  // so it is written once.
  const plan = planStructureWrite(
    new Map([['sec-1', undefined]]),
    [incoming('sec-1', 'new')],
    SECTION_DEFAULTS,
  );

  assert.equal(plan.writes[0]?.created, false);
  assert.equal('pagesSyncedThrough' in (plan.writes[0]?.fields ?? {}), false);
});

test('a document the tree no longer names is deleted, and the skip does not reach it', () => {
  // Deletion is by absence from the incoming set, which is a fact rather than an
  // inference: the caller has just read the whole tree in one request.
  const plan = planStructureWrite(
    new Map([
      ['sec-1', 'same'],
      ['gone', 'whatever'],
    ]),
    [incoming('sec-1', 'same')],
    SECTION_DEFAULTS,
  );

  assert.deepEqual(plan.writes, []);
  assert.deepEqual(plan.deletes, ['gone']);
});

test('a plan keys its writes by the Firestore document id, not the Graph id', () => {
  // A OneNote id can contain a `/`, which Firestore forbids in a document id. The stored
  // map is keyed by the encoded form, so the plan has to encode before it compares — a
  // comparison against the raw id would find nothing stored and create every document.
  const graphId = '0-abc/def!1-ghi';
  const documentId = encodeMirrorId(graphId);
  assert.notEqual(documentId, graphId);

  const plan = planStructureWrite(
    new Map([[documentId, 'same']]),
    [incoming(graphId, 'same')],
    SECTION_DEFAULTS,
  );

  assert.deepEqual(plan.writes, []);
  assert.deepEqual(plan.deletes, [], 'and it is not read as a document the tree dropped');
});

// ---------------------------------------------------------------------------
// Content drift
// ---------------------------------------------------------------------------

test('pageHasDrifted compares Graph against Graph, and only for a present copy', () => {
  const stored = {
    id: 'p1',
    title: 'Page',
    lastModifiedDateTime: '2026-08-19T12:00:00Z',
    contentState: 'present' as const,
  };
  const live = { id: 'p1', title: 'Page', lastModifiedDateTime: '2026-08-19T12:00:00Z' };

  assert.equal(pageHasDrifted(stored, live), false, 'identical stamps are not drift');
  assert.equal(
    pageHasDrifted(stored, { ...live, lastModifiedDateTime: '2026-08-19T12:00:01Z' }),
    true,
    "one second later is drift \u2014 both sides are Graph's own string",
  );
  assert.equal(
    pageHasDrifted({ ...stored, contentState: 'stale' }, { ...live, lastModifiedDateTime: 'x' }),
    false,
    'a copy already stale has nothing to invalidate',
  );
  assert.equal(
    pageHasDrifted({ ...stored, contentState: 'missing' }, { ...live, lastModifiedDateTime: 'x' }),
    false,
    'nor has one recorded as missing',
  );
  assert.equal(
    pageHasDrifted(stored, { ...live, lastModifiedDateTime: '' }),
    false,
    'an absent timestamp is not evidence of a change',
  );
});
