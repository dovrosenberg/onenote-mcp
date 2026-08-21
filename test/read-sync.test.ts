// The inline refresh every read tool runs before it reads the mirror.
//
// Two things are tested here and nothing else touches either: which sync reports license
// the claim "the mirror equals OneNote", and how often a refresh is allowed to run.
//
// The second is the one with a cost attached. Without an interval rule a conversation of
// thirty tool calls runs thirty syncs, and OneNote allows 400 requests an hour across
// everything this server does — so every assertion below about a refresh that did *not*
// run is an assertion about the hourly budget. The clock is injected rather than waited
// on, the way `SyncDeps.now` is in test/mirror-sync.test.ts.

import test from 'node:test';
import assert from 'node:assert/strict';

import { setEventSink } from '../src/logging.ts';
import { MirrorLeaseHeldError } from '../src/mirror-store.ts';
import type { SyncReport } from '../src/mirror-sync.ts';
import {
  INLINE_SYNC_MIN_INTERVAL_MS,
  INLINE_SYNC_RETRY_INTERVAL_MS,
  createReadSync,
  freshnessOf,
} from '../src/read-sync.ts';

setEventSink(() => {});

/** A run that reached the end with nothing outstanding. Fields are overridden per test. */
function report(overrides: Partial<SyncReport> = {}): SyncReport {
  return {
    mode: 'incremental',
    outcome: 'complete',
    done: true,
    graphRequests: 1,
    sectionsVisited: 0,
    pagesUpdated: 0,
    pagesDeleted: 0,
    pagesFailed: 0,
    unknownNotebookIds: 0,
    unknownActiveNotebookIds: 0,
    sectionsSkippedInactive: 0,
    treeRead: true,
    durationMs: 120,
    ...overrides,
  };
}

/** A clock a test moves by hand. */
function clock(): { now: () => number; advance: (ms: number) => void } {
  let at = 1_000_000;
  return {
    now: () => at,
    advance: (ms) => {
      at += ms;
    },
  };
}

// ---------------------------------------------------------------------------
// freshnessOf
// ---------------------------------------------------------------------------

test('a run that finished everything licenses the OneNote claim', () => {
  assert.equal(freshnessOf(report()), 'current');
});

test('each way a run can return with work left behind reads as behind', () => {
  // Four separate ways, and each of them leaves part of the mirror holding whatever it
  // held before. A single `outcome === complete` check would pass three of them.
  const partial: Array<readonly [string, Partial<SyncReport>]> = [
    ['the run failed outright', { outcome: 'failed', done: false }],
    ['the budget stopped it with sections outstanding', { outcome: 'budget-exhausted', done: false }],
    // Not fatal to a sync — it carries on against the structure already in Firestore —
    // but a notebook or section created since the last good tree read is absent from it.
    ['the expanded-tree read failed', { treeRead: false }],
    // The page keeps its old stored copy, which may be several edits behind.
    ['a page content fetch failed', { pagesFailed: 1 }],
  ];

  for (const [why, overrides] of partial) {
    assert.equal(freshnessOf(report(overrides)), 'behind', why);
  }
});

// ---------------------------------------------------------------------------
// createReadSync
// ---------------------------------------------------------------------------

test('a finished refresh is not repeated inside the interval', () => {
  // The whole cost control. A tool call every few seconds must not be a sync every few
  // seconds.
  const { now, advance } = clock();
  let runs = 0;
  const sync = createReadSync(() => {
    runs += 1;
    return Promise.resolve(report());
  }, now);

  return (async () => {
    assert.equal(await sync.refresh(), 'current');
    advance(INLINE_SYNC_MIN_INTERVAL_MS - 1);
    assert.equal(await sync.refresh(), 'current', 'the held answer, not a new run');
    assert.equal(runs, 1);

    advance(2);
    assert.equal(await sync.refresh(), 'current');
    assert.equal(runs, 2, 'past the interval it runs again');
  })();
});

test('a refresh that did not finish is suppressed for longer, not shorter', async () => {
  // A budget that ran out, a lease that was held, a Graph that refused — none of which
  // thirty seconds makes better. Retrying at the same cadence would spend the hourly
  // budget learning the same thing over and over, which is the failure the sync's own
  // budget exists to prevent.
  const { now, advance } = clock();
  let runs = 0;
  const sync = createReadSync(() => {
    runs += 1;
    return Promise.resolve(report({ outcome: 'budget-exhausted', done: false }));
  }, now);

  assert.equal(await sync.refresh(), 'behind');
  assert.equal(runs, 1);

  // Past the interval a *finished* refresh would use, and still suppressed.
  advance(INLINE_SYNC_MIN_INTERVAL_MS + 1);
  assert.equal(await sync.refresh(), 'behind');
  assert.equal(runs, 1);

  advance(INLINE_SYNC_RETRY_INTERVAL_MS);
  assert.equal(await sync.refresh(), 'behind');
  assert.equal(runs, 2);
});

test('two callers arriving together share one run', async () => {
  // The second would be refused by the sync lease anyway. Sharing gives it the first
  // one's answer rather than a `behind` it did not earn.
  const { now } = clock();
  let runs = 0;
  let release: (() => void) | undefined;
  const gate = new Promise<void>((resolve) => {
    release = resolve;
  });

  const sync = createReadSync(async () => {
    runs += 1;
    await gate;
    return report();
  }, now);

  const both = Promise.all([sync.refresh(), sync.refresh()]);
  release?.();

  assert.deepEqual(await both, ['current', 'current']);
  assert.equal(runs, 1);
});

test('a refresh that throws answers behind and never reaches the caller', async () => {
  // A failed refresh must not fail the read. The tool still has its Graph fallback, and
  // the only consequence is the weaker label.
  const { now, advance } = clock();
  const sync = createReadSync(() => Promise.reject(new Error('firestore down')), now);

  assert.equal(await sync.refresh(), 'behind');

  // And the failure is subject to the same retry interval, so a persistent outage costs
  // one attempt every five minutes rather than one per tool call.
  advance(INLINE_SYNC_MIN_INTERVAL_MS + 1);
  assert.equal(await sync.refresh(), 'behind');
});

test('the scheduler holding the lease is a behind, not an error', async () => {
  // An ordinary overlap: the scheduled run is doing the work this refresh wanted done.
  const { now } = clock();
  const sync = createReadSync(() => Promise.reject(new MirrorLeaseHeldError('sweep')), now);

  assert.equal(await sync.refresh(), 'behind');
});

test('a failed run does not wedge the next attempt', async () => {
  // The in-flight promise has to be cleared however the run ends, or one rejection would
  // leave every later refresh awaiting a promise that already settled.
  const { now, advance } = clock();
  let runs = 0;
  const sync = createReadSync(() => {
    runs += 1;
    return runs === 1 ? Promise.reject(new Error('once')) : Promise.resolve(report());
  }, now);

  assert.equal(await sync.refresh(), 'behind');
  advance(INLINE_SYNC_RETRY_INTERVAL_MS);
  assert.equal(await sync.refresh(), 'current');
  assert.equal(runs, 2);
});
