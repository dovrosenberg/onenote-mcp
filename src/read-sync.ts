// The inline sync every read tool attempts before it reads the mirror.
//
// The read path is three steps: refresh the mirror, read the mirror, and say whether the
// refresh finished. A refresh that finished means the mirror equals OneNote as of a few
// seconds ago, so the answer is reported as `source: "onenote"`; a refresh that did not
// means the copy may be behind, and the answer says `source: "mirror"` instead. That is
// the whole reason this module exists — not to make reads faster, but to make the source
// label an honest claim rather than a guess about how long ago the scheduler last ran.
//
// **It is the incremental sync, not a sweep.** `runIncremental` costs one Graph request
// when nothing has changed — the expanded tree — plus one per changed section and one per
// changed page. A sweep enumerates every page of a section, section by section, and a
// full sweep visits every selected section regardless of timestamps; either would take
// tens of seconds and tens of requests on a tool call a human is waiting on. What stays
// the sweep's job, on the scheduler, is reconciling a section's stored pages against its
// live listing: deleting pages Graph no longer returns, discovering pages moved in under a
// stamp no watermark reaches, and re-fetching a page whose stored stamp or title disagrees
// with the listing. Nothing there marks a page — a mark deletes the page-content document
// and nothing re-fetches a marked page, so the sweep's answer to a disagreement is always
// a re-fetch.
//
// **Everything here is bounded twice, and the second bound is the one that matters.**
// The per-run budget caps what one refresh may spend. The interval caps how often a
// refresh may run at all, and without it a conversation of thirty tool calls would run
// thirty syncs. The arithmetic against OneNote's 400-requests-per-hour limit:
//
// | Situation | Cost |
// |---|---|
// | A quiet account, a call every few seconds | 1 request per 30s, so 120/hour |
// | A mirror far behind, refresh after refresh reporting `behind` | 12 requests per 5 min, so 144/hour |
//
// Neither number is comfortable beside the scheduler's own runs, which is why the retry
// interval is five minutes rather than one: a refresh that could not finish is unlikely
// to finish thirty seconds later, and spending the hourly budget discovering that would
// break the interactive tools for the rest of the hour — the exact failure the sync's own
// budget exists to prevent.
//
// **Nothing here is depended on across requests.** The interval state is in memory, so a
// fresh Cloud Run instance simply refreshes on its first read. That is a slower first
// call, never a wrong answer.

import { logEvent } from './logging.ts';
import { MirrorLeaseHeldError } from './mirror-store.ts';
import type { SyncReport } from './mirror-sync.ts';

/**
 * How many Graph requests one inline refresh may spend.
 *
 * Deliberately far below the scheduler's 120: this runs inside a tool call, and a
 * refresh that needs more than a dozen requests is a backfill rather than a catch-up.
 * Exceeding it is not a failure — the run keeps the watermarks it earned and reports
 * `behind`, and the next scheduled run continues from there.
 */
export const INLINE_SYNC_REQUEST_BUDGET = 12;

/**
 * Wall clock one inline refresh may spend.
 *
 * Chosen against a person waiting on a tool call rather than against Cloud Run's 300s
 * request ceiling. It does not bound a single hung Graph call: `withRequestTimeout` in
 * ./graph-throttle.ts caps that at 60 seconds, and the sync checks its deadline before
 * starting an operation rather than during one.
 */
export const INLINE_SYNC_TIME_BUDGET_MS = 15_000;

/** How long a finished refresh keeps the mirror labelled current without re-running. */
export const INLINE_SYNC_MIN_INTERVAL_MS = 30_000;

/**
 * How long a refresh that did not finish suppresses the next attempt.
 *
 * Longer than the interval above, not shorter. A refresh reports `behind` because the
 * budget ran out, the lease was held, or Graph refused it — none of which thirty seconds
 * makes better, and all of which would otherwise have every read in a conversation pay
 * the full request budget to learn the same thing again.
 */
export const INLINE_SYNC_RETRY_INTERVAL_MS = 300_000;

/**
 * Whether the mirror can be claimed to equal OneNote.
 *
 * `current` is what turns a mirror-answered tool result into `source: "onenote"`.
 * `behind` is every other outcome, and there are many: the budget ran out, the tree read
 * failed, a page could not be fetched, the scheduler held the lease, Firestore was
 * unreachable. They are one value because the tools do the same thing with all of them.
 */
export type ReadFreshness = 'current' | 'behind';

/** What the read tools call. ./tools.ts binds the real incremental sync to this. */
export interface ReadSync {
  /** Bring the mirror up to date if it is time to. Never throws. */
  refresh(): Promise<ReadFreshness>;
}

/**
 * Does this report license the claim "the mirror equals OneNote"?
 *
 * Four things have to hold, and each of them is a way the run could have left part of
 * the mirror behind while still returning:
 *
 * - `outcome` complete, so nothing failed outright.
 * - `done`, so the request or time budget did not stop the run with sections outstanding.
 * - `treeRead`, because a failed expanded-tree read is *not* fatal to a sync — it carries
 *   on against the structure already in Firestore — and a notebook or section created
 *   since the last successful tree read is absent from that structure.
 * - `pagesFailed` zero. It counts four causes across five increment sites in
 *   ./mirror-sync.ts, and each leaves part of the mirror where it was or mis-describes it:
 *   a page whose content fetch failed keeps whatever copy was stored; a section whose page
 *   enumeration failed is reconciled not at all, which is counted from the incremental
 *   pass and from the sweep separately; a page whose Firestore write threw kept its old
 *   document, and that now includes the short-circuit's `putPageMetadata`, so a transient
 *   Firestore error on a page whose content was confirmed *identical* labels every read
 *   `source: "mirror"` for `INLINE_SYNC_RETRY_INTERVAL_MS`; and a page whose ink render
 *   failed, which is the one cause where the page itself is fine — its text is stored and
 *   only the handwriting is missing, and the label degrades anyway.
 *
 *   The last two are pessimistic on purpose. Neither is worth a separate freshness value:
 *   the tools do the same thing with every `behind`, and a run reporting one of them has
 *   demonstrated that something in the write path is failing.
 *
 * Exported so the rule is asserted directly rather than through a fake sync.
 */
export function freshnessOf(report: SyncReport): ReadFreshness {
  return report.outcome === 'complete' &&
    report.done &&
    report.treeRead &&
    report.pagesFailed === 0
    ? 'current'
    : 'behind';
}

/**
 * Wrap one incremental sync in the interval rule.
 *
 * Two callers arriving together share one run rather than starting two: the second would
 * be refused by the sync lease anyway, and sharing gives it the first one's answer
 * instead of a `behind` it did not earn.
 *
 * `now` is injected so a test does not wait in real time, the way `SyncDeps.now` is.
 */
export function createReadSync(
  run: () => Promise<SyncReport>,
  now: () => number = Date.now,
): ReadSync {
  let inFlight: Promise<ReadFreshness> | null = null;
  let holdUntil = 0;
  let held: ReadFreshness = 'behind';

  const attempt = async (): Promise<ReadFreshness> => {
    let freshness: ReadFreshness;
    try {
      freshness = freshnessOf(await run());
    } catch (err) {
      // A refresh that failed never fails the read. Graph is still reachable through the
      // tool's own fallback, and the mirror still holds whatever it held before — the
      // only consequence is that the answer is labelled `mirror` rather than `onenote`.
      logEvent('read-sync-skipped', { reason: reasonOf(err) });
      freshness = 'behind';
    }

    held = freshness;
    holdUntil =
      now() + (freshness === 'current' ? INLINE_SYNC_MIN_INTERVAL_MS : INLINE_SYNC_RETRY_INTERVAL_MS);
    return freshness;
  };

  return {
    refresh: () => {
      if (now() < holdUntil) return Promise.resolve(held);
      if (inFlight !== null) return inFlight;

      const started = attempt().finally(() => {
        inFlight = null;
      });
      inFlight = started;
      return started;
    },
  };
}

/** Why a refresh did not run to completion. A fixed set; these reach a log line. */
function reasonOf(err: unknown): string {
  if (err instanceof MirrorLeaseHeldError) return 'lease-held';
  return 'failed';
}
