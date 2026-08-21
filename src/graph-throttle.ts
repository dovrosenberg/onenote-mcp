// The one place this service decides how fast it is allowed to talk to Graph.
//
// OneNote's documented limits for a delegated app are 5 concurrent requests, 120 per
// minute, and 400 per hour. They are mandatory to consider and will generate failures 
// if ignored (see CLAUDE429 with code 10007) and then lock the account.  So the gate 
// below caps concurrency and spaces request starts, and it is the thing that 
// makes a walk survive rather than any retry.
//
// Retrying is the second half. 429 and 503 are retried: a 400 or a 404 means the
// request itself is wrong, and repeating it spends quota to fail again — that is the
// second best practice in Microsoft's own throttling guidance. `Retry-After` is honoured
// when Graph sends one, and a doubling backoff is used when it does not.
//
// One 500 is retried, and the rule around it is deliberately narrow. Measured 2026-08-19
// and recorded in api-overview.md: every `$expand` on /me/onenote/notebooks answered 500
// with OData code 19999 for seven minutes across 18 attempts, then recovered with no
// change to the request, while un-expanded calls on the same collection answered 200
// throughout. That is a transient service fault rather than a wrong request, and without
// a retry it takes down search_pages, find_page_by_name, list_pages_by_name,
// create_page_by_name and append_to_page_by_name, all of which go through
// getExpandedTree(). So a 500 is retried only when the body carries code 19999 and only
// on a GET — 19999 is the code the service uses when it will not say what went wrong, so
// it can also mean something permanent, and PATCH /pages/{id}/content is not safe to
// repeat blindly.
//
// The gate is injected rather than global so a test can run without waiting: the
// `createGraph*` factories pass PRODUCTION_GATE, and a bare constructor gets UNGATED.

import { logEvent } from './logging.ts';

/** Anything that runs a Graph request under some policy. */
export interface RequestGate {
  run<T>(operation: () => Promise<T>): Promise<T>;
}

/** What a gate needs to know about a failure to decide whether to retry it. */
export interface RetryableError {
  readonly status?: number;
  /** From the `Retry-After` header, already in milliseconds. */
  readonly retryAfterMs?: number;
  /** The response body, read for the OData code on a 500. */
  readonly body?: string;
  /** The HTTP verb. Absent means GET, which is what `GraphRequestError` defaults to. */
  readonly method?: string;
}

/**
 * The OData code Graph returns when it will not say what went wrong.
 *
 * "Something failed, the API cannot share any more information at the time of the
 * request." It covers both the transient `$expand` fault above and permanent ones — an
 * account-wide `/sections` request with no `$filter` answers 500/19999 every time — which
 * is why the retry is capped rather than open-ended.
 */
export const OPAQUE_SERVER_ERROR_CODE = '19999';

export interface GateOptions {
  /** At most this many requests in flight. The documented limit is 5. */
  readonly maxConcurrent?: number;
  /** At least this long between two request starts. 120/minute is one every 500 ms. */
  readonly minIntervalMs?: number;
  /** How many times one request is retried after a 429 or 503. */
  readonly maxRetries?: number;
  /** Wait after the first throttled attempt when Graph sends no `Retry-After`. */
  readonly baseBackoffMs?: number;
  /** Injected so a test does not wait in real time. */
  readonly sleep?: (ms: number) => Promise<void>;
  readonly now?: () => number;
}

/** 4 rather than 5, so a stray request from elsewhere in the process is not the one over. */
export const MAX_CONCURRENT = 4;

/** 120 requests a minute is one every 500 ms. */
export const MIN_INTERVAL_MS = 500;

export const MAX_RETRIES = 3;
export const BASE_BACKOFF_MS = 2_000;

/**
 * The longest a single retry will wait before giving up instead.
 *
 * Graph decides how long a 429 lasts and says so in `Retry-After`, and OneNote's answer
 * can be minutes — CLAUDE.md records five retries spanning three minutes all refused
 * after one burst. Honouring that verbatim is correct for the account and wrong for the
 * process: Cloud Run cuts a request at 300 seconds, and the mirror sync budgets 240, both
 * checked *before* an operation starts. One request that sleeps for three minutes inside
 * the gate blows through both and the run is killed mid-flight.
 *
 * So a wait longer than this is not shortened — that would hammer a service which has
 * just asked for room — it is declined. The caller sees the 429, the sync leaves that
 * section's watermark where it is, and the next scheduled run picks it up. Waiting is the
 * thing being given up on, not the work.
 */
export const MAX_RETRY_WAIT_MS = 30_000;

/** Runs everything immediately. The default for a directly constructed client. */
export const UNGATED: RequestGate = { run: (operation) => operation() };

/**
 * A gate that caps concurrency, spaces starts, and retries throttled requests.
 *
 * Requests queue in arrival order. A retry re-enters the queue rather than holding its
 * slot, so one throttled request cannot block the others while it waits.
 */
export function createGate(options: GateOptions = {}): RequestGate {
  const maxConcurrent = options.maxConcurrent ?? MAX_CONCURRENT;
  const minIntervalMs = options.minIntervalMs ?? MIN_INTERVAL_MS;
  const maxRetries = options.maxRetries ?? MAX_RETRIES;
  const baseBackoffMs = options.baseBackoffMs ?? BASE_BACKOFF_MS;
  const sleep = options.sleep ?? defaultSleep;
  const now = options.now ?? Date.now;

  let inFlight = 0;
  let lastStart = -Infinity;
  const waiting: (() => void)[] = [];

  const release = (): void => {
    inFlight -= 1;
    waiting.shift()?.();
  };

  const acquire = async (): Promise<void> => {
    if (inFlight >= maxConcurrent) {
      await new Promise<void>((resolve) => waiting.push(resolve));
    }
    inFlight += 1;

    // Spacing is applied after the slot is taken, so the interval bounds the request
    // rate rather than the queue length.
    const wait = lastStart + minIntervalMs - now();
    if (wait > 0) await sleep(wait);
    lastStart = now();
  };

  return {
    async run<T>(operation: () => Promise<T>): Promise<T> {
      for (let attempt = 0; ; attempt += 1) {
        await acquire();

        let result: T;
        try {
          result = await operation();
        } catch (err) {
          // The slot is given up before the backoff: holding it would idle a quarter of
          // the allowed concurrency for the length of the wait.
          release();
          const wait = retryWait(err, attempt, maxRetries, baseBackoffMs);
          if (wait === null) throw err;
          await sleep(wait);
          continue;
        }

        release();
        return result;
      }
    },
  };
}

/** How long to wait before retrying, or null when the error is not retryable. */
export function retryWait(
  err: unknown,
  attempt: number,
  maxRetries: number,
  baseBackoffMs: number,
): number | null {
  if (attempt >= maxRetries) return null;

  const failure = err as RetryableError;
  const status = failure.status;

  if (status === 500) {
    // No Retry-After branch: the service sends none on this, and the doubling backoff is
    // what carried a request through the seven-minute window that motivated the rule.
    return isTransientServerError(failure) ? capped(baseBackoffMs * 2 ** attempt) : null;
  }

  if (status !== 429 && status !== 503) return null;

  const retryAfterMs = failure.retryAfterMs;
  if (typeof retryAfterMs === 'number' && retryAfterMs > 0) return capped(retryAfterMs);

  return capped(baseBackoffMs * 2 ** attempt);
}

/** A wait this long is declined rather than shortened. See MAX_RETRY_WAIT_MS. */
function capped(waitMs: number): number | null {
  return waitMs > MAX_RETRY_WAIT_MS ? null : waitMs;
}

/**
 * True for the one 500 worth repeating: OData code 19999, on a read.
 *
 * A body that is not JSON, or is JSON of another shape, answers false — an unrecognised
 * 500 is left alone rather than guessed at.
 */
function isTransientServerError(failure: RetryableError): boolean {
  const method = failure.method ?? 'GET';
  if (method !== 'GET') return false;

  const body = failure.body;
  if (body === undefined) return false;

  try {
    const parsed = JSON.parse(body) as { error?: { code?: unknown } };
    return parsed.error?.code === OPAQUE_SERVER_ERROR_CODE;
  } catch {
    return false;
  }
}

/** `Retry-After` in seconds, or an HTTP date, converted to milliseconds. */
export function parseRetryAfter(header: string | null, now: number = Date.now()): number | undefined {
  if (header === null || header.trim() === '') return undefined;

  const seconds = Number(header);
  if (Number.isFinite(seconds) && seconds >= 0) return seconds * 1000;

  const date = Date.parse(header);
  if (Number.isNaN(date)) return undefined;
  return Math.max(0, date - now);
}

// ---------------------------------------------------------------------------
// Clock skew
// ---------------------------------------------------------------------------

/**
 * How far this process's clock may sit from Graph's before it is worth an event.
 *
 * An HTTP `Date` header carries whole seconds and the value is read after a round trip,
 * so nothing below a second or so is measurable this way. Five seconds is well clear of
 * that and well below every margin in the sync that depends on the two clocks being
 * close.
 */
export const CLOCK_SKEW_LOG_THRESHOLD_MS = 5_000;

/**
 * At most one skew event this often, however many responses arrive.
 *
 * A threshold alone is not a limit: a clock that is simply wrong is over it on every
 * response, and this service can make hundreds of requests an hour. That is log volume
 * with no extra information in it, and it buries `graph-auth-failure` and
 * `token-cache-write-refused`, the two events an operator alerts on. Ten minutes still
 * reports a skew that appears and a skew that goes away within one scheduled sync
 * interval.
 *
 * The window is measured on this process's own clock, which is the one under suspicion.
 * A clock that is wrong by a constant still measures elapsed time correctly, and one
 * jumping far enough to break the window is a fault this event would report anyway.
 */
export const CLOCK_SKEW_LOG_INTERVAL_MS = 10 * 60_000;

/** When the last `graph-clock-skew` line was written, on the caller's clock. */
let lastSkewLogAtMs = -Infinity;

/**
 * Clear the interval limit so one test's event does not suppress the next test's.
 *
 * Tests only; nothing in `src/` calls it, the same arrangement `setEventSink` has.
 */
export function resetClockSkewThrottle(): void {
  lastSkewLogAtMs = -Infinity;
}

/**
 * Record how far this process's clock is from Graph's, from a header already sent.
 *
 * Every cross-clock comparison in the mirror sync — the watermark overlap in
 * `overlapFrom`, and the settle guard in `storedPageIsCurrent` — is sized on an
 * assumption about this number. `api-overview.md` bounds it loosely: one probe on
 * 2026-08-21 put it inside a bracket that contains a full round trip, taken against a
 * developer workstation rather than the deployed service, which rules out minutes and
 * hours and settles nothing tighter. This reports it continuously from the running
 * service's own clock, and costs no request — `Date` is on every response.
 *
 * Returns the signed skew in milliseconds, positive when this process is ahead of Graph,
 * or null when the header is absent or unreadable. Resolution is whole seconds, because
 * that is all an HTTP date carries; a value under a second is rounding, not skew.
 *
 * The argument is a `Headers` and a clock, and deliberately not the response or the URL:
 * all three call sites have a URL in scope, and a URL carries a page id. What is logged
 * is one number.
 */
export function recordClockSkew(headers: Headers, now: () => number = Date.now): number | null {
  const header = headers.get('date');
  if (header === null) return null;

  const serverMs = Date.parse(header);
  if (Number.isNaN(serverMs)) return null;

  const atMs = now();
  const skewMs = atMs - serverMs;

  if (
    Math.abs(skewMs) >= CLOCK_SKEW_LOG_THRESHOLD_MS &&
    atMs - lastSkewLogAtMs >= CLOCK_SKEW_LOG_INTERVAL_MS
  ) {
    lastSkewLogAtMs = atMs;
    logEvent('graph-clock-skew', { skewMs: Math.round(skewMs) });
  }

  return skewMs;
}

function defaultSleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}
