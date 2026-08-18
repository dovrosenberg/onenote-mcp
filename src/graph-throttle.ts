// The one place this service decides how fast it is allowed to talk to Graph.
//
// OneNote's documented limits for a delegated app are 5 concurrent requests, 120 per
// minute, and 400 per hour. They are mandatory to consider and will generate failures 
// if ignored (see CLAUDE429 with code 10007) and then lock the account.  So the gate 
// below caps concurrency and spaces request starts, and it is the thing that 
// makes a walk survive rather than any retry.
//
// Retrying is the second half. Only 429 and 503 are retried: a 400 or a 404 means the
// request itself is wrong, and repeating it spends quota to fail again — that is the
// second best practice in Microsoft's own throttling guidance. `Retry-After` is honoured
// when Graph sends one, and a doubling backoff is used when it does not.
//
// The gate is injected rather than global so a test can run without waiting: the
// `createGraph*` factories pass PRODUCTION_GATE, and a bare constructor gets UNGATED.

/** Anything that runs a Graph request under some policy. */
export interface RequestGate {
  run<T>(operation: () => Promise<T>): Promise<T>;
}

/** What a gate needs to know about a failure to decide whether to retry it. */
export interface RetryableError {
  readonly status?: number;
  /** From the `Retry-After` header, already in milliseconds. */
  readonly retryAfterMs?: number;
}

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

  const status = (err as RetryableError).status;
  if (status !== 429 && status !== 503) return null;

  const retryAfterMs = (err as RetryableError).retryAfterMs;
  if (typeof retryAfterMs === 'number' && retryAfterMs > 0) return retryAfterMs;

  return baseBackoffMs * 2 ** attempt;
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

function defaultSleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}
