// The request gate: concurrency, spacing, and which failures are retried.
//
// Time is injected rather than waited on — `sleep` records what it was asked for and
// advances a clock — so the test asserts the policy rather than the wall clock. What no
// test here proves is that the numbers are the right ones; 5 concurrent and 120 per
// minute come from Microsoft's documented limits, and the one measurement behind them is
// in the `Graph request budget` section of CLAUDE.md.

import test from 'node:test';
import assert from 'node:assert/strict';

import {
  MAX_CONCURRENT,
  MIN_INTERVAL_MS,
  UNGATED,
  createGate,
  parseRetryAfter,
  retryWait,
} from '../src/graph-throttle.ts';

/** A clock that only moves when something sleeps. */
function fakeClock(): {
  now: () => number;
  sleep: (ms: number) => Promise<void>;
  slept: number[];
} {
  let clock = 1_000;
  const slept: number[] = [];
  return {
    now: () => clock,
    sleep: (ms: number) => {
      slept.push(ms);
      clock += ms;
      return Promise.resolve();
    },
    slept,
  };
}

function deferred(): { promise: Promise<void>; resolve: () => void } {
  let resolve = (): void => {};
  const promise = new Promise<void>((res) => {
    resolve = res;
  });
  return { promise, resolve };
}

test('the gate never runs more than maxConcurrent at once', async () => {
  const clock = fakeClock();
  const gate = createGate({ maxConcurrent: 2, minIntervalMs: 0, ...clock });

  let inFlight = 0;
  let peak = 0;
  const blockers = Array.from({ length: 6 }, () => deferred());

  const runs = blockers.map((blocker) =>
    gate.run(async () => {
      inFlight += 1;
      peak = Math.max(peak, inFlight);
      await blocker.promise;
      inFlight -= 1;
    }),
  );

  // Let the first wave start, then release everything.
  await Promise.resolve();
  assert.equal(peak, 2, 'a third request must wait for a slot');
  for (const blocker of blockers) blocker.resolve();
  await Promise.all(runs);

  assert.equal(peak, 2);
});

test('request starts are spaced by minIntervalMs', async () => {
  const clock = fakeClock();
  const gate = createGate({ maxConcurrent: 1, minIntervalMs: 500, ...clock });

  await gate.run(() => Promise.resolve('a'));
  await gate.run(() => Promise.resolve('b'));
  await gate.run(() => Promise.resolve('c'));

  // The first goes immediately; each later one waits out the interval.
  assert.deepEqual(clock.slept, [500, 500]);
});

test('a 429 is retried after the Retry-After the response carried', async () => {
  const clock = fakeClock();
  const gate = createGate({ maxConcurrent: 2, minIntervalMs: 0, ...clock });

  let attempts = 0;
  const result = await gate.run(() => {
    attempts += 1;
    if (attempts === 1) {
      return Promise.reject(Object.assign(new Error('throttled'), { status: 429, retryAfterMs: 7_000 }));
    }
    return Promise.resolve('ok');
  });

  assert.equal(result, 'ok');
  assert.equal(attempts, 2);
  assert.deepEqual(clock.slept, [7_000], "Graph's own wait is used, not the local backoff");
});

test('a 429 with no Retry-After backs off, doubling each attempt', async () => {
  const clock = fakeClock();
  const gate = createGate({
    maxConcurrent: 2,
    minIntervalMs: 0,
    baseBackoffMs: 1_000,
    maxRetries: 3,
    ...clock,
  });

  let attempts = 0;
  await assert.rejects(() =>
    gate.run(() => {
      attempts += 1;
      return Promise.reject(Object.assign(new Error('throttled'), { status: 429 }));
    }),
  );

  assert.equal(attempts, 4, 'the first attempt plus maxRetries');
  assert.deepEqual(clock.slept, [1_000, 2_000, 4_000]);
});

test('a 400 is not retried, because repeating it spends quota to fail again', async () => {
  const clock = fakeClock();
  const gate = createGate({ maxConcurrent: 2, minIntervalMs: 0, ...clock });

  let attempts = 0;
  await assert.rejects(() =>
    gate.run(() => {
      attempts += 1;
      return Promise.reject(Object.assign(new Error('bad request'), { status: 400 }));
    }),
  );

  assert.equal(attempts, 1);
  assert.deepEqual(clock.slept, []);
});

test('a throttled request gives up its slot while it waits', async () => {
  const clock = fakeClock();
  const gate = createGate({ maxConcurrent: 1, minIntervalMs: 0, ...clock });

  let started = 0;
  let attempts = 0;

  const throttled = gate.run(() => {
    attempts += 1;
    started += 1;
    if (attempts === 1) {
      return Promise.reject(Object.assign(new Error('throttled'), { status: 503 }));
    }
    return Promise.resolve('retried');
  });
  const other = gate.run(() => {
    started += 1;
    return Promise.resolve('other');
  });

  assert.deepEqual(await Promise.all([throttled, other]), ['retried', 'other']);
  // Three starts from two calls: the retry is a third pass through the gate, and the
  // second call was not blocked behind the backoff.
  assert.equal(started, 3);
});

test('retryWait covers what the gate decides', () => {
  assert.equal(retryWait({ status: 429 }, 0, 3, 1_000), 1_000);
  assert.equal(retryWait({ status: 503 }, 2, 3, 1_000), 4_000);
  assert.equal(retryWait({ status: 429, retryAfterMs: 250 }, 0, 3, 1_000), 250);
  assert.equal(retryWait({ status: 429 }, 3, 3, 1_000), null, 'attempts are capped');
  assert.equal(retryWait({ status: 404 }, 0, 3, 1_000), null);
  assert.equal(retryWait(new Error('network'), 0, 3, 1_000), null);
});

test('Retry-After is read as seconds or as an HTTP date', () => {
  assert.equal(parseRetryAfter('30'), 30_000);
  assert.equal(parseRetryAfter('0'), 0);
  assert.equal(parseRetryAfter(null), undefined);
  assert.equal(parseRetryAfter(''), undefined);
  assert.equal(parseRetryAfter('not a number'), undefined);

  const now = Date.parse('2026-03-04T10:00:00Z');
  assert.equal(parseRetryAfter('Wed, 04 Mar 2026 10:00:20 GMT', now), 20_000);
  // A date already past means wait no time at all, not a negative wait.
  assert.equal(parseRetryAfter('Wed, 04 Mar 2026 09:59:00 GMT', now), 0);
});

test('UNGATED runs immediately and the production defaults match the documented limits', async () => {
  assert.equal(await UNGATED.run(() => Promise.resolve(42)), 42);
  // 5 concurrent and 120 a minute are the documented limits; these stay under both.
  assert.ok(MAX_CONCURRENT < 5);
  assert.ok(MIN_INTERVAL_MS >= 500);
});
