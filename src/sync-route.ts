// POST /sync — the scheduler's way in, authenticated by a shared secret.
//
// Structurally ./keepalive.ts: the secret is compared in constant time before any work
// happens, the route is not mounted at all when no secret is configured, only POST is
// answered, and it sits outside the bearer gate because a scheduler cannot run the OAuth
// flow — it has no browser and nowhere to keep a refresh token. It is its own variable
// rather than the keepalive secret because the two do different things and a credential
// should reach one of them, not both.
//
// **Four paths rather than one path with a mode.** `src/logging.ts` records the method,
// the path and the status, and deliberately records no query string and no body. A mode
// carried in either would appear nowhere in the request log, and "which job ran, and did
// it answer?" is the first question when the mirror looks wrong. A body would also need
// a JSON parser on a route outside the bearer gate. And a time-based rule would make
// behaviour depend on the container's clock and remove the ability to force a sweep on
// demand, which is exactly the move the keepalive runbook documents as the way to prove
// a job works.
//
// **A budget-exhausted run answers 200, not 503.** It is a normal outcome with a report
// attached, and the work it did is committed. A 503 would make the scheduler retry
// immediately and spend the next hour's Graph budget inside this one, which is the
// failure the budget exists to prevent.

import { timingSafeEqual } from 'node:crypto';

import { Router, type Request, type Response } from 'express';

import { GraphAuthError } from './graph-auth.ts';
import { logEvent } from './logging.ts';
import { MirrorLeaseHeldError } from './mirror-store.ts';
import type { SyncReport } from './mirror-sync.ts';

export const SYNC_PATH = '/sync';
export const SYNC_HEADER = 'x-sync-secret';

/** What the route drives. ./tools.ts binds the real sync to this. */
export interface SyncTarget {
  runIncremental(): Promise<SyncReport>;
  runSweep(): Promise<SyncReport>;
  runFullSweep(): Promise<SyncReport>;
  runSweepAll(): Promise<SyncReport>;
}

export function syncRouter(secret: string, target: SyncTarget): Router {
  const router = Router();

  const run = (mode: keyof SyncTarget) => async (req: Request, res: Response): Promise<void> => {
    // Set first, so it is present on every answer including the 401.
    res.setHeader('Cache-Control', 'no-store');

    if (!secretMatches(secret, req.get(SYNC_HEADER))) {
      logEvent('sync-unauthorized');
      res.status(401).json({ status: 'unauthorized' });
      return;
    }

    try {
      const report = await target[mode]();
      res.status(200).json({ status: 'ok', ...report });
    } catch (err) {
      if (err instanceof MirrorLeaseHeldError) {
        // Not a failure. The nightly sweep landing on a still-running incremental is an
        // ordinary overlap, and the scheduler retrying later is the right response.
        logEvent('sync-conflict', { mode });
        res.status(409).json({ status: 'conflict', heldBy: err.heldBy, retryable: true });
        return;
      }

      if (err instanceof GraphAuthError) {
        // The one failure a human has to act on: the refresh token is gone and no retry
        // will fix it. Its own reason reaches the operator rather than being flattened.
        res.status(503).json({
          status: 'failed',
          reason: err.reason,
          retryable: err.retryable,
          detail: err.message,
        });
        return;
      }

      logEvent('sync-failed', { mode, reason: 'unexpected' });
      res.status(503).json({ status: 'failed', reason: 'unexpected', retryable: true });
    }
  };

  router.post('/', run('runIncremental'));
  router.post('/sweep', run('runSweep'));
  router.post('/sweep/full', run('runFullSweep'));
  router.post('/sweep/all', run('runSweepAll'));

  // Only POST. A GET would let a link preview, or anything that crawls a URL, spend a
  // slice of the hourly Graph budget.
  for (const path of ['/', '/sweep', '/sweep/full', '/sweep/all']) {
    router.all(path, (_req: Request, res: Response) => {
      res.setHeader('Allow', 'POST');
      res.status(405).json({ status: 'method-not-allowed' });
    });
  }

  return router;
}

/**
 * Constant-time comparison.
 *
 * `timingSafeEqual` throws on a length mismatch rather than answering false, so the
 * lengths are compared first — which leaks the length and nothing else.
 */
function secretMatches(expected: string, provided: string | undefined): boolean {
  if (provided === undefined) return false;

  const a = Buffer.from(provided, 'utf8');
  const b = Buffer.from(expected, 'utf8');
  if (a.length !== b.length) return false;

  return timingSafeEqual(a, b);
}
