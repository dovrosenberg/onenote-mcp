// The route that keeps the Microsoft refresh token from expiring through disuse.
//
// Microsoft's delegated refresh tokens lapse after roughly 90 days without use, and the
// only recovery is a device-code sign-in at a browser — `npm run bootstrap`, run by a
// person, on a machine this service is not. Nothing else in this repository can prevent
// that, because the token only slides forward when it is actually exchanged, and it is
// only exchanged when a tool call arrives after the held access token has expired. A
// connector nobody uses for three months is a connector that needs a human.
//
// So a scheduler calls this route, it calls `GraphAuth.refresh()`, and that forces the
// exchange whether or not anyone has used the service. See the keepalive section of
// README.md for the Cloud Scheduler job.
//
// Three things about it are decisions rather than details:
//
// - **It is not `/healthz`.** That route is public, reports nothing, and costs nothing.
//   This one spends a token-endpoint round trip and a Firestore write on every call, so
//   it is authenticated and it is mounted separately.
// - **The secret is its own variable, not the Layer-1 client secret and not the token
//   signing key.** A scheduler cannot run the OAuth flow — it has no browser and no way
//   to hold a refresh token — so bearer auth is not available to it, and reusing either
//   OAuth credential here would put a credential that can reach the whole MCP surface
//   into a second place for a route that only needs to say "refresh now".
// - **The route is absent when the secret is unset**, rather than present and refusing.
//   An operator who has not configured it gets a 404 from the scheduler and a failed job,
//   which says what is wrong; a 401 would read as a mistyped secret.

import { Router, type Request, type Response } from 'express';
import { timingSafeEqual } from 'node:crypto';

import { GraphAuthError } from './graph-auth.ts';
import { logEvent } from './logging.ts';

/** Where the route is mounted. The scheduler job's URL is this path on MCP_PUBLIC_URL. */
export const KEEPALIVE_PATH = '/keepalive';

/** The header carrying the shared secret. */
export const KEEPALIVE_HEADER = 'x-keepalive-secret';

/**
 * The slice of `GraphAuth` this route calls. Declared narrowly so the handler is
 * testable with a plain object; `GraphAuth` satisfies it structurally.
 */
export interface KeepaliveTarget {
  refresh(): Promise<string>;
}

/**
 * Build the keepalive route.
 *
 * Answers 200 when the refresh succeeded, 401 when the secret is absent or wrong, and
 * 503 when the refresh failed. The 503 body names the `GraphAuthError` reason and
 * whether a retry can fix it, which is the difference between a scheduler job that
 * should keep trying and one whose failure means a human has to sign in again.
 */
export function keepaliveRouter(secret: string, target: KeepaliveTarget): Router {
  const router = Router();

  router.post('/', async (req: Request, res: Response) => {
    res.setHeader('Cache-Control', 'no-store');

    // Checked before anything else, so an unauthenticated flood costs a string
    // comparison rather than a request to Entra and a write to Firestore.
    if (!secretMatches(secret, req.get(KEEPALIVE_HEADER))) {
      logEvent('keepalive-unauthorized');
      res.status(401).json({ status: 'unauthorized' });
      return;
    }

    try {
      // The return value is discarded on purpose. What this call is for is its side
      // effect: Entra issues a replacement refresh token with a fresh inactivity window
      // and ./token-cache.ts writes it to Firestore.
      await target.refresh();
      res.status(200).json({ status: 'ok' });
    } catch (err) {
      // A GraphAuthError's message is safe to return — it names the Firestore document
      // path and the action to take, and no account identifier. Anything else is
      // reported as a bare label: an arbitrary error's message may carry a request body.
      if (err instanceof GraphAuthError) {
        res
          .status(503)
          .json({ status: 'failed', reason: err.reason, retryable: err.retryable, detail: err.message });
        return;
      }

      logEvent('keepalive-failed', { reason: 'unexpected' });
      res.status(503).json({ status: 'failed', reason: 'unexpected', retryable: true });
    }
  });

  // Anything but POST. The refresh is not a read, and answering a GET would let a link
  // or a crawler spend a token exchange.
  router.all('/', (req: Request, res: Response) => {
    res.setHeader('Allow', 'POST');
    res.status(405).json({ status: 'method-not-allowed', method: req.method });
  });

  return router;
}

/**
 * Constant-time comparison of the presented secret.
 *
 * The length is compared first because `timingSafeEqual` throws on a length mismatch.
 * That leaks the length of the configured secret, which is not what protects it — the
 * 32-character minimum in ./config.ts is.
 */
function secretMatches(expected: string, provided: string | undefined): boolean {
  if (provided === undefined) return false;

  const a = Buffer.from(provided, 'utf8');
  const b = Buffer.from(expected, 'utf8');
  if (a.length !== b.length) return false;

  return timingSafeEqual(a, b);
}
