// The one place a Firestore client is constructed.
//
// Until issue #30 there was exactly one consumer — the MSAL token cache — and the client
// lived inside `createFirestoreTokenCachePlugin`. The page mirror is the second, and two
// `Firestore` objects against one database would open two gRPC channels and run two
// credential refreshers for no benefit. So the client is built here and memoised.
//
// Memoised rather than threaded through every constructor, for the same reason
// PRODUCTION_GATE in ./graph-throttle.ts is a module-level constant: it is a
// process-wide shared resource, and passing it down four call sites would put the wiring
// in every signature between here and the leaf. The rule this repository already states
// for GraphAuth is the same rule — "One GraphAuth per process… two MSAL clients would
// each hold their own view of the Firestore cache".
//
// Construction opens no connection. The client connects lazily on the first read or
// write, which is what lets test/tools.test.ts build the real thing with no credential
// and no backend. If that ever stops being true, that test is where it shows up first.

import { Firestore } from '@google-cloud/firestore';

import type { FirestoreConfig } from './config.ts';

/**
 * Keyed by project id so a test — or a future second database — cannot silently get a
 * client pointed at a different project than it asked for. In the deployed service there
 * is only ever one entry.
 */
const clients = new Map<string, Firestore>();

/** The process's Firestore client for `config`, built on first use. */
export function firestoreFor(config: FirestoreConfig): Firestore {
  const key = config.projectId ?? '';

  const existing = clients.get(key);
  if (existing !== undefined) return existing;

  // projectId is left off entirely rather than passed as undefined, so the client uses
  // its own inference path — the Cloud Run metadata server.
  const client = new Firestore(
    config.projectId === undefined ? {} : { projectId: config.projectId },
  );
  clients.set(key, client);
  return client;
}

/**
 * Drop the memoised clients. Tests only.
 *
 * Nothing in `src/` calls this. It exists so a test that asserts something about
 * construction is not deciding the outcome of the next one.
 */
export function resetFirestoreClients(): void {
  clients.clear();
}
