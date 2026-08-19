// MSAL's ICachePlugin backed by one Firestore document.
//
// Cloud Run containers are ephemeral, so the MSAL cache cannot live on the container
// filesystem: every refresh returns a new refresh token, and losing it means re-running
// the bootstrap CLI by hand. The cache is therefore one document, rewritten in place.
//
// The document holds the serialized blob, a server timestamp, and the blob the last
// write replaced. MSAL owns the blob's structure and changes it between library
// versions, so nothing here parses it — with one deliberate exception, `isEmptyCache`,
// which is written to be ignorant of every key name. See the note there.
//
// Two failure modes decide the shape of this file, and both are about not sending a
// human back to `npm run bootstrap` for something that did not need one:
//
// - **A write that empties the document is refused.** The blob is the only copy of the
//   refresh token. MSAL removes credentials from its in-memory cache on some failures,
//   and `afterCacheAccess` runs in a `finally` block, so a serialization that lost the
//   account can reach this code while the stored one is still good. The guard is in
//   `overwriteWouldEmptyCache`, and the replaced blob is kept in `previousCache` so even
//   a write this guard does let through can be undone by hand.
// - **A backend failure is not a credential failure.** Firestore being unreachable, or a
//   revoked `roles/datastore.user` binding, used to surface through
//   `acquireTokenSilent` as the same error a dead refresh token produces, and the
//   message for that one says to re-run the bootstrap CLI. Those failures are now
//   `TokenCacheUnavailableError`, which ./graph-auth.ts reports as retryable.

import { setTimeout as delay } from 'node:timers/promises';

import { FieldValue, Firestore, type DocumentReference } from '@google-cloud/firestore';
import type { ICachePlugin, TokenCacheContext } from '@azure/msal-node';

import type { FirestoreConfig } from './config.ts';
import { logEvent } from './logging.ts';

const CACHE_FIELD = 'cache';
const UPDATED_AT_FIELD = 'updatedAt';

/**
 * The blob the most recent write replaced.
 *
 * It exists for the case the guard below cannot catch: a write that is not empty but is
 * still wrong. Recovery is a manual copy of this field over `cache`, which is worth
 * having because the alternative is a device-code sign-in. One generation is kept rather
 * than a history — the cache is rewritten on every refresh, so a history would grow
 * without bound and the useful copy is always the most recent good one.
 */
const PREVIOUS_CACHE_FIELD = 'previousCache';

/** How many times a write is attempted before it is reported as unavailable. */
const WRITE_ATTEMPTS = 3;

/** Backoff before the second and third attempt. Short: a Graph request is waiting. */
const WRITE_BACKOFF_MS = [250, 1_000];

export class TokenCacheError extends Error {
  readonly documentPath: string;

  constructor(message: string, documentPath: string) {
    super(message);
    this.name = 'TokenCacheError';
    this.documentPath = documentPath;
  }
}

/**
 * The Firestore backend did not answer, or the write did not commit.
 *
 * Separate from `TokenCacheError` because the two call for opposite responses: a
 * malformed document needs the bootstrap CLI, and this needs a retry. ./graph-auth.ts
 * looks for this type in the cause chain and says so in the message the operator reads.
 */
export class TokenCacheUnavailableError extends Error {
  readonly documentPath: string;
  readonly operation: 'read' | 'write';

  constructor(
    operation: 'read' | 'write',
    documentPath: string,
    options: { cause?: unknown } = {},
  ) {
    super(
      `The Microsoft token cache at ${documentPath} could not be ${operation === 'read' ? 'read' : 'written'}: the Firestore backend did not answer. The stored credential is unaffected; this is a retry, not a re-authorization.`,
    );
    this.name = 'TokenCacheUnavailableError';
    this.documentPath = documentPath;
    this.operation = operation;
    if ('cause' in options) this.cause = options.cause;
  }
}

/**
 * The slice of a Firestore DocumentSnapshot this module reads. Declared narrowly so the
 * decoder is callable from a test with a plain object; `DocumentSnapshot` satisfies it
 * structurally, so the call sites need no cast.
 */
export interface CacheSnapshot {
  readonly exists: boolean;
  data(): Record<string, unknown> | undefined;
}

/**
 * Decode a snapshot of the cache document.
 *
 * Returns null when the document is absent — that is the pre-bootstrap state, not an
 * error. Returns '' when the document exists with an empty `cache` field, which is a
 * real but empty cache; the caller distinguishes the two.
 *
 * @throws {TokenCacheError} if the document exists but `cache` is missing or not a
 * string. Treating that as an empty cache would make MSAL report no accounts, and the
 * next write would overwrite whatever was actually stored there.
 */
export function readCache(snapshot: CacheSnapshot, documentPath: string): string | null {
  if (!snapshot.exists) return null;

  const data = snapshot.data();
  if (data === undefined) return null;

  const value = data[CACHE_FIELD];
  if (typeof value !== 'string') {
    throw new TokenCacheError(
      `Token cache document ${documentPath} has no usable "cache" field (found ${typeof value}). Re-run the bootstrap CLI to recreate it.`,
      documentPath,
    );
  }

  return value;
}

/**
 * Does this serialized cache hold nothing?
 *
 * This is the one place that looks inside the blob, and it is written so that it cannot
 * be wrong about a schema it does not know. It reads no key name: a cache is empty when
 * it parses to a JSON object whose every value is an empty container. MSAL's format is
 * `{"Account": {...}, "RefreshToken": {...}, …}`, so an emptied cache is
 * `{"Account": {}, "RefreshToken": {}, …}` whatever those keys are called this version.
 *
 * Everything unrecognised answers false, which is the safe direction: the only thing
 * this decides is whether to refuse a write, so failing open costs the guard and failing
 * closed would block every write and strand the refresh token in memory.
 */
export function isEmptyCache(serialized: string): boolean {
  if (serialized.trim() === '') return true;

  let parsed: unknown;
  try {
    parsed = JSON.parse(serialized);
  } catch {
    return false;
  }
  if (typeof parsed !== 'object' || parsed === null || Array.isArray(parsed)) return false;

  return Object.values(parsed).every(isEmptyContainer);
}

function isEmptyContainer(value: unknown): boolean {
  if (value === null || value === undefined) return true;
  if (Array.isArray(value)) return value.length === 0;
  if (typeof value === 'object') return Object.keys(value).length === 0;
  return false;
}

/**
 * Would writing `candidate` over `stored` throw away the only copy of the credential?
 *
 * True only when something real is being replaced by nothing. An empty document being
 * filled is the bootstrap case and an empty cache over an already-empty one changes
 * nothing, so neither is refused.
 */
export function overwriteWouldEmptyCache(stored: string | null, candidate: string): boolean {
  if (stored === null || stored === '') return false;
  return !isEmptyCache(stored) && isEmptyCache(candidate);
}

export class FirestoreTokenCachePlugin implements ICachePlugin {
  readonly #firestore: Firestore;
  readonly #documentPath: string;
  readonly #ref: DocumentReference;
  /** The blob last handed to, or written on behalf of, MSAL. null means "document absent". */
  #lastKnown: string | null = null;

  constructor(firestore: Firestore, documentPath: string) {
    this.#firestore = firestore;
    this.#documentPath = documentPath;
    this.#ref = firestore.doc(documentPath);
  }

  /**
   * Load the stored cache into MSAL's in-memory cache.
   *
   * The read is deliberately not in a transaction: it is a plain read, and the
   * compare-and-set that protects against a competing write happens in
   * `afterCacheAccess`.
   */
  async beforeCacheAccess(context: TokenCacheContext): Promise<void> {
    let snapshot: CacheSnapshot;
    try {
      snapshot = await this.#ref.get();
    } catch (err) {
      // A malformed document is TokenCacheError and comes from readCache below. This is
      // the backend itself failing, which is a different message and a different action.
      throw new TokenCacheUnavailableError('read', this.#documentPath, { cause: err });
    }

    const serialized = readCache(snapshot, this.#documentPath);
    this.#lastKnown = serialized;

    if (serialized !== null && serialized !== '') {
      context.tokenCache.deserialize(serialized);
    }
    // An absent or empty document leaves MSAL's in-memory cache untouched, which is
    // already the empty state.
  }

  /**
   * Write the cache back when MSAL says it changed.
   *
   * `--max-instances=1` does not prevent two instances existing during a revision
   * transition, so the read-modify-write runs inside a transaction that re-reads the
   * document and folds in a competing write before serialising.
   *
   * Three things can stop a write committing, and they are not the same event:
   * a competing writer, which the transaction folds in and retries; a candidate blob
   * that would empty the document, which is refused and logged; and the backend not
   * answering, which is retried a few times and then raised as
   * `TokenCacheUnavailableError`.
   */
  async afterCacheAccess(context: TokenCacheContext): Promise<void> {
    if (!context.cacheHasChanged) return;

    let written: string | undefined;
    let refused = false;

    await this.#commit(async () => {
      await this.#firestore.runTransaction(async (tx) => {
        const stored = readCache(await tx.get(this.#ref), this.#documentPath);

        if (stored !== null && stored !== '' && stored !== this.#lastKnown) {
          // Another instance wrote between our read and now. MSAL's deserialize merges
          // into the in-memory cache rather than replacing it, so folding the stored
          // blob back in keeps both sides' entries.
          context.tokenCache.deserialize(stored);
        }

        // serialize() runs on every attempt, so a retry after contention commits a blob
        // computed from the re-read rather than one computed before the conflict.
        const candidate = context.tokenCache.serialize();

        if (overwriteWouldEmptyCache(stored, candidate)) {
          // Nothing is written and nothing is thrown. Throwing would turn whatever
          // failure emptied the in-memory cache into a second, less informative one —
          // `afterCacheAccess` runs in MSAL's `finally`, so the real error is already on
          // its way out. What matters is that the stored refresh token survives.
          refused = true;
          written = undefined;
          return;
        }

        refused = false;
        written = candidate;

        const document: Record<string, unknown> = {
          [CACHE_FIELD]: candidate,
          [UPDATED_AT_FIELD]: FieldValue.serverTimestamp(),
        };
        // Only a real blob that is actually being replaced is worth keeping. Writing the
        // candidate here as well would double the document for nothing.
        if (stored !== null && stored !== '' && stored !== candidate) {
          document[PREVIOUS_CACHE_FIELD] = stored;
        }

        tx.set(this.#ref, document);
      });
    });

    if (refused) {
      // Worth an alert: it means MSAL handed over a cache with no credentials in it, and
      // the next process to start reads the copy this refusal preserved.
      logEvent('token-cache-write-refused', { documentPath: this.#documentPath });
      return;
    }

    // Assigned only after the transaction resolves. Firestore re-runs the callback on
    // contention, so a value assigned inside it could be a blob that never committed.
    this.#lastKnown = written ?? this.#lastKnown;
  }

  /**
   * Run a write, retrying a backend failure a few times before giving up.
   *
   * `runTransaction` already retries its own contention. This is the layer above that:
   * a Firestore outage of a few seconds, or the gRPC channel dropping, costs a wait
   * rather than a refresh token that never reached the document. A `TokenCacheError` is
   * not retried — a malformed document does not repair itself.
   */
  async #commit(attempt: () => Promise<void>): Promise<void> {
    let last: unknown;

    for (let tries = 0; tries < WRITE_ATTEMPTS; tries += 1) {
      if (tries > 0) await delay(WRITE_BACKOFF_MS[tries - 1] ?? 1_000);

      try {
        await attempt();
        return;
      } catch (err) {
        if (err instanceof TokenCacheError) throw err;
        last = err;
      }
    }

    logEvent('token-cache-write-failed', {
      documentPath: this.#documentPath,
      attempts: WRITE_ATTEMPTS,
    });
    throw new TokenCacheUnavailableError('write', this.#documentPath, { cause: last });
  }
}

export function createFirestoreTokenCachePlugin(
  config: FirestoreConfig,
): FirestoreTokenCachePlugin {
  // projectId is left off entirely rather than passed as undefined, so the client uses
  // its own inference path — the Cloud Run metadata server.
  const firestore = new Firestore(
    config.projectId === undefined ? {} : { projectId: config.projectId },
  );
  return new FirestoreTokenCachePlugin(firestore, config.cacheDocumentPath);
}
