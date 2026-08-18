// MSAL's ICachePlugin backed by one Firestore document.
//
// Cloud Run containers are ephemeral, so the MSAL cache cannot live on the container
// filesystem: every refresh returns a new refresh token, and losing it means re-running
// the bootstrap CLI by hand. The cache is therefore one document, rewritten in place.
//
// The document holds exactly two fields — the serialized blob and a server timestamp.
// MSAL owns that blob's structure and changes it between library versions, so nothing
// here parses it.

import { FieldValue, Firestore, type DocumentReference } from '@google-cloud/firestore';
import type { ICachePlugin, TokenCacheContext } from '@azure/msal-node';

import type { FirestoreConfig } from './config.ts';

const CACHE_FIELD = 'cache';
const UPDATED_AT_FIELD = 'updatedAt';

export class TokenCacheError extends Error {
  readonly documentPath: string;

  constructor(message: string, documentPath: string) {
    super(message);
    this.name = 'TokenCacheError';
    this.documentPath = documentPath;
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
    const serialized = readCache(await this.#ref.get(), this.#documentPath);
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
   */
  async afterCacheAccess(context: TokenCacheContext): Promise<void> {
    if (!context.cacheHasChanged) return;

    let written: string | undefined;

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
      written = context.tokenCache.serialize();
      tx.set(this.#ref, {
        [CACHE_FIELD]: written,
        [UPDATED_AT_FIELD]: FieldValue.serverTimestamp(),
      });
    });

    // Assigned only after runTransaction resolves. Firestore re-runs the callback on
    // contention, so a value assigned inside it could be a blob that never committed.
    this.#lastKnown = written ?? this.#lastKnown;
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
