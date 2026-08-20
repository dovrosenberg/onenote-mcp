// Firestore access for the page mirror. Calls, and as little judgement as possible.
//
// Everything in this file that could be *decided* wrongly lives in ./mirror-schema.ts
// instead: document ids, field shapes, the oversize threshold, the watermark arithmetic,
// how a hand-edited document is read. That is not tidiness. This file has no automated
// test and cannot get one on this machine — there is no Firestore emulator here, and
// CLAUDE.md forbids substituting an in-memory fake because the behaviour at stake is
// transaction retry under contention and FieldValue.serverTimestamp(), and a fake would
// assert the fake. So the split is what keeps the untestable surface down to method
// calls whose failure is obvious in a live run.
//
// Four things in here are not obvious:
//
// **The lease is a transaction, and it is the only thing preventing a double spend.**
// Two scheduler jobs, or a retry landing on a run that is still going, would otherwise
// both walk the tree and both fetch content — twice the Graph budget for one result.
//
// **Structure writes are batched and skipped wholesale when the tree hash is unchanged.**
// The account has 55 notebooks and 568 sections, so rewriting every structure document
// every fifteen minutes would be ~2.5 million writes a month for a tree that changes
// when someone adds a notebook.
//
// **Nothing here filters on a `deleted` field, because there is no such field.** A
// deleted page is removed and a tombstone written to its own collection. See the note in
// ./mirror-schema.ts for why.
//
// **Every query in here has a matching composite index in scripts/gcp-bootstrap.sh.** A
// missing one fails at query time with FAILED_PRECONDITION and a console link, which
// nothing in CI would catch.

import {
  FieldValue,
  type CollectionReference,
  type DocumentReference,
  type Firestore,
  type Query,
} from '@google-cloud/firestore';

import { logEvent } from './logging.ts';
import {
  NOTEBOOKS_COLLECTION,
  PAGES_COLLECTION,
  PAGE_CONTENT_COLLECTION,
  SECTIONS_COLLECTION,
  SECTION_GROUPS_COLLECTION,
  SYNC_STATE_PATH,
  TOMBSTONES_COLLECTION,
  encodeMirrorId,
  leaseIsHeld,
  readSelection,
  readSyncState,
  type MirrorNotebook,
  type MirrorPage,
  type MirrorPageContent,
  type MirrorSection,
  type MirrorSectionGroup,
  type MirrorSyncState,
  type MirrorTombstone,
  type NotebookSelection,
  type SyncMode,
} from './mirror-schema.ts';

/**
 * A Firestore read or write did not answer.
 *
 * Its own type for the reason `TokenCacheUnavailableError` is: on the read path the
 * mirror is an optimisation and this must fall through to Graph, and on the sync path it
 * must end the slice with the cursor intact. A bare rejection would be indistinguishable
 * from a decode failure, which is not retryable.
 */
export class MirrorUnavailableError extends Error {
  readonly operation: string;

  constructor(operation: string, options: { cause?: unknown } = {}) {
    super(
      `The page mirror could not be ${operation}: the Firestore backend did not answer. Reads fall back to Microsoft Graph; nothing stored is affected.`,
    );
    this.name = 'MirrorUnavailableError';
    this.operation = operation;
    if ('cause' in options) this.cause = options.cause;
  }
}

/** Raised when the lease is held by a run that has not expired. */
export class MirrorLeaseHeldError extends Error {
  readonly heldBy: SyncMode;

  constructor(heldBy: SyncMode) {
    super(`A ${heldBy} sync is already running.`);
    this.name = 'MirrorLeaseHeldError';
    this.heldBy = heldBy;
  }
}

/** How many documents go in one batched write. Firestore's own limit is 500. */
const BATCH_LIMIT = 450;

/** How many pages an unscoped title scan reads before it reports itself truncated. */
export const SCAN_LIMIT = 5_000;

/** The projection an unscoped title scan reads. Never the HTML. */
const SCAN_FIELDS = [
  'id',
  'title',
  'sectionId',
  'notebookId',
  'sectionPath',
  'lastModifiedDateTime',
] as const;

export interface ScanResult {
  readonly pages: MirrorPage[];
  /** True when SCAN_LIMIT was reached, so the answer is a sample. */
  readonly truncated: boolean;
}

export class MirrorStore {
  readonly #root: DocumentReference;

  constructor(firestore: Firestore, rootDocumentPath: string) {
    this.#root = firestore.doc(rootDocumentPath);
  }

  // -------------------------------------------------------------------------
  // Selection and sync state
  // -------------------------------------------------------------------------

  /** The hand-edited notebook list. This method's document is never written here. */
  async getSelection(): Promise<NotebookSelection> {
    return readSelection(await this.#read('reading the notebook selection', this.#root));
  }

  async getSyncState(): Promise<MirrorSyncState> {
    return readSyncState(await this.#read('reading the sync state', this.#stateRef()));
  }

  /** Merge fields into the sync state. Absent fields are left alone. */
  async patchSyncState(patch: Partial<MirrorSyncState>): Promise<void> {
    await this.#run('writing the sync state', () =>
      this.#stateRef().set({ ...patch, updatedAt: FieldValue.serverTimestamp() }, { merge: true }),
    );
  }

  /**
   * Take the lease, or refuse.
   *
   * A transaction rather than a read-then-write, because the thing being prevented is
   * exactly the interleaving a read-then-write allows: two runs both seeing no lease and
   * both proceeding. `--max-instances=1` does not help — a revision transition runs two
   * instances, and the nightly sweep and a retried incremental can overlap on one.
   */
  async acquireLease(mode: SyncMode, nowIso: string): Promise<void> {
    await this.#run('taking the sync lease', () =>
      this.#firestoreOf(this.#stateRef()).runTransaction(async (tx) => {
        const snapshot = await tx.get(this.#stateRef());
        const state = readSyncState(snapshot.data());

        if (leaseIsHeld(state, Date.parse(nowIso)) && state.runningMode !== null) {
          throw new MirrorLeaseHeldError(state.runningMode);
        }

        tx.set(
          this.#stateRef(),
          { runningMode: mode, runningSince: nowIso },
          { merge: true },
        );
      }),
    );
  }

  /**
   * Release the lease, but only if this run still holds it.
   *
   * The conditional part is not defensive. Cloud Run cuts a request at 300 seconds and
   * throttles CPU outside a request, so an overrunning run is *frozen* rather than
   * killed: it resumes whenever the next request arrives, possibly many minutes later.
   * By then the lease has expired on age and another run has taken it. An unconditional
   * clear would then release the live run's lease and let a third start alongside it,
   * both spending the same hourly Graph budget — which is the exact failure the lease
   * exists to prevent, reached by way of the lease itself.
   *
   * `runningSince` is the discriminator: it is the ISO timestamp this run wrote when it
   * took the lease, so a run that no longer sees its own value knows it has been
   * superseded and leaves the document alone.
   *
   * Never allowed to fail the run. The slice's work is already committed by the time this
   * is called, and an exception here would turn a successful sync into a reported failure
   * that the scheduler then retries. A lease that fails to clear expires on age instead,
   * which costs one skipped run at worst.
   */
  async releaseLease(heldSince: string): Promise<void> {
    try {
      await this.#firestoreOf(this.#stateRef()).runTransaction(async (tx) => {
        const state = readSyncState((await tx.get(this.#stateRef())).data());
        if (state.runningSince !== heldSince) {
          // Superseded. Someone else's lease; not ours to clear.
          logEvent('mirror-lease-superseded');
          return;
        }
        tx.set(this.#stateRef(), { runningMode: null, runningSince: null }, { merge: true });
      });
    } catch (err) {
      logEvent('mirror-lease-release-failed', { reason: reasonOf(err) });
    }
  }

  // -------------------------------------------------------------------------
  // Structure
  // -------------------------------------------------------------------------

  async listNotebooks(): Promise<MirrorNotebook[]> {
    return this.#query('listing mirrored notebooks', this.#notebooks().orderBy('displayName'));
  }

  async listSectionsUnder(parentId: string): Promise<MirrorSection[]> {
    return this.#query(
      'listing mirrored sections',
      this.#sections().where('parentId', '==', parentId).orderBy('displayName'),
    );
  }

  async listSectionGroupsUnder(parentId: string): Promise<MirrorSectionGroup[]> {
    return this.#query(
      'listing mirrored section groups',
      this.#sectionGroups().where('parentId', '==', parentId).orderBy('displayName'),
    );
  }

  /** Every mirrored section, oldest watermark first, which is the sync's visit order. */
  async listSectionsToSync(): Promise<MirrorSection[]> {
    return this.#query(
      'listing sections to sync',
      this.#sections().where('mirrored', '==', true).orderBy('pagesSyncedThrough', 'asc'),
    );
  }

  /** Every section and section group, for the read path's expanded-tree equivalent. */
  async listAllSections(): Promise<MirrorSection[]> {
    return this.#query('listing every mirrored section', this.#sections());
  }

  async listAllSectionGroups(): Promise<MirrorSectionGroup[]> {
    return this.#query('listing every mirrored section group', this.#sectionGroups());
  }

  async getSection(sectionId: string): Promise<MirrorSection | null> {
    return this.#one('reading a mirrored section', this.#sections().doc(encodeMirrorId(sectionId)));
  }

  async getSectionGroup(groupId: string): Promise<MirrorSectionGroup | null> {
    return this.#one(
      'reading a mirrored section group',
      this.#sectionGroups().doc(encodeMirrorId(groupId)),
    );
  }

  async getNotebook(notebookId: string): Promise<MirrorNotebook | null> {
    return this.#one(
      'reading a mirrored notebook',
      this.#notebooks().doc(encodeMirrorId(notebookId)),
    );
  }

  /**
   * Replace the whole structure, in batches, and delete what the tree no longer holds.
   *
   * Called only when the tree hash moved. Deletion is by absence from the incoming set
   * rather than by a `lastSeenAt` sweep, because the caller has just read the complete
   * tree in one request — so absence here is a fact rather than an inference.
   */
  async putStructure(structure: {
    notebooks: readonly MirrorNotebook[];
    sectionGroups: readonly MirrorSectionGroup[];
    sections: readonly MirrorSection[];
  }): Promise<void> {
    await this.#replaceCollection(this.#notebooks(), structure.notebooks);
    await this.#replaceCollection(this.#sectionGroups(), structure.sectionGroups);
    await this.#replaceCollection(this.#sections(), structure.sections);
  }

  /** Advance one section's watermark. Written only after its whole changed set landed. */
  async setSectionWatermark(sectionId: string, watermarkIso: string): Promise<void> {
    await this.#run('advancing a section watermark', () =>
      this.#sections()
        .doc(encodeMirrorId(sectionId))
        .set({ pagesSyncedThrough: watermarkIso }, { merge: true }),
    );
  }

  async setSectionSweepResult(
    sectionId: string,
    fields: { pageCount: number; lastSweptAt: string },
  ): Promise<void> {
    await this.#run('recording a section sweep', () =>
      this.#sections().doc(encodeMirrorId(sectionId)).set(fields, { merge: true }),
    );
  }

  async setChildGroupsKnown(groupId: string, known: boolean): Promise<void> {
    await this.#run('recording nested section groups', () =>
      this.#sectionGroups()
        .doc(encodeMirrorId(groupId))
        .set({ childGroupsKnown: known }, { merge: true }),
    );
  }

  // -------------------------------------------------------------------------
  // Pages
  // -------------------------------------------------------------------------

  async getPage(pageId: string): Promise<MirrorPage | null> {
    return this.#one('reading a mirrored page', this.#pages().doc(encodeMirrorId(pageId)));
  }

  async getPageContent(pageId: string): Promise<MirrorPageContent | null> {
    return this.#one(
      'reading mirrored page content',
      this.#pageContent().doc(encodeMirrorId(pageId)),
    );
  }

  /** Pages in one section, newest first. `limit` is a result count. */
  async listPagesInSection(sectionId: string, limit?: number): Promise<MirrorPage[]> {
    let query: Query = this.#pages()
      .where('sectionId', '==', sectionId)
      .orderBy('lastModified', 'desc');
    if (limit !== undefined) query = query.limit(limit);

    return this.#query('listing mirrored pages', query);
  }

  /** Every mirrored page id in one section, for the sweep's set difference. */
  async listPageIdsInSection(sectionId: string): Promise<string[]> {
    const pages = await this.#query<{ id: string }>(
      'listing mirrored page ids',
      this.#pages().where('sectionId', '==', sectionId).select('id'),
    );
    return pages.map((page) => page.id);
  }

  /**
   * Page metadata across the mirror, projected, for a title substring search.
   *
   * Firestore cannot match a substring, so the filtering happens in the caller. What
   * makes that affordable is `.select()`: the projection is applied server-side, so a
   * scan of 2000 pages transfers a few hundred kilobytes of titles rather than the
   * gigabytes of HTML those documents would carry if the two collections were one.
   *
   * `limit` is SCAN_LIMIT plus one, so the result can say whether it was truncated
   * rather than silently answering from a sample.
   */
  async scanPages(scope: { notebookId?: string; sectionId?: string } = {}): Promise<ScanResult> {
    let query: Query = this.#pages();
    if (scope.sectionId !== undefined) query = query.where('sectionId', '==', scope.sectionId);
    else if (scope.notebookId !== undefined) {
      query = query.where('notebookId', '==', scope.notebookId);
    }

    const pages = await this.#query<MirrorPage>(
      'scanning mirrored page titles',
      query
        .orderBy('lastModified', 'desc')
        .select(...SCAN_FIELDS)
        .limit(SCAN_LIMIT + 1),
    );

    return { pages: pages.slice(0, SCAN_LIMIT), truncated: pages.length > SCAN_LIMIT };
  }

  /** Write a page's metadata and, when it is inline, its content. */
  async putPage(page: MirrorPage, content: MirrorPageContent | null): Promise<void> {
    await this.#run('writing a mirrored page', async () => {
      const batch = this.#firestoreOf(this.#root).batch();
      batch.set(this.#pages().doc(encodeMirrorId(page.id)), {
        ...page,
        // Parallel to lastModifiedDateTime, which stays the ISO string Graph sent because
        // every tool result prints it. This one exists so Firestore can order and
        // range-filter, which it cannot do on a string reliably across time zones.
        lastModified: new Date(page.lastModifiedDateTime),
        contentSyncedAt: FieldValue.serverTimestamp(),
      });

      const contentRef = this.#pageContent().doc(encodeMirrorId(page.id));
      if (content === null) batch.delete(contentRef);
      else batch.set(contentRef, { ...content, fetchedAt: FieldValue.serverTimestamp() });

      await batch.commit();
    });
  }

  /** Update only the metadata Graph's page list carries, leaving content untouched. */
  async putPageMetadata(
    page: Pick<
      MirrorPage,
      'id' | 'title' | 'titleLower' | 'sectionId' | 'notebookId' | 'sectionPath' | 'lastModifiedDateTime'
    >,
  ): Promise<void> {
    await this.#run('writing mirrored page metadata', () =>
      this.#pages()
        .doc(encodeMirrorId(page.id))
        .set({ ...page, lastModified: new Date(page.lastModifiedDateTime) }, { merge: true }),
    );
  }

  /**
   * Mark a page's content stale, so the next read goes to Graph.
   *
   * What a write tool calls. The content document is deleted rather than kept and
   * flagged, because a read that fell through on the flag and a read that fell through
   * on a missing document take the same path, and keeping superseded HTML around only
   * creates a chance of serving it.
   */
  async markPageStale(pageId: string): Promise<void> {
    await this.#run('marking a mirrored page stale', async () => {
      const batch = this.#firestoreOf(this.#root).batch();
      // `update`, not `set({merge:true})`. A merging set *creates* the document when it
      // is absent, and the write tools reach the whole account while only the selection
      // is mirrored — so every write to an unmirrored page would leave a stub carrying
      // nothing but `contentState: 'stale'`. Harmless to read, but it fills the queried
      // collection with rows for pages the mirror does not hold.
      batch.update(this.#pages().doc(encodeMirrorId(pageId)), { contentState: 'stale' });
      batch.delete(this.#pageContent().doc(encodeMirrorId(pageId)));

      try {
        await batch.commit();
      } catch (err) {
        // NOT_FOUND means there was no copy to invalidate, which is the ordinary answer
        // for a page outside the mirrored set. Anything else is a real failure.
        if (isNotFound(err)) return;
        throw err;
      }
    });
  }

  /** Remove a page and its content, and record why it went. */
  async deletePage(tombstone: MirrorTombstone): Promise<void> {
    await this.#run('deleting a mirrored page', async () => {
      const batch = this.#firestoreOf(this.#root).batch();
      const encoded = encodeMirrorId(tombstone.id);
      batch.delete(this.#pages().doc(encoded));
      batch.delete(this.#pageContent().doc(encoded));
      batch.set(this.#tombstones().doc(encoded), {
        ...tombstone,
        deletedAt: FieldValue.serverTimestamp(),
      });
      await batch.commit();
    });
  }

  // -------------------------------------------------------------------------
  // Internals
  // -------------------------------------------------------------------------

  #stateRef(): DocumentReference {
    return this.#root.collection(SYNC_STATE_PATH[0]).doc(SYNC_STATE_PATH[1]);
  }

  #notebooks(): CollectionReference {
    return this.#root.collection(NOTEBOOKS_COLLECTION);
  }

  #sectionGroups(): CollectionReference {
    return this.#root.collection(SECTION_GROUPS_COLLECTION);
  }

  #sections(): CollectionReference {
    return this.#root.collection(SECTIONS_COLLECTION);
  }

  #pages(): CollectionReference {
    return this.#root.collection(PAGES_COLLECTION);
  }

  #pageContent(): CollectionReference {
    return this.#root.collection(PAGE_CONTENT_COLLECTION);
  }

  #tombstones(): CollectionReference {
    return this.#root.collection(TOMBSTONES_COLLECTION);
  }

  #firestoreOf(ref: DocumentReference): Firestore {
    return ref.firestore;
  }

  async #read(
    operation: string,
    ref: DocumentReference,
  ): Promise<Record<string, unknown> | undefined> {
    return this.#run(operation, async () => (await ref.get()).data());
  }

  async #one<T>(operation: string, ref: DocumentReference): Promise<T | null> {
    return this.#run(operation, async () => {
      const snapshot = await ref.get();
      return snapshot.exists ? (snapshot.data() as T) : null;
    });
  }

  async #query<T>(operation: string, query: Query): Promise<T[]> {
    return this.#run(operation, async () =>
      (await query.get()).docs.map((doc) => doc.data() as T),
    );
  }

  /**
   * Replace a whole collection with `documents`, deleting anything else in it.
   *
   * Reads the existing ids first so the delete set is exact. That read is one query
   * against a collection of at most a few hundred documents, and it runs only when the
   * tree hash moved — which is when someone added or renamed a notebook, not every
   * fifteen minutes.
   */
  async #replaceCollection(
    collection: CollectionReference,
    documents: readonly { id: string }[],
  ): Promise<void> {
    await this.#run(`replacing ${collection.id}`, async () => {
      const existing = new Set((await collection.select().get()).docs.map((doc) => doc.id));

      let batch = this.#firestoreOf(this.#root).batch();
      let queued = 0;
      const flush = async (): Promise<void> => {
        if (queued === 0) return;
        await batch.commit();
        batch = this.#firestoreOf(this.#root).batch();
        queued = 0;
      };

      for (const document of documents) {
        const encoded = encodeMirrorId(document.id);
        existing.delete(encoded);
        batch.set(collection.doc(encoded), {
          ...document,
          lastSeenAt: FieldValue.serverTimestamp(),
        });
        queued += 1;
        if (queued >= BATCH_LIMIT) await flush();
      }

      for (const staleId of existing) {
        batch.delete(collection.doc(staleId));
        queued += 1;
        if (queued >= BATCH_LIMIT) await flush();
      }

      await flush();
    });
  }

  /**
   * Run a Firestore call, turning a backend failure into `MirrorUnavailableError`.
   *
   * `MirrorLeaseHeldError` passes through: it is this module's own answer, not a backend
   * failure, and the route reports it as a 409 rather than a 503.
   */
  async #run<T>(operation: string, call: () => Promise<T>): Promise<T> {
    try {
      return await call();
    } catch (err) {
      if (err instanceof MirrorLeaseHeldError) throw err;
      throw new MirrorUnavailableError(operation, { cause: err });
    }
  }
}

/** Firestore's gRPC status for a document that is not there. */
function isNotFound(err: unknown): boolean {
  return (err as { code?: unknown }).code === 5;
}

/** A reason string for a log line. Never a message, which can carry a request body. */
function reasonOf(err: unknown): string {
  return err instanceof Error ? err.name : 'unknown';
}

export function createMirrorStore(firestore: Firestore, rootDocumentPath: string): MirrorStore {
  return new MirrorStore(firestore, rootDocumentPath);
}
