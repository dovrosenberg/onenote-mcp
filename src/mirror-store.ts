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
  NEW_GROUP_DEFAULTS,
  NEW_SECTION_DEFAULTS,
  NOTEBOOKS_COLLECTION,
  PAGES_COLLECTION,
  PAGE_CONTENT_COLLECTION,
  SECTIONS_COLLECTION,
  SECTION_GROUPS_COLLECTION,
  SYNC_STATE_PATH,
  TOMBSTONES_COLLECTION,
  encodeMirrorId,
  leaseIsHeld,
  planStructureWrite,
  readSelection,
  readSyncState,
  type MirrorNotebook,
  type MirrorPage,
  type MirrorPageContent,
  type MirrorPageDigest,
  type MirrorSection,
  type MirrorSectionGroup,
  type MirrorSyncState,
  type MirrorTombstone,
  type NotebookSelection,
  type NotebookStructureWrite,
  type SectionGroupStructureWrite,
  type SectionStructureWrite,
  type StructureIdentity,
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

  /**
   * Every mirrored section, oldest watermark first, which is the sync's visit order.
   *
   * The `orderBy` excludes any document that does not carry `pagesSyncedThrough`, so a
   * section created without it is invisible to the sync permanently and nothing reports a
   * section that was never enumerated. `NEW_SECTION_DEFAULTS` is what guarantees the field
   * on a create; see its docstring.
   */
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
   * Bring the structure collections in line with the tree, and delete what it no longer
   * holds.
   *
   * Called only when the tree hash moved. Deletion is by absence from the incoming set
   * rather than by a `lastSeenAt` sweep, because the caller has just read the complete
   * tree in one request — so absence here is a fact rather than an inference.
   *
   * The creation defaults are per collection and are applied only to a document that was
   * not already there. A notebook has none: every field on it comes from the tree.
   */
  async putStructure(structure: {
    notebooks: readonly NotebookStructureWrite[];
    sectionGroups: readonly SectionGroupStructureWrite[];
    sections: readonly SectionStructureWrite[];
  }): Promise<void> {
    await this.#reconcileCollection(this.#notebooks(), structure.notebooks, {});
    await this.#reconcileCollection(
      this.#sectionGroups(),
      structure.sectionGroups,
      NEW_GROUP_DEFAULTS,
    );
    await this.#reconcileCollection(this.#sections(), structure.sections, NEW_SECTION_DEFAULTS);
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

  /** Every section whose page listing is currently held. Normally none. */
  async listHeldSections(): Promise<MirrorSection[]> {
    // A single-field range filter, so Firestore's automatic index serves it and
    // scripts/gcp-bootstrap.sh needs no composite index for it.
    return this.#query(
      'listing sections with writes in flight',
      this.#sections().where('pendingWrites', '>', 0),
    );
  }

  /**
   * Raise this section's page-listing hold, before a write Graph has not seen yet.
   *
   * `update`, not `set({merge:true})`, for the reason `markPageStale` gives and with a
   * worse failure if it were ignored: a merging set would *create* the section document,
   * and the write tools reach the whole account while only the selection is mirrored — so
   * a create in an unmirrored section would leave a section stub carrying nothing but a
   * counter, which `listAllSections` then feeds to `expandedTree` as a section with no
   * name. NOT_FOUND means there is no listing to hold back.
   */
  async holdSectionListing(sectionId: string, sinceIso: string): Promise<void> {
    await this.#updateSection('holding a section page listing', sectionId, {
      pendingWrites: FieldValue.increment(1),
      pendingWritesSince: sinceIso,
    });
  }

  /** Lower it again, after the write's resync has landed. */
  async releaseSectionListing(sectionId: string): Promise<void> {
    await this.#updateSection('releasing a section page listing', sectionId, {
      pendingWrites: FieldValue.increment(-1),
    });
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

  /** Every stored page in one section, projected to what a sweep reconciles on. */
  async listPageDigestsInSection(sectionId: string): Promise<MirrorPageDigest[]> {
    return this.#query<MirrorPageDigest>(
      'listing mirrored page digests',
      this.#pages()
        .where('sectionId', '==', sectionId)
        .select('id', 'title', 'lastModifiedDateTime', 'contentState'),
    );
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

  async #updateSection(
    operation: string,
    sectionId: string,
    fields: Record<string, unknown>,
  ): Promise<void> {
    await this.#run(operation, async () => {
      try {
        await this.#sections().doc(encodeMirrorId(sectionId)).update(fields);
      } catch (err) {
        if (isNotFound(err)) return;
        throw err;
      }
    });
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
   * Bring one collection in line with `documents`, deleting anything else in it.
   *
   * Every decision about what that means — skip, merge, create, delete — is
   * `planStructureWrite` in ./mirror-schema.ts, which is pure and unit-tested. What is
   * left here is the query, the batching and the calls. Do not put a branch back: this
   * module has no test, so a condition written here is a condition nothing checks.
   *
   * The query reads `identity` alone and supplies both halves of the plan's input: the
   * id set that drives deletion, and the string that drives the skip.
   *
   * Two things here are assumptions about Firestore rather than facts this repository can
   * check, because there is no backend on this machine. Both are settled by watching the
   * first post-deploy sync: it should report a structure pass, leave `backfillComplete`
   * where it was, and leave `pagesSyncedThrough` set on every section it did not create.
   *
   * `select('identity')` is assumed to be a **projection** — every document in the
   * collection comes back, carrying that one field or, for a document written before the
   * field existed, carrying nothing. Firestore documents it that way. If it filtered
   * instead, every such document would be absent from the result, take the create branch,
   * and be written with `pagesSyncedThrough: null` — the account-wide re-backfill this
   * whole design removes. It is self-limiting if wrong: one re-backfill, after which every
   * document carries `identity`.
   *
   * `{ merge: true }` on the `set` below is the line the whole design rests on, and it is
   * the one nothing can catch. `planStructureWrite` deliberately omits the sync-owned
   * fields from `write.fields`, so a `set` without the merge option would replace each
   * document with the tree fields alone and clear `pagesSyncedThrough` on every section
   * whose identity moved. That is narrower than the bug this design removes — it needs an
   * identity change rather than any hash change — and exactly as silent. The store fake in
   * test/mirror-sync.test.ts models the merge rather than asserting it, so deleting this
   * option leaves the suite green.
   *
   * A skipped document's `lastSeenAt` is not refreshed. Nothing reads it — deletion is by
   * absence from the incoming set — and refreshing it would cost the write the skip exists
   * to avoid.
   */
  async #reconcileCollection(
    collection: CollectionReference,
    documents: readonly (StructureIdentity & { id: string })[],
    createDefaults: object,
  ): Promise<void> {
    await this.#run(`reconciling ${collection.id}`, async () => {
      const stored = new Map(
        (await collection.select('identity').get()).docs.map((doc) => [
          doc.id,
          doc.get('identity') as unknown,
        ]),
      );

      const plan = planStructureWrite(stored, documents, createDefaults);

      let batch = this.#firestoreOf(this.#root).batch();
      let queued = 0;
      const flush = async (): Promise<void> => {
        if (queued === 0) return;
        await batch.commit();
        batch = this.#firestoreOf(this.#root).batch();
        queued = 0;
      };

      for (const write of plan.writes) {
        batch.set(
          collection.doc(write.documentId),
          { ...write.fields, lastSeenAt: FieldValue.serverTimestamp() },
          { merge: true },
        );
        queued += 1;
        if (queued >= BATCH_LIMIT) await flush();
      }

      for (const documentId of plan.deletes) {
        batch.delete(collection.doc(documentId));
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
