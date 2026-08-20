// The tool registry: which tool modules the deployed server exposes, and the Graph
// client they all share.
//
// This is its own file rather than a function in ./mcp-tools.ts because the tool modules
// import ./mcp-tools.ts for `ToolDefinition` and the argument helpers. Putting the
// wiring there too would make that pair of modules import each other.
//
// The Graph client is built once per process and shared by every request. Building one
// per request would throw away MSAL's in-memory access token and hit the token endpoint
// every time — see the note on `createGraphAuth` in ./graph-auth.ts.

import type { Config } from './config.ts';
import { firestoreFor } from './firestore.ts';
import { createGraphAuth, type GraphAuth } from './graph-auth.ts';
import { createGraphStructure } from './graph-structure.ts';
import type { ToolDefinition } from './mcp-tools.ts';
import { createMirrorBlobStore } from './mirror-blobs.ts';
import { MirrorReader } from './mirror-reader.ts';
import { createMirrorStore } from './mirror-store.ts';
import type { MirrorWriteSync } from './write-tools.ts';
import {
  resyncPage,
  runFullSweep,
  runIncremental,
  runSweep,
  runSweepAll,
  type SyncContent,
  type SyncDeps,
} from './mirror-sync.ts';
import { createGraphPageContent } from './page-content.ts';
import {
  INLINE_SYNC_REQUEST_BUDGET,
  INLINE_SYNC_TIME_BUDGET_MS,
  createReadSync,
  type ReadSync,
} from './read-sync.ts';
import { createPageTools } from './page-tools.ts';
import { createGraphPageWrite } from './page-write.ts';
import { createStructureTools } from './structure-tools.ts';
import type { SyncTarget } from './sync-route.ts';
import { createWriteTools } from './write-tools.ts';

/**
 * Build the process's one Graph auth object.
 *
 * Separate from `createTools` because the keepalive route in ./keepalive.ts needs the
 * same object. Two of them would mean two MSAL clients, each with its own in-memory
 * access token and its own view of the Firestore cache — a forced refresh through one
 * would leave the other holding a superseded blob until its next read, and both would be
 * writing the same document.
 */
export function createGraphAuthFor(config: Config): GraphAuth {
  const { graph, firestore } = config;
  if (graph === undefined || firestore === undefined) {
    throw new Error("internal: createGraphAuthFor needs the 'graph' and 'firestore' config groups");
  }

  return createGraphAuth(graph, firestore);
}

/**
 * Every tool this server exposes.
 *
 * Issue #15 adds the browsing tools, #16 the reading tool, and #18 the three writing
 * tools.
 *
 * They are built over three Graph clients because the endpoints answer with different
 * things — JSON for structure, a `multipart/mixed` body for page content, 204 with no
 * body for a write — but all three take the same auth object, so one MSAL client and one
 * token cache serve the whole process. All three also share the process-wide request
 * gate, which is what keeps reads and writes together inside OneNote's per-user limits.
 *
 * The writing tools come last so `tools/list` reads as browse, read, then write.
 */
/**
 * The mirror reader, when reads are configured to use it.
 *
 * Undefined turns the whole feature off: every tool module treats an absent mirror as
 * "always Graph", so `MIRROR_READ_ENABLED=false` is a complete rollback with no code
 * change and no data migration. The bucket is required by `loadConfig`'s cross-field
 * rule, so the check here is a type narrowing rather than a reachable branch.
 */
export function createMirrorReaderFor(config: Config): MirrorReader | undefined {
  const mirror = config.mirror;
  const firestore = config.firestore;
  if (mirror === undefined || firestore === undefined || !mirror.readEnabled) return undefined;

  if (mirror.bucket === undefined) {
    throw new Error('internal: MIRROR_READ_ENABLED is true without MIRROR_BUCKET');
  }

  return new MirrorReader(
    createMirrorStore(firestoreFor(firestore), mirror.rootDocumentPath),
    createMirrorBlobStore(mirror.bucket, firestore.projectId),
  );
}

/**
 * What the write tools call after a successful write.
 *
 * Bound whenever a mirror exists at all, not only when reads are enabled: a mirror being
 * filled by the sync while `MIRROR_READ_ENABLED` is false still holds copies a write
 * supersedes, and keeping them current then is what makes turning reads on later safe
 * rather than a race with whatever was written in between.
 *
 * It shares the Graph content client with the read tools, so a resync runs through the
 * same process-wide request gate every other Graph call does — a write burst cannot
 * outrun the per-user rate limit by going through this path.
 */
export function createMirrorWriteSyncFor(
  config: Config,
  auth: GraphAuth,
  content: SyncContent,
): MirrorWriteSync | undefined {
  const mirror = config.mirror;
  const firestore = config.firestore;
  if (mirror === undefined || firestore === undefined) return undefined;
  if (mirror.syncSecret === undefined && !mirror.readEnabled) return undefined;
  if (mirror.bucket === undefined) return undefined;

  const store = createMirrorStore(firestoreFor(firestore), mirror.rootDocumentPath);
  const blobs = createMirrorBlobStore(mirror.bucket, firestore.projectId);
  const deps = { store, blobs, content };

  return {
    resyncPage: (pageId, hint) => resyncPage(deps, pageId, hint),
    markPageStale: (pageId) => store.markPageStale(pageId),
    sectionOfPage: async (pageId) => (await store.getPage(pageId))?.sectionId ?? null,
    // The hold's timestamp is stamped here rather than with a server timestamp, for the
    // reason the sync lease stamps `runningSince` here: the value is compared against
    // `Date.now()` in the read path, and mixing a Firestore `Timestamp` into that
    // comparison would mean decoding one on every page listing.
    holdSectionListing: (sectionId) =>
      store.holdSectionListing(sectionId, new Date().toISOString()),
    releaseSectionListing: (sectionId) => store.releaseSectionListing(sectionId),
  };
}

export function createTools(
  auth: GraphAuth,
  mirror?: MirrorReader,
  writeSync?: MirrorWriteSync,
  readSync?: ReadSync,
): ToolDefinition[] {
  // One page-content client serves both the reading tool and `append_to_page`, which
  // reads a page's layout before it writes so its content does not land on handwriting.
  const content = createGraphPageContent(auth);

  // One structure client serves the browsing tools and `create_page_by_name`, which
  // resolves a notebook/section-group/section path through the same resolver the
  // `_by_name` browsing tools use.
  const structure = createGraphStructure(auth);

  return [
    // The mirror and its inline sync are passed to the read tools only. Writes always go
    // to Graph, which stays the source of truth; the mirror is a copy that the sync
    // refreshes, and the read tools refresh it themselves before every read.
    ...createStructureTools(structure, {}, mirror, readSync),
    ...createPageTools(content, undefined, mirror, readSync),
    // The mirror reaches the write tools only as a write-sync: a successful write
    // re-reads its page from Graph and stores it, so the next read answers with what was
    // just written rather than falling through until the next scheduled run.
    ...createWriteTools(createGraphPageWrite(auth), content, structure, writeSync),
  ];
}

/**
 * The sync's dependencies, or undefined when the mirror is not configured at all.
 *
 * Shared by the scheduled route and by the inline refresh the read tools run, so the two
 * cannot end up talking to different Firestore roots or different buckets. Each caller
 * gets its own object; the Firestore client underneath is memoised by project id in
 * ./firestore.ts, and all three Graph clients pass through the same process-wide request
 * gate, so an inline refresh cannot outrun the per-user rate limit.
 */
function syncDepsFor(config: Config, auth: GraphAuth): SyncDeps | undefined {
  const mirror = config.mirror;
  const firestore = config.firestore;
  if (mirror === undefined || firestore === undefined || mirror.bucket === undefined) {
    return undefined;
  }

  return {
    graph: createGraphStructure(auth),
    content: createGraphPageContent(auth),
    store: createMirrorStore(firestoreFor(firestore), mirror.rootDocumentPath),
    blobs: createMirrorBlobStore(mirror.bucket, firestore.projectId),
  };
}

/**
 * The refresh every read tool attempts before it reads the mirror.
 *
 * Built only when reads are enabled, because it exists to make the `source` a tool
 * reports an honest claim and a tool with no mirror to read reports `onenote` anyway.
 * With `MIRROR_READ_ENABLED=false` this is undefined, every read goes straight to Graph,
 * and no tool call spends a Graph request on a sync — which is what keeps the flag a
 * complete rollback.
 *
 * The budgets are the inline ones in ./read-sync.ts, not `MIRROR_SYNC_REQUEST_BUDGET`:
 * that variable sizes a scheduled job with 300 seconds to spend, and this runs inside a
 * tool call a person is waiting on.
 */
export function createReadSyncFor(config: Config, auth: GraphAuth): ReadSync | undefined {
  if (config.mirror?.readEnabled !== true) return undefined;

  const deps = syncDepsFor(config, auth);
  if (deps === undefined) {
    throw new Error('internal: MIRROR_READ_ENABLED is true without MIRROR_BUCKET');
  }

  return createReadSync(() =>
    runIncremental(deps, {
      requestBudget: INLINE_SYNC_REQUEST_BUDGET,
      timeBudgetMs: INLINE_SYNC_TIME_BUDGET_MS,
    }),
  );
}

/**
 * Bind the page mirror's sync to the route, or answer undefined.
 *
 * Undefined whenever the mirror is not configured to sync, which is what keeps
 * `POST /sync` unmounted and the path a 404. A 404 tells an operator the service is
 * unconfigured; a 401 would read as a mistyped secret.
 *
 * The bucket is required by `loadConfig`'s cross-field rule whenever the sync secret is
 * set, so the check here cannot fail in practice — it is a type narrowing, and a throw
 * rather than a silent skip because a sync with nowhere to put a rendered PNG should not
 * start.
 */
export function createSyncTargetFor(config: Config, auth: GraphAuth): SyncTarget | undefined {
  const mirror = config.mirror;
  const firestore = config.firestore;
  if (mirror?.syncSecret === undefined || firestore === undefined) return undefined;

  const deps = syncDepsFor(config, auth);
  if (deps === undefined) {
    throw new Error("internal: MIRROR_SYNC_SECRET is set without MIRROR_BUCKET");
  }

  const options = { requestBudget: mirror.syncRequestBudget };

  return {
    runIncremental: () => runIncremental(deps, options),
    runSweep: () => runSweep(deps, options),
    runFullSweep: () => runFullSweep(deps, options),
    runSweepAll: () => runSweepAll(deps, options),
  };
}
