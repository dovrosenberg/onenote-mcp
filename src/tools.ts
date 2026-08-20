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
import type { MirrorInvalidator } from './write-tools.ts';
import { runFullSweep, runIncremental, runSweep, type SyncDeps } from './mirror-sync.ts';
import { createGraphPageContent } from './page-content.ts';
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
 * The invalidator the write tools call after a successful write.
 *
 * Bound whenever a mirror exists at all, not only when reads are enabled: a mirror being
 * filled by the sync while `MIRROR_READ_ENABLED` is false still holds copies that a write
 * supersedes, and marking them stale then is what makes turning reads on later safe
 * rather than a race with whatever was written in between.
 */
export function createMirrorInvalidatorFor(config: Config): MirrorInvalidator | undefined {
  const mirror = config.mirror;
  const firestore = config.firestore;
  if (mirror === undefined || firestore === undefined) return undefined;
  if (mirror.syncSecret === undefined && !mirror.readEnabled) return undefined;

  return createMirrorStore(firestoreFor(firestore), mirror.rootDocumentPath);
}

export function createTools(
  auth: GraphAuth,
  mirror?: MirrorReader,
  invalidator?: MirrorInvalidator,
): ToolDefinition[] {
  // One page-content client serves both the reading tool and `append_to_page`, which
  // reads a page's layout before it writes so its content does not land on handwriting.
  const content = createGraphPageContent(auth);

  // One structure client serves the browsing tools and `create_page_by_name`, which
  // resolves a notebook/section-group/section path through the same resolver the
  // `_by_name` browsing tools use.
  const structure = createGraphStructure(auth);

  return [
    // The mirror is passed to the read tools only. Writes always go to Graph, which
    // stays the source of truth; the mirror is a copy that the sync refreshes.
    ...createStructureTools(structure, {}, mirror),
    ...createPageTools(content, undefined, mirror),
    // The mirror reaches the write tools only as an invalidator: a successful write
    // marks its page stale so the next read goes to Graph rather than serving the copy
    // the write just superseded.
    ...createWriteTools(createGraphPageWrite(auth), content, structure, invalidator),
  ];
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

  if (mirror.bucket === undefined) {
    throw new Error("internal: MIRROR_SYNC_SECRET is set without MIRROR_BUCKET");
  }

  const deps: SyncDeps = {
    graph: createGraphStructure(auth),
    content: createGraphPageContent(auth),
    store: createMirrorStore(firestoreFor(firestore), mirror.rootDocumentPath),
    blobs: createMirrorBlobStore(mirror.bucket, firestore.projectId),
  };
  const options = { requestBudget: mirror.syncRequestBudget };

  return {
    runIncremental: () => runIncremental(deps, options),
    runSweep: () => runSweep(deps, options),
    runFullSweep: () => runFullSweep(deps, options),
  };
}
