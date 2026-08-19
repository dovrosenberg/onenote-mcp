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
import { createGraphAuth, type GraphAuth } from './graph-auth.ts';
import { createGraphStructure } from './graph-structure.ts';
import type { ToolDefinition } from './mcp-tools.ts';
import { createGraphPageContent } from './page-content.ts';
import { createPageTools } from './page-tools.ts';
import { createGraphPageWrite } from './page-write.ts';
import { createStructureTools } from './structure-tools.ts';
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
export function createTools(auth: GraphAuth): ToolDefinition[] {
  // One page-content client serves both the reading tool and `append_to_page`, which
  // reads a page's layout before it writes so its content does not land on handwriting.
  const content = createGraphPageContent(auth);

  return [
    ...createStructureTools(createGraphStructure(auth)),
    ...createPageTools(content),
    ...createWriteTools(createGraphPageWrite(auth), content),
  ];
}
