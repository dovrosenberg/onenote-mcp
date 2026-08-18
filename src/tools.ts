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
import { createGraphAuth } from './graph-auth.ts';
import { createGraphStructure } from './graph-structure.ts';
import type { ToolDefinition } from './mcp-tools.ts';
import { createStructureTools } from './structure-tools.ts';

/**
 * Every tool this server exposes.
 *
 * Issue #15 adds the four browsing tools. Reading (#16) and writing (#18) append to the
 * list returned here.
 */
export function createTools(config: Config): ToolDefinition[] {
  const { graph, firestore } = config;
  if (graph === undefined || firestore === undefined) {
    throw new Error("internal: createTools needs the 'graph' and 'firestore' config groups");
  }

  const structure = createGraphStructure(createGraphAuth(graph, firestore));

  return [...createStructureTools(structure)];
}
