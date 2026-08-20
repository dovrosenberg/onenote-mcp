// The contract every MCP tool in this server implements, the registry the JSON-RPC layer
// reads, and the one place a thrown error becomes something a model can read.
//
// A tool carries its own JSON Schema and receives arguments that nothing has validated.
// That follows from ./mcp-server.ts building the protocol surface on the SDK's low-level
// `Server` rather than on `McpServer` — see the comment there for why. The helpers at the
// bottom of this file are what turn raw arguments into typed values, and a bad argument
// into a readable failure rather than a TypeError.
//
// The registry of which tool modules the server exposes lives in ./tools.ts, not here:
// those modules import this one for `ToolDefinition` and the argument helpers, and
// wiring them up here as well would make the two files import each other.

import type { CallToolResult, Tool } from '@modelcontextprotocol/sdk/types.js';

import { GraphAuthError } from './graph-auth.ts';
import { GraphRequestError, GraphResponseError } from './graph-structure.ts';
import { InkParseError, InkRenderError } from './ink.ts';
import { MirrorUnavailableError } from './mirror-store.ts';
import { NameLookupError } from './name-lookup.ts';

/**
 * One tool: what `tools/list` advertises, and what `tools/call` invokes.
 *
 * `handle` receives the arguments object verbatim, never null — an omitted `arguments`
 * member becomes `{}` before it gets here, so every tool can read from it directly.
 */
export interface ToolDefinition {
  readonly name: string;
  readonly title: string;
  readonly description: string;
  readonly inputSchema: Tool['inputSchema'];
  readonly annotations?: Tool['annotations'];
  handle(args: Readonly<Record<string, unknown>>): Promise<CallToolResult>;
}

/**
 * An argument the caller got wrong. Separate from every other error because it is the
 * caller's to fix: the message names the argument and what was expected, and the model
 * on the other end can retry the call without a human.
 */
export class ToolInputError extends Error {
  readonly argument: string;

  constructor(argument: string, expected: string) {
    super(`Argument '${argument}' ${expected}.`);
    this.name = 'ToolInputError';
    this.argument = argument;
  }
}

/** The `tools/list` payload. Drops `handle` and keeps the declared shape. */
export function toolDescriptors(tools: readonly ToolDefinition[]): Tool[] {
  return tools.map((tool) => ({
    name: tool.name,
    title: tool.title,
    description: tool.description,
    inputSchema: tool.inputSchema,
    ...(tool.annotations === undefined ? {} : { annotations: tool.annotations }),
  }));
}

/** Index by name, rejecting a duplicate at construction rather than shadowing silently. */
export function indexTools(tools: readonly ToolDefinition[]): Map<string, ToolDefinition> {
  const byName = new Map<string, ToolDefinition>();
  for (const tool of tools) {
    if (byName.has(tool.name)) {
      throw new Error(`internal: two tools registered as '${tool.name}'`);
    }
    byName.set(tool.name, tool);
  }
  return byName;
}

/**
 * Turn a thrown error into a tool result the caller can act on.
 *
 * Every failure below is expected in normal operation — an expired refresh token, a page
 * id that no longer exists, a document resvg will not render — so each becomes an
 * `isError` result rather than a JSON-RPC error. The distinction matters to the client:
 * a JSON-RPC error is a protocol fault, and an `isError` result is an answer the model is
 * meant to read and respond to.
 *
 * Nothing here quotes page content, a notebook or section name, or a token. The Graph
 * body is reduced to its `error.code` and `error.message`, which is where "20266,
 * maximum sections exceeded" lives and is the only part worth showing.
 */
export function toolErrorResult(toolName: string, err: unknown): CallToolResult {
  return { isError: true, content: [{ type: 'text', text: toolErrorText(toolName, err) }] };
}

function toolErrorText(toolName: string, err: unknown): string {
  if (err instanceof ToolInputError) {
    return `${toolName}: ${err.message}`;
  }
  if (err instanceof GraphAuthError) {
    // The message already ends in `npm run bootstrap`, which is the action a human has
    // to take. Passing it through is the whole point of GraphAuthError existing.
    return `${toolName}: ${err.message}`;
  }
  if (err instanceof GraphRequestError) {
    const detail = graphErrorDetail(err.body);
    return (
      `${toolName}: Microsoft Graph returned ${err.status} ${err.statusText}` +
      `${detail === null ? '' : ` (${detail})`}.`
    );
  }
  if (err instanceof GraphResponseError) {
    return `${toolName}: ${err.message}`;
  }
  if (err instanceof InkParseError || err instanceof InkRenderError) {
    return `${toolName}: ${err.message}`;
  }
  if (err instanceof NameLookupError) {
    // The message lists the sibling names that were actually there, which is what lets
    // the calling model fix the argument without another browsing call.
    return `${toolName}: ${err.message}`;
  }
  if (err instanceof MirrorUnavailableError) {
    // Only reachable when the local mirror *and* Microsoft Graph both failed — a mirror
    // failure alone falls through to Graph and never becomes a tool error. Its own arm so
    // the caller learns the local copy is not the problem to fix. No cause message: it
    // wraps an arbitrary Firestore error, which can carry a document path.
    return (
      `${toolName}: ${err.message} The direct read from Microsoft Graph also failed; ` +
      'check the service logs.'
    );
  }
  // Anything else is a bug in this server. Report that it happened and nothing more:
  // an arbitrary error's message may carry a request body, and this output is read by a
  // client that may log it.
  return `${toolName}: the server hit an unexpected error. Check the service logs.`;
}

/** How much of Graph's `error.message` reaches the caller. */
const MAX_GRAPH_DETAIL_CHARS = 300;

/**
 * `{ "error": { "code": ..., "message": ... } }` is the OData error shape Graph returns.
 * A body that is not that shape yields null, so a stray HTML error page from a proxy is
 * dropped rather than pasted into a tool result.
 */
function graphErrorDetail(body: string): string | null {
  let parsed: unknown;
  try {
    parsed = JSON.parse(body);
  } catch {
    return null;
  }
  if (typeof parsed !== 'object' || parsed === null || !('error' in parsed)) return null;

  const error = (parsed as { error: unknown }).error;
  if (typeof error !== 'object' || error === null) return null;

  const code = 'code' in error && typeof error.code === 'string' ? error.code : null;
  const message =
    'message' in error && typeof error.message === 'string'
      ? error.message.slice(0, MAX_GRAPH_DETAIL_CHARS)
      : null;

  if (code !== null && message !== null) return `${code}: ${message}`;
  return code ?? message;
}

// ---------------------------------------------------------------------------
// Argument helpers. Every tool in project-spec.md takes ids and short strings plus one
// optional count, so these three cover the surface; each throws ToolInputError, which
// toolErrorResult turns into a message naming the argument. A tool with a fixed set of
// allowed values checks it itself and throws the same error — see `containerType` in
// ./structure-tools.ts.
// ---------------------------------------------------------------------------

/** A non-empty string. Whitespace-only is treated as missing: an id cannot be blank. */
export function requiredString(args: Readonly<Record<string, unknown>>, name: string): string {
  const value = args[name];
  if (typeof value !== 'string' || value.trim() === '') {
    throw new ToolInputError(name, 'is required and must be a non-empty string');
  }
  return value;
}

/** The same, but an absent or null argument yields undefined rather than throwing. */
export function optionalString(
  args: Readonly<Record<string, unknown>>,
  name: string,
): string | undefined {
  const value = args[name];
  if (value === undefined || value === null) return undefined;
  if (typeof value !== 'string' || value.trim() === '') {
    throw new ToolInputError(name, 'must be a non-empty string when given');
  }
  return value;
}

/**
 * A boolean, or undefined when absent.
 *
 * Strict about the type rather than truthy: a model passing the string `"false"` means
 * false, and coercing it to true would silently do the opposite of what was asked.
 */
export function optionalBoolean(
  args: Readonly<Record<string, unknown>>,
  name: string,
): boolean | undefined {
  const value = args[name];
  if (value === undefined || value === null) return undefined;
  if (typeof value !== 'boolean') {
    throw new ToolInputError(name, 'must be true or false when given');
  }
  return value;
}

/** A whole number inside `[min, max]`, or undefined when absent. */
export function optionalInteger(
  args: Readonly<Record<string, unknown>>,
  name: string,
  range: { min: number; max: number },
): number | undefined {
  const value = args[name];
  if (value === undefined || value === null) return undefined;
  if (typeof value !== 'number' || !Number.isInteger(value)) {
    throw new ToolInputError(name, 'must be a whole number when given');
  }
  if (value < range.min || value > range.max) {
    throw new ToolInputError(name, `must be between ${range.min} and ${range.max}`);
  }
  return value;
}
