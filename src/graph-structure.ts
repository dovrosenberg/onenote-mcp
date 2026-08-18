// Read-only traversal of the OneNote structure: notebooks, section groups, sections,
// and the page list inside one section. Ported from the recon script in Appendix A of
// project-spec.md and generalised — pagination and nesting are handled here rather than
// assumed away as the script could.
//
// The account-wide page list `/me/onenote/pages` is never called. It fails with error
// 20266, "maximum sections exceeded", once the account has enough sections across all
// notebooks, and a notebook-per-year with a section-group-per-month reaches that quickly.
// Page listing is always scoped to `/sections/{id}/pages`. A test in
// test/graph-structure.test.ts scans every file under src/ for that path.
//
// Nesting is real: a notebook holds sections and section groups, and a section group
// holds further sections and section groups. `getNotebookTree` walks that recursion;
// nothing here assumes one level.

import type { GraphAuth } from './graph-auth.ts';
import {
  UNGATED,
  createGate,
  parseRetryAfter,
  type RequestGate,
} from './graph-throttle.ts';

export const GRAPH_ROOT = 'https://graph.microsoft.com/v1.0';

/** `$select` for structure nodes; display names are needed to show a tree at all. */
const NODE_SELECT = '$select=id,displayName&$orderby=displayName';

/**
 * One request that returns a notebook's sections, its section groups, and those groups'
 * sections. `$select` is repeated inside every expand clause because it is worth 5.7x on
 * the response: 441 KB without it against 78 KB with it, for the same tree. The separator
 * inside a clause carrying both `$select` and `$expand` is a semicolon, not a comma.
 *
 * This is the shape the OneNote throttling guidance asks for — see the `Graph request
 * budget` section of CLAUDE.md. The per-container walk that returns the same data costs
 * `1 + 2 x containers`, which is 195 requests on the real account against this one.
 */
const EXPANDED_TREE_URL =
  `${GRAPH_ROOT}/me/onenote/notebooks?$select=id,displayName` +
  `&$expand=sections($select=id,displayName),` +
  `sectionGroups($select=id,displayName;$expand=sections($select=id,displayName))`;

/**
 * Sections anywhere in the account whose name contains the query, with their parents.
 *
 * This reaches a section at any nesting depth in one request, which neither a walk nor
 * `$expand` does — `parentSectionGroup` is what says where the section actually sits, and
 * `$expand` is capped at two levels (a third nesting level answers 400, and `$levels=max`
 * answers 400 as well).
 *
 * `contains(tolower(displayName), '…')` rather than `tolower(displayName) eq '…'`:
 * the second answers 500 with code 19999 on this endpoint, which is undocumented and was
 * found by testing. `contains` is therefore how the comparison is made case-insensitive,
 * and the exact match is applied to the results by the caller. The account-wide
 * `/me/onenote/sections` with no `$filter` at all also answers 500, so this URL is only
 * usable with the filter present.
 */
const SECTIONS_BY_NAME_URL =
  `${GRAPH_ROOT}/me/onenote/sections?$select=id,displayName` +
  `&$expand=parentNotebook($select=id,displayName),parentSectionGroup($select=id,displayName)` +
  `&$filter=`;

/**
 * One container with both its child relationships inlined.
 *
 * Two requests become one. `listSections` and `listSectionGroups` still exist because a
 * caller sometimes wants only one of the two, but nothing should ask for both separately.
 */
function containerChildrenUrl(kind: ContainerKind, containerId: string): string {
  return (
    `${GRAPH_ROOT}/me/onenote/${kind}/${encodeURIComponent(containerId)}?$select=id,displayName` +
    `&$expand=sections($select=id,displayName),sectionGroups($select=id,displayName)`
  );
}

/** The most result pages one list call will follow before giving up. */
const MAX_PAGE_FOLLOWS = 50;

/**
 * How many containers a walk visits at once.
 *
 * The gate enforces this too, but the walk bounds itself as well: a client constructed
 * without a gate — which is what a test does — must not be able to open 108 requests at
 * once, because that is the shape that got this repository throttled.
 */
const WALK_CONCURRENCY = 4;

/** Graph's own ceiling on `$top`; above it the request is a 400 with code 20129. */
export const MAX_GRAPH_TOP = 100;

/** How deep section groups may nest before the walk is treated as a cycle. */
const MAX_SECTION_GROUP_DEPTH = 20;

/** How much of an error body reaches the message; the whole body stays on `.body`. */
const MAX_BODY_CHARS = 2000;

/** A non-2xx response from Graph. Carries the status and the body, per issue #11. */
export class GraphRequestError extends Error {
  readonly url: string;
  readonly status: number;
  readonly statusText: string;
  readonly body: string;
  /** From `Retry-After`, in milliseconds. Only Graph knows how long a 429 lasts. */
  readonly retryAfterMs: number | undefined;

  constructor(
    url: string,
    status: number,
    statusText: string,
    body: string,
    retryAfterMs?: number,
  ) {
    super(
      `GET ${url} failed: ${status} ${statusText}${body === '' ? '' : ` ${truncate(body, MAX_BODY_CHARS)}`}`,
    );
    this.name = 'GraphRequestError';
    this.url = url;
    this.status = status;
    this.statusText = statusText;
    this.body = body;
    this.retryAfterMs = retryAfterMs;
  }
}

/**
 * A 2xx response whose body is not the shape the caller needs, or a listing that would
 * not terminate. Separate from GraphRequestError because there is no status to act on:
 * a caller can retry a 429 and cannot usefully retry this.
 */
export class GraphResponseError extends Error {
  readonly url: string;

  constructor(message: string, url: string) {
    super(message);
    this.name = 'GraphResponseError';
    this.url = url;
  }
}

/**
 * The slice of `GraphAuth` this module calls. Declared narrowly so the client is
 * testable with a plain object; `GraphAuth` satisfies it structurally.
 */
export interface TokenSource {
  getAccessToken(): Promise<string>;
}

/** The slice of `fetch` this module uses, so a test can supply its own. */
export type FetchLike = (
  url: string,
  init: { headers: Record<string, string> },
) => Promise<Response>;

/** `notebooks` and `sectionGroups` expose the same two child relationships. */
export type ContainerKind = 'notebooks' | 'sectionGroups';

export interface Notebook {
  readonly id: string;
  readonly displayName: string;
}

export interface Section {
  readonly id: string;
  readonly displayName: string;
}

export interface SectionGroup {
  readonly id: string;
  readonly displayName: string;
}

export interface PageSummary {
  readonly id: string;
  readonly title: string;
  readonly lastModifiedDateTime: string;
}

/** What one container directly holds, before any recursion. */
export interface ContainerChildren {
  readonly sections: Section[];
  readonly sectionGroups: SectionGroup[];
}

/** A section found by name, with the containers it sits in. */
export interface SectionWithParents extends Section {
  readonly parentNotebook: Notebook | null;
  readonly parentSectionGroup: SectionGroup | null;
}

/** A section group with everything below it already resolved. */
export interface SectionGroupNode extends SectionGroup {
  readonly sections: Section[];
  readonly sectionGroups: SectionGroupNode[];
}

/** A notebook with the whole section-group tree below it resolved. */
export interface NotebookTree extends Notebook {
  readonly sections: Section[];
  readonly sectionGroups: SectionGroupNode[];
}

/**
 * A section group as the expanded tree returns it: its own sections, but not its nested
 * section groups. The nested `$expand` reaches one level, so anything deeper is absent
 * from the response rather than empty in it — `nestedGroupsUnknown` says which.
 */
export interface ExpandedSectionGroup extends SectionGroup {
  readonly sections: Section[];
}

/** One notebook out of the expanded tree. */
export interface ExpandedNotebook extends Notebook {
  readonly sections: Section[];
  readonly sectionGroups: ExpandedSectionGroup[];
}

/**
 * One GET against Graph with a bearer token.
 *
 * @throws {GraphRequestError} on any non-2xx response, carrying the status and body.
 * @throws {GraphResponseError} if a 2xx body is not a JSON object.
 */
export async function graphGet(
  url: string,
  accessToken: string,
  fetchImpl: FetchLike,
): Promise<Record<string, unknown>> {
  const response = await fetchImpl(url, { headers: { Authorization: `Bearer ${accessToken}` } });

  if (!response.ok) {
    // The body is read before it is thrown away: Graph puts the actual reason in it,
    // and error 20266 is only distinguishable from any other 400 by that text.
    throw new GraphRequestError(
      url,
      response.status,
      response.statusText,
      await safeText(response),
      parseRetryAfter(response.headers.get('retry-after')),
    );
  }

  let body: unknown;
  try {
    body = await response.json();
  } catch (err) {
    throw new GraphResponseError(
      `GET ${url} returned ${response.status} with a body that is not JSON: ${describeError(err)}`,
      url,
    );
  }

  if (typeof body !== 'object' || body === null || Array.isArray(body)) {
    throw new GraphResponseError(`GET ${url} returned JSON that is not an object.`, url);
  }

  return body as Record<string, unknown>;
}

/**
 * Read a whole OData collection, following `@odata.nextLink` until it stops appearing.
 *
 * Graph decides its own page size and ignores a `$top` larger than that size, so a
 * single response is never proof that a collection is complete. `limit` stops the walk
 * early once enough items are in hand, which is what a caller asking for the most
 * recent N pages wants.
 */
async function collectValues(
  firstUrl: string,
  accessToken: string,
  fetchImpl: FetchLike,
  limit?: number,
  gate: RequestGate = UNGATED,
): Promise<unknown[]> {
  const items: unknown[] = [];
  let url: string | undefined = firstUrl;

  for (let follows = 0; url !== undefined; follows += 1) {
    if (follows >= MAX_PAGE_FOLLOWS) {
      throw new GraphResponseError(
        `GET ${firstUrl} kept returning an @odata.nextLink after ${MAX_PAGE_FOLLOWS} pages, so the listing was abandoned.`,
        firstUrl,
      );
    }

    // Gated per request rather than per listing: following @odata.nextLink issues
    // several requests and each counts against the limits separately. The gate also
    // holds the retry, so a 429 on page three does not lose the first two.
    const body: Record<string, unknown> = await gate.run(() =>
      graphGet(url as string, accessToken, fetchImpl),
    );

    const value = body['value'];
    if (!Array.isArray(value)) {
      throw new GraphResponseError(`GET ${url} returned no "value" array.`, url);
    }
    items.push(...value);

    if (limit !== undefined && items.length >= limit) return items.slice(0, limit);

    const next: unknown = body['@odata.nextLink'];
    if (next === undefined || next === null) break;
    if (typeof next !== 'string' || next === '') {
      throw new GraphResponseError(`GET ${url} returned an unusable @odata.nextLink.`, url);
    }
    // The link is an absolute URL Graph built, including its own paging cursor. It is
    // followed verbatim; rebuilding it from parts loses the cursor.
    url = next;
  }

  return limit === undefined ? items : items.slice(0, limit);
}

/**
 * The structure half of the Graph client: everything above page content.
 *
 * A token is fetched per request rather than held, because `GraphAuth` already serves a
 * still-valid access token from MSAL's in-memory cache without a network call, and it
 * is the only thing that knows when the token has expired.
 */
export class GraphStructure {
  readonly #tokens: TokenSource;
  readonly #fetch: FetchLike;
  readonly #gate: RequestGate;

  /**
   * `gate` defaults to UNGATED so a test runs at full speed. The deployed server builds
   * this through `createGraphStructure`, which passes the shared production gate — that
   * is what keeps the process inside OneNote's 5 concurrent and 120-per-minute limits.
   */
  constructor(
    tokens: TokenSource,
    fetchImpl: FetchLike = globalThis.fetch,
    gate: RequestGate = UNGATED,
  ) {
    this.#tokens = tokens;
    this.#fetch = fetchImpl;
    this.#gate = gate;
  }

  /** Every notebook the signed-in account can see, by display name. */
  async listNotebooks(): Promise<Notebook[]> {
    const url = `${GRAPH_ROOT}/me/onenote/notebooks?${NODE_SELECT}`;
    return (await this.#collect(url)).map((item) => toNode(item, url));
  }

  /** The sections directly under a notebook or a section group. */
  async listSections(containerKind: ContainerKind, containerId: string): Promise<Section[]> {
    const url = `${this.#containerUrl(containerKind, containerId)}/sections?${NODE_SELECT}`;
    return (await this.#collect(url)).map((item) => toNode(item, url));
  }

  /** The section groups directly under a notebook or a section group. */
  async listSectionGroups(
    containerKind: ContainerKind,
    containerId: string,
  ): Promise<SectionGroup[]> {
    const url = `${this.#containerUrl(containerKind, containerId)}/sectionGroups?${NODE_SELECT}`;
    return (await this.#collect(url)).map((item) => toNode(item, url));
  }

  /**
   * Both child relationships of one container, in one request.
   *
   * `$expand` on the container itself replaces the two list calls this used to make. On
   * a walk that visits every container, that halves the request count, and requests are
   * the scarce thing: the budget is 400 an hour.
   *
   * @throws {GraphRequestError} on a non-2xx response.
   * @throws {GraphResponseError} if the body is not the expected shape.
   */
  async listContainerChildren(
    containerKind: ContainerKind,
    containerId: string,
  ): Promise<ContainerChildren> {
    const url = containerChildrenUrl(containerKind, containerId);
    const body = await this.#get(url);
    return {
      sections: toNodeArray(body['sections'], url).map((item) => toNode(item, url)),
      sectionGroups: toNodeArray(body['sectionGroups'], url).map((item) => toNode(item, url)),
    };
  }

  /**
   * Pages in one section, most recently modified first, at most `top` of them.
   *
   * Scoped to the section on purpose. The account-wide list is the endpoint that fails
   * with error 20266 on this account's structure.
   */
  async listPagesInSection(sectionId: string, top = 50): Promise<PageSummary[]> {
    if (!Number.isInteger(top) || top < 1) {
      throw new RangeError(`top must be a positive integer, got ${String(top)}`);
    }

    // Graph rejects a $top above 100 outright — 400 with code 20129 — so the request
    // asks for at most that and `collectValues` follows @odata.nextLink until `top`
    // items are in hand. `top` is therefore a result count, not a page size.
    const url =
      `${GRAPH_ROOT}/me/onenote/sections/${encodeURIComponent(sectionId)}/pages` +
      `?$top=${Math.min(top, MAX_GRAPH_TOP)}` +
      `&$orderby=lastModifiedDateTime desc&$select=id,title,lastModifiedDateTime`;

    return (await this.#collect(url, top)).map((item) => toPageSummary(item, url));
  }

  /**
   * The whole tree under one notebook, recursing through nested section groups.
   *
   * Section groups are the UI's "tab groups", and this account nests them, so a walk
   * that stopped at the notebook's direct children would silently miss most sections.
   *
   * @throws {GraphResponseError} if the nesting exceeds MAX_SECTION_GROUP_DEPTH, which
   * means a container reported itself, directly or through a chain, as its own child.
   */
  async getNotebookTree(notebook: Notebook): Promise<NotebookTree> {
    const { sections, sectionGroups } = await this.listContainerChildren(
      'notebooks',
      notebook.id,
    );

    return {
      id: notebook.id,
      displayName: notebook.displayName,
      sections,
      sectionGroups: await this.#expandGroups(sectionGroups, 1),
    };
  }

  /**
   * Sections anywhere in the account whose name contains `displayName`, with parents.
   *
   * Substring rather than equality, because equality with `tolower()` is refused by this
   * endpoint. The caller narrows the result to a full match; this is the widest query the
   * service will answer case-insensitively.
   *
   * @throws {GraphRequestError} on a non-2xx response.
   */
  async findSectionsByName(displayName: string): Promise<SectionWithParents[]> {
    const filter = `contains(tolower(displayName), ${quoteOData(displayName.trim().toLowerCase())})`;
    const url = `${SECTIONS_BY_NAME_URL}${encodeURIComponent(filter)}`;
    return (await this.#collect(url)).map((item) => toSectionWithParents(item, url));
  }

  /**
   * Pages in one section whose title matches, compared case-insensitively by Graph.
   *
   * `tolower(title)` is accepted here even though it is rejected on sections, so this
   * replaces reading every title in the section and comparing them locally — which was
   * bounded at 100 by Graph's own `$top` ceiling and could therefore miss a match.
   *
   * @throws {GraphRequestError} on a non-2xx response.
   */
  async findPagesByTitle(sectionId: string, title: string): Promise<PageSummary[]> {
    const filter = `tolower(title) eq ${quoteOData(title.trim().toLowerCase())}`;
    const url =
      `${GRAPH_ROOT}/me/onenote/sections/${encodeURIComponent(sectionId)}/pages` +
      `?$select=id,title,lastModifiedDateTime&$filter=${encodeURIComponent(filter)}`;
    return (await this.#collect(url)).map((item) => toPageSummary(item, url));
  }

  /**
   * Pages in one section whose title contains `query`, compared case-insensitively.
   *
   * The comparison happens in the service, so a section holding thousands of pages costs
   * the same one request as a section holding three, and nothing is truncated on the way
   * back. This is what an unscoped title search runs per section.
   *
   * @throws {GraphRequestError} on a non-2xx response.
   */
  async findPagesMatchingTitle(sectionId: string, query: string): Promise<PageSummary[]> {
    const filter = `contains(tolower(title), ${quoteOData(query.trim().toLowerCase())})`;
    const url =
      `${GRAPH_ROOT}/me/onenote/sections/${encodeURIComponent(sectionId)}/pages` +
      `?$select=id,title,lastModifiedDateTime&$orderby=${encodeURIComponent('lastModifiedDateTime desc')}` +
      `&$top=${MAX_GRAPH_TOP}&$filter=${encodeURIComponent(filter)}`;
    return (await this.#collect(url, MAX_GRAPH_TOP)).map((item) => toPageSummary(item, url));
  }

  /**
   * Every notebook with its sections and one level of section group, in one request.
   *
   * This is the cheap way to answer a question about names. It does not replace
   * `getFullTree`: a section group nested inside a section group is not in the response,
   * and its sections are not either. A caller that finds nothing here and needs to be
   * sure has to walk the notebook it cares about.
   *
   * @throws {GraphRequestError} on a non-2xx response.
   * @throws {GraphResponseError} if the body is not the expected shape.
   */
  async getExpandedTree(): Promise<ExpandedNotebook[]> {
    const items = await this.#collect(EXPANDED_TREE_URL);
    return items.map((item) => toExpandedNotebook(item, EXPANDED_TREE_URL));
  }

  /**
   * Every notebook, each with its tree resolved.
   *
   * Built on `getExpandedTree`, so the notebooks, their sections, their section groups
   * and those groups' sections all arrive in the first request. What remains is one
   * request per section group, to find whether anything nests below it — Graph caps
   * `$expand` at two levels, so nothing cheaper answers that. On the account this was
   * measured against that is 44 requests rather than the 195 the per-container walk cost.
   *
   * Prefer `getExpandedTree` when one level of section group is enough. This exists for
   * the case where it is not.
   */
  async getFullTree(): Promise<NotebookTree[]> {
    const expanded = await this.getExpandedTree();

    return mapWithLimit(expanded, WALK_CONCURRENCY, async (notebook) => ({
      id: notebook.id,
      displayName: notebook.displayName,
      sections: notebook.sections,
      sectionGroups: await this.#expandGroups(notebook.sectionGroups, 1),
    }));
  }

  async #expandGroups(groups: SectionGroup[], depth: number): Promise<SectionGroupNode[]> {
    if (groups.length === 0) return [];

    if (depth > MAX_SECTION_GROUP_DEPTH) {
      throw new GraphResponseError(
        `Section groups nested deeper than ${MAX_SECTION_GROUP_DEPTH} levels, which Graph does not produce; the walk was abandoned rather than followed into a cycle.`,
        `${GRAPH_ROOT}/me/onenote/sectionGroups`,
      );
    }

    // One request per group, and at most WALK_CONCURRENCY of them in flight. The
    // request is what says whether anything nests below this group; its sections come
    // back in the same response.
    return mapWithLimit(groups, WALK_CONCURRENCY, async (group): Promise<SectionGroupNode> => {
      const children = await this.listContainerChildren('sectionGroups', group.id);
      return {
        id: group.id,
        displayName: group.displayName,
        sections: children.sections,
        sectionGroups: await this.#expandGroups(children.sectionGroups, depth + 1),
      };
    });
  }

  #containerUrl(containerKind: ContainerKind, containerId: string): string {
    return `${GRAPH_ROOT}/me/onenote/${containerKind}/${encodeURIComponent(containerId)}`;
  }

  /** One entity, gated. */
  async #get(url: string): Promise<Record<string, unknown>> {
    const accessToken = await this.#tokens.getAccessToken();
    return this.#gate.run(() => graphGet(url, accessToken, this.#fetch));
  }

  /** One collection, with every request it makes passing through the gate. */
  async #collect(url: string, limit?: number): Promise<unknown[]> {
    const accessToken = await this.#tokens.getAccessToken();
    return collectValues(url, accessToken, this.#fetch, limit, this.#gate);
  }
}

/**
 * The gate every client built by a factory in this repository shares.
 *
 * One per process, because the limits are per app per user and two clients pacing
 * themselves independently would together exceed what either one allows.
 */
export const PRODUCTION_GATE: RequestGate = createGate();

/** Build the client from the server's Graph auth, sharing the process-wide gate. */
export function createGraphStructure(auth: GraphAuth): GraphStructure {
  return new GraphStructure(auth, globalThis.fetch, PRODUCTION_GATE);
}

function toNode(item: unknown, url: string): { id: string; displayName: string } {
  const record = asRecord(item, url);
  return {
    id: requireString(record, 'id', url),
    // A notebook or section can come back with a null displayName; an empty string is
    // preferable to `undefined` reaching a caller that only formats it.
    displayName: optionalString(record, 'displayName') ?? '',
  };
}

/**
 * One notebook out of the expanded response.
 *
 * An absent `sections` or `sectionGroups` is read as empty rather than raised on: Graph
 * omits an expanded relationship that holds nothing, and a notebook with no section
 * groups is ordinary.
 */
function toExpandedNotebook(item: unknown, url: string): ExpandedNotebook {
  const record = asRecord(item, url);
  return {
    ...toNode(record, url),
    sections: toNodeArray(record['sections'], url).map((node) => toNode(node, url)),
    sectionGroups: toNodeArray(record['sectionGroups'], url).map((group) => {
      const groupRecord = asRecord(group, url);
      return {
        ...toNode(groupRecord, url),
        sections: toNodeArray(groupRecord['sections'], url).map((node) => toNode(node, url)),
      };
    }),
  };
}

/** An expanded relationship: absent means empty, anything but an array is a fault. */
function toNodeArray(value: unknown, url: string): unknown[] {
  if (value === undefined || value === null) return [];
  if (!Array.isArray(value)) {
    throw new GraphResponseError(
      `GET ${url} returned an expanded relationship that is not an array.`,
      url,
    );
  }
  return value;
}

/**
 * `Promise.all` with a ceiling on how many run at once, preserving input order.
 *
 * Written out rather than pulled in as a dependency because it is fifteen lines and the
 * thing it prevents — an unbounded fan-out against an API that allows five concurrent
 * requests — is the most expensive mistake this repository has made.
 */
async function mapWithLimit<T, R>(
  items: readonly T[],
  limit: number,
  run: (item: T) => Promise<R>,
): Promise<R[]> {
  const results: R[] = new Array<R>(items.length);
  let cursor = 0;

  const worker = async (): Promise<void> => {
    for (;;) {
      const index = cursor;
      if (index >= items.length) return;
      cursor += 1;
      results[index] = await run(items[index] as T);
    }
  };

  await Promise.all(Array.from({ length: Math.min(limit, items.length) }, () => worker()));
  return results;
}

/** A single-quoted OData string literal; an embedded quote is doubled. */
export function quoteOData(value: string): string {
  return `'${value.replaceAll("'", "''")}'`;
}

function toSectionWithParents(item: unknown, url: string): SectionWithParents {
  const record = asRecord(item, url);
  const notebook = record['parentNotebook'];
  const group = record['parentSectionGroup'];
  return {
    ...toNode(record, url),
    // A section directly under a notebook has no parent section group, and Graph returns
    // null rather than omitting it. Both readings end as null here.
    parentNotebook: isRecord(notebook) ? toNode(notebook, url) : null,
    parentSectionGroup: isRecord(group) ? toNode(group, url) : null,
  };
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function toPageSummary(item: unknown, url: string): PageSummary {
  const record = asRecord(item, url);
  return {
    id: requireString(record, 'id', url),
    title: optionalString(record, 'title') ?? '',
    lastModifiedDateTime: optionalString(record, 'lastModifiedDateTime') ?? '',
  };
}

function asRecord(item: unknown, url: string): Record<string, unknown> {
  if (typeof item !== 'object' || item === null || Array.isArray(item)) {
    throw new GraphResponseError(`GET ${url} returned a non-object inside "value".`, url);
  }
  return item as Record<string, unknown>;
}

function requireString(record: Record<string, unknown>, key: string, url: string): string {
  const value = record[key];
  if (typeof value !== 'string' || value === '') {
    // The offending value is not printed: these bodies carry user content, and the
    // repository's hygiene rules keep page and notebook names out of anything loggable.
    throw new GraphResponseError(`GET ${url} returned an item with no usable "${key}".`, url);
  }
  return value;
}

function optionalString(record: Record<string, unknown>, key: string): string | undefined {
  const value = record[key];
  return typeof value === 'string' ? value : undefined;
}

/** An error body that cannot be read must not mask the status that caused the throw. */
async function safeText(response: Response): Promise<string> {
  try {
    return await response.text();
  } catch {
    return '';
  }
}

function truncate(text: string, max: number): string {
  return text.length <= max ? text : `${text.slice(0, max)}… (${text.length} chars)`;
}

function describeError(err: unknown): string {
  if (err instanceof Error) return `${err.name}: ${err.message}`;
  return String(err);
}
