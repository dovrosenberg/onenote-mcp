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
// Decoding lives in ./graph-decode.ts. That module imports GraphResponseError and the
// result types from here, so the dependency runs one way and this file keeps the URLs,
// the requests and the error types.
import {
  asRecord,
  describeError,
  mapWithLimit,
  quoteOData,
  requireString,
  safeText,
  toExpandedNotebook,
  toNode,
  toNodeArray,
  toPageSummary,
  toSectionWithParents,
  truncate,
} from './graph-decode.ts';

export const GRAPH_ROOT = 'https://graph.microsoft.com/v1.0';

/**
 * The only origin this module will send the access token to.
 *
 * `@odata.nextLink` is followed verbatim, because it carries Graph's own paging cursor
 * and cannot be rebuilt from parts. That means a URL out of a response body decides
 * where the next request goes, and every request carries the Graph bearer token in an
 * Authorization header. A link naming another host would send that token to it. Graph is
 * the one that writes these links, so this check is not expected to fire; it is here so
 * that the token cannot leave this origin whether or not that stays true.
 */
const GRAPH_ORIGIN = new URL(GRAPH_ROOT).origin;

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
  `${GRAPH_ROOT}/me/onenote/notebooks?$select=id,displayName,lastModifiedDateTime` +
  `&$expand=sections($select=id,displayName,lastModifiedDateTime),` +
  `sectionGroups($select=id,displayName;$expand=sections($select=id,displayName,lastModifiedDateTime))`;

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
  /** The verb, so a failed write does not report itself as a failed read. */
  readonly method: string;

  constructor(
    url: string,
    status: number,
    statusText: string,
    body: string,
    retryAfterMs?: number,
    method = 'GET',
  ) {
    super(
      `${method} ${url} failed: ${status} ${statusText}${body === '' ? '' : ` ${truncate(body, MAX_BODY_CHARS)}`}`,
    );
    this.name = 'GraphRequestError';
    this.url = url;
    this.status = status;
    this.statusText = statusText;
    this.body = body;
    this.retryAfterMs = retryAfterMs;
    this.method = method;
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

/**
 * The slice of `fetch` this repository uses, so a test can supply its own.
 *
 * Every call in this module is a GET and passes only `headers`. `method` and `body` are
 * here for ./page-write.ts, which PATCHes and POSTs through the same type so one fake
 * fetch in a test can serve both halves of the client.
 */
export type FetchLike = (
  url: string,
  init: { headers: Record<string, string>; method?: string; body?: string },
) => Promise<Response>;

/** `notebooks` and `sectionGroups` expose the same two child relationships. */
export type ContainerKind = 'notebooks' | 'sectionGroups';

export interface Notebook {
  readonly id: string;
  readonly displayName: string;
  /**
   * Present only on the expanded-tree call, which is the one URL that asks for it.
   * `NODE_SELECT` deliberately does not, so `list_notebooks` and `list_sections` keep
   * the URLs their tests assert.
   */
  readonly lastModifiedDateTime?: string;
}

export interface Section {
  readonly id: string;
  readonly displayName: string;
  /**
   * Present only on the expanded-tree call. Measured 2026-08-19 (api-overview.md): this
   * moves when a page in the section is created, edited or deleted, and does not move
   * otherwise — which is what lets the mirror's incremental sync visit only the sections
   * that changed instead of all of them.
   */
  readonly lastModifiedDateTime?: string;
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
/** Is this an absolute URL on the Graph origin? A relative or malformed URL is not. */
function isGraphUrl(value: string): boolean {
  try {
    return new URL(value).origin === GRAPH_ORIGIN;
  } catch {
    return false;
  }
}

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
    if (!isGraphUrl(next)) {
      // The link itself is not quoted. It came from a response body, and this repository's
      // output can reach a public log.
      throw new GraphResponseError(
        `GET ${url} returned an @odata.nextLink pointing somewhere other than ${GRAPH_ORIGIN}, so it was not followed.`,
        url,
      );
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
   * Pages in one section modified at or after `sinceIso`, newest first.
   *
   * The mirror's second tier (issue #30). One request per section whose timestamp moved,
   * against `listPagesInSection`'s "most recent N whatever their age".
   *
   * `$filter=lastModifiedDateTime ge {iso}` with an unquoted ISO-8601 UTC literal was
   * measured accepted on 2026-08-19 (api-overview.md). If the service ever stops
   * accepting it the answer is a 400, and the caller's fallback is
   * `listPagesInSection` plus a client-side cutoff — the documented default sort for a
   * section's pages is already `lastModifiedTime desc`, confirmed in the same run, so
   * that fallback costs the same one request. No `$orderby` is sent here for that
   * reason: it would be a second unverified query option on a call whose filter is the
   * thing under suspicion.
   *
   * `sinceIso` must already carry the caller's overlap. This method does no date
   * arithmetic — the watermark rule lives with the sync, which is the only thing that
   * knows when its pass started.
   */
  async listPagesChangedSince(sectionId: string, sinceIso: string): Promise<PageSummary[]> {
    const url =
      `${GRAPH_ROOT}/me/onenote/sections/${encodeURIComponent(sectionId)}/pages` +
      `?$select=id,title,lastModifiedDateTime&$top=${MAX_GRAPH_TOP}` +
      `&$filter=${encodeURIComponent(`lastModifiedDateTime ge ${sinceIso}`)}`;

    return (await this.#collect(url)).map((item) => toPageSummary(item, url));
  }

  /**
   * Every page id in one section, and nothing else.
   *
   * This is the mirror's deletion sweep, and it is as cheap as deletion detection gets
   * on this account. Graph has no /delta on any OneNote resource and no tombstone for a
   * deleted page, and the account-wide page list — the one call that would enumerate
   * everything in one request per 100 pages — is the banned one, error 20266. So the
   * floor is one request per section plus one per additional 100 pages, and what makes
   * that affordable is that only the mirrored notebooks are swept.
   *
   * `$select=id` alone, because the caller compares id sets and reads nothing else. No
   * `top` argument: a sweep that stopped early would report pages as deleted that are
   * merely past the cutoff, which is the one mistake here that destroys data.
   */
  async listPageIds(sectionId: string): Promise<string[]> {
    const url =
      `${GRAPH_ROOT}/me/onenote/sections/${encodeURIComponent(sectionId)}/pages` +
      `?$select=id&$top=${MAX_GRAPH_TOP}`;

    return (await this.#collect(url)).map((item) => requireString(asRecord(item, url), 'id', url));
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
