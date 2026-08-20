// The write half of the Graph client: appending to a page, creating one, and replacing
// a title.
//
// OneNote does not have a text append. Changing an existing page is
// `PATCH /me/onenote/pages/{id}/content` with a JSON array of change objects, each
// naming a target element, an action, and the HTML to apply. Success is 204 with an
// empty body. The three shapes this module sends were measured against the live service
// on 2026-08-18 by the spike in issue #17; the full record, including every error code,
// is in `api-overview.md` under **Writing page content**.
//
// Four things here follow from that spike rather than from the documentation:
//
// `target: "body"` is the first top-level div, not the page. A page authored in the
// OneNote client has several sibling top-level divs, so an append lands at the end of the
// **first outline**. Reaching another one means reading `?includeIDs=true` and targeting
// that div's generated id, which is issue #27's problem, not this module's.
//
// A page created here deliberately omits `data-absolute-enabled` from its `<body>`. Graph
// then wraps the whole submission in one `<div data-id="_default">`, which makes `body`
// cover the entire page — so `append_to_page` appends to the bottom of a page this server
// created, and to the first outline of a page a human wrote in the client. Setting the
// attribute would give sibling outlines matching a client-authored page and make the two
// cases behave the same badly rather than the same well.
//
// The title is stored verbatim. `content` is not parsed: `<p>x</p>` produces a page whose
// title reads `<p>x</p>`, and `&`, `<`, `>` and `"` survive unescaped. So `updatePageTitle`
// escapes nothing. `createPage` is the opposite case — the title there is an element in a
// submitted HTML document and has to be escaped, which `escapeHtml` below does.
//
// The change array is applied as a unit. One change naming a missing target fails the
// whole request with 400 code 20120 and applies none of the others, so a caller that
// wants two edits in one call can rely on all-or-nothing.

import {
  GRAPH_ROOT,
  GraphRequestError,
  GraphResponseError,
  PRODUCTION_GATE,
  type FetchLike,
  type TokenSource,
  withRequestTimeout,
} from './graph-structure.ts';
import { UNGATED, parseRetryAfter, type RequestGate } from './graph-throttle.ts';
import type { GraphAuth } from './graph-auth.ts';

/** One entry in a PATCH change array. */
export interface PageChange {
  /** `body`, `title`, a generated id, or `#{data-id}`. `title` and `body` take no `#`. */
  readonly target: string;
  readonly action: 'append' | 'insert' | 'prepend' | 'replace';
  readonly content: string;
  /** `before` or `after`; Graph defaults to `after` when it is absent. */
  readonly position?: 'before' | 'after';
}

/** What `createPage` returns: enough to read the page back, and a link to open it. */
export interface CreatedPage {
  readonly id: string;
  /** The title Graph echoed. The create response carries it even when a later
   *  `GET /pages/{id}?$select=title` still answers `""`. */
  readonly title: string;
  /** `oneNoteWebUrl`, which opens the page in a browser. Null if Graph omitted it. */
  readonly webUrl: string | null;
  /** `oneNoteClientUrl`, which opens it in the desktop client. Null if Graph omitted it. */
  readonly clientUrl: string | null;
}

/** The URL a change array is PATCHed to. */
export function pageContentPatchUrl(pageId: string): string {
  return `${GRAPH_ROOT}/me/onenote/pages/${encodeURIComponent(pageId)}/content`;
}

/** The URL a new page is POSTed to. */
export function sectionPagesUrl(sectionId: string): string {
  return `${GRAPH_ROOT}/me/onenote/sections/${encodeURIComponent(sectionId)}/pages`;
}

/**
 * The document submitted to create a page.
 *
 * The title is an element here rather than a header in the body, because everything that
 * finds a page by name — `find_page_by_name`, `list_pages_by_name`, `search_pages` —
 * matches on the page title Graph stores, and that comes from `<title>`. A page whose
 * name exists only as an `<h1>` is unfindable by name.
 *
 * `<body>` carries no `data-absolute-enabled`; see the note at the top of this file.
 */
export function createPageHtml(title: string, bodyHtml: string): string {
  return (
    '<!DOCTYPE html>\n' +
    '<html>\n' +
    '  <head>\n' +
    `    <title>${escapeHtml(title)}</title>\n` +
    '  </head>\n' +
    `  <body>${bodyHtml}</body>\n` +
    '</html>\n'
  );
}

/** The five characters that would otherwise close or open a tag in a submitted document. */
export function escapeHtml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/**
 * The write half of the Graph client.
 *
 * Kept apart from `GraphStructure` because nothing here is a GET and nothing here reads
 * an OData collection: two of the three calls answer 204 with no body at all.
 */
export class GraphPageWrite {
  readonly #tokens: TokenSource;
  readonly #fetch: FetchLike;
  readonly #gate: RequestGate;

  /**
   * `gate` defaults to UNGATED so a test runs at full speed. `createGraphPageWrite`
   * passes the process-wide gate: the OneNote limits are per app per user and count
   * writes alongside reads, so a write that paced itself separately would push the
   * reads over.
   *
   * Retrying is safe for all three calls even though a POST creates something. The gate
   * retries only 429 and 503, and a throttled request was refused rather than performed.
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

  /**
   * Apply a change array to one page.
   *
   * @throws {GraphRequestError} on any non-2xx response.
   * @throws {RangeError} if the array is empty — Graph answers 400 code 20125 to that,
   * and spending a request to learn it is pointless.
   */
  async patchPage(pageId: string, changes: readonly PageChange[]): Promise<void> {
    if (changes.length === 0) {
      throw new RangeError('a PATCH needs at least one change; Graph rejects an empty array');
    }

    const url = pageContentPatchUrl(pageId);
    const accessToken = await this.#tokens.getAccessToken();

    await this.#gate.run(async () => {
      const response = await this.#fetch(url, {
        method: 'PATCH',
        headers: {
          Authorization: `Bearer ${accessToken}`,
          'Content-Type': 'application/json',
        },
        body: JSON.stringify(changes),
      });

      if (!response.ok) throw await requestError(url, 'PATCH', response);
      // Success is 204 with an empty body. Nothing is read from it.
    });
  }

  /**
   * Append HTML to the end of a page's body.
   *
   * On a page created by `createPage` that is the bottom of the page. On a page authored
   * in the OneNote client it is the end of the first outline, because that is what `body`
   * addresses there.
   *
   * @throws {GraphRequestError} on any non-2xx response.
   */
  async appendToPage(pageId: string, html: string): Promise<void> {
    await this.patchPage(pageId, [{ target: 'body', action: 'append', content: html }]);
  }

  /**
   * Replace a page's title with `title`, verbatim.
   *
   * `replace` is the only action `title` accepts — `append` is 400 code 20141 — and the
   * target is written `title`, never `#title`, which is 400 code 20149.
   *
   * @throws {GraphRequestError} on any non-2xx response.
   */
  async updatePageTitle(pageId: string, title: string): Promise<void> {
    await this.patchPage(pageId, [{ target: 'title', action: 'replace', content: title }]);
  }

  /**
   * Create a page in one section from a title and a fragment of body HTML.
   *
   * The request is `Content-Type: text/html` rather than the `multipart/form-data` shape
   * the documentation shows: multipart exists to carry binary parts, and this tool takes
   * no binary. A page needing an embedded image is not something this server writes.
   *
   * @throws {GraphRequestError} on any non-2xx response.
   * @throws {GraphResponseError} if the 2xx body is not a page object with an id.
   */
  async createPage(sectionId: string, title: string, bodyHtml: string): Promise<CreatedPage> {
    const url = sectionPagesUrl(sectionId);
    const accessToken = await this.#tokens.getAccessToken();

    return this.#gate.run(async () => {
      const response = await this.#fetch(url, {
        method: 'POST',
        headers: {
          Authorization: `Bearer ${accessToken}`,
          // The charset is explicit: the body is a JavaScript string, `fetch` sends it
          // as UTF-8, and a page whose title holds a non-ASCII character comes back
          // mangled if the service guesses something else.
          'Content-Type': 'text/html; charset=utf-8',
        },
        body: createPageHtml(title, bodyHtml),
      });

      if (!response.ok) throw await requestError(url, 'POST', response);

      let body: unknown;
      try {
        body = await response.json();
      } catch (err) {
        throw new GraphResponseError(
          `POST ${url} returned ${response.status} with a body that is not JSON: ${describeError(err)}`,
          url,
        );
      }

      return toCreatedPage(body, url);
    });
  }
}

/** Build the client from the server's Graph auth, sharing the process-wide gate. */
export function createGraphPageWrite(auth: GraphAuth): GraphPageWrite {
  return new GraphPageWrite(auth, withRequestTimeout(globalThis.fetch), PRODUCTION_GATE);
}

/**
 * The error for a failed write, with the body read before it is thrown away.
 *
 * The body is where the PATCH codes live — 20120 for a target that cannot be located,
 * 20134 for one that is not updateable — and those are the only thing that distinguishes
 * one 400 from another.
 */
async function requestError(
  url: string,
  method: string,
  response: Response,
): Promise<GraphRequestError> {
  return new GraphRequestError(
    url,
    response.status,
    response.statusText,
    await safeText(response),
    parseRetryAfter(response.headers.get('retry-after')),
    method,
  );
}

/**
 * The created page out of Graph's response.
 *
 * Only `id` is required. The links are optional because a missing link is not a reason to
 * fail a create that already happened — the page exists either way, and a tool that threw
 * here would report a failure the caller cannot act on.
 */
function toCreatedPage(body: unknown, url: string): CreatedPage {
  if (typeof body !== 'object' || body === null || Array.isArray(body)) {
    throw new GraphResponseError(`POST ${url} returned JSON that is not an object.`, url);
  }

  const record = body as Record<string, unknown>;
  const id = record['id'];
  if (typeof id !== 'string' || id === '') {
    throw new GraphResponseError(
      `POST ${url} succeeded but the response carried no page id, so the new page cannot be named.`,
      url,
    );
  }

  return {
    id,
    title: typeof record['title'] === 'string' ? record['title'] : '',
    webUrl: linkHref(record['links'], 'oneNoteWebUrl'),
    clientUrl: linkHref(record['links'], 'oneNoteClientUrl'),
  };
}

/** `links: { oneNoteWebUrl: { href: "…" } }`, or null if any part of that is missing. */
function linkHref(links: unknown, name: string): string | null {
  if (typeof links !== 'object' || links === null) return null;
  const link = (links as Record<string, unknown>)[name];
  if (typeof link !== 'object' || link === null) return null;
  const href = (link as Record<string, unknown>)['href'];
  return typeof href === 'string' && href !== '' ? href : null;
}

/** An unreadable body must not mask the status that is about to be thrown. */
async function safeText(response: Response): Promise<string> {
  try {
    return await response.text();
  } catch {
    return '';
  }
}

function describeError(err: unknown): string {
  return err instanceof Error ? err.message : String(err);
}
