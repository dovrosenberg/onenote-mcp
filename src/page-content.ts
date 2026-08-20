// Fetching one page's content from Graph, with its handwriting.
//
// `?includeInkML=true` is what makes the difference: without it the response is HTML
// with `<!-- InkNode is not supported -->` where the handwriting was, and the strokes
// are simply gone. With it the response is `multipart/mixed` — one part the same HTML,
// another the InkML that ./ink.ts turns into a picture.
//
// One fetch serves both halves. A caller asking for a page should not pay two round
// trips, and should not have to know that ink is a separate concern.
//
// The trimming itself lives in ./page-html.ts. `fetchRaw` returns the parts as Graph
// sent them; `fetchContent` is the whole page — trimmed HTML and rendered ink — from
// that one response.

import {
  GraphRequestError,
  GRAPH_ROOT,
  PRODUCTION_GATE,
  type FetchLike,
  type TokenSource,
  withRequestTimeout,
} from './graph-structure.ts';
import { UNGATED, parseRetryAfter, type RequestGate } from './graph-throttle.ts';
import { renderInk, DEFAULT_RENDER_WIDTH, type InkImage } from './ink.ts';
import { splitMultipart, findPart, type MultipartPart } from './multipart.ts';
import { trimPageHtml } from './page-html.ts';

/** A page's content as Graph returned it, already split into parts. */
export interface RawPageContent {
  /** The whole response body, before any splitting. */
  readonly raw: string;
  /** The response `Content-Type`, which carries the multipart boundary. */
  readonly contentType: string | null;
  /** The parts, or an empty array when the response was not multipart. */
  readonly parts: MultipartPart[];
}

/** One page: its typed content as trimmed HTML, and its handwriting as a PNG. */
export interface PageContent {
  /** The trimmed HTML, or null when the response carried no HTML part. */
  readonly html: string | null;
  /** The rendered handwriting, or null when the page has none. Null is normal. */
  readonly ink: InkImage | null;
}

/** Anything in a content response that could hold an `<ink>` root. */
export interface InkCandidate {
  /** Where the text came from: `body`, or `part-<n>`. */
  readonly label: string;
  readonly text: string;
}

/**
 * The URL of one page's content, with the handwriting and the element ids included.
 *
 * `includeIDs=true` is what makes Graph emit the generated `id` attributes a PATCH can
 * target. Without it a page authored in the OneNote client comes back with no id and no
 * `data-id` anywhere, so nothing on it is addressable and the write tools in
 * ./page-write.ts can only reach `body` and `title`. The two parameters compose:
 * measured on 2026-08-18, `includeIDs=true&includeInkML=true` answers 200 with
 * `multipart/mixed` and the ids present in the HTML part, so ./multipart.ts is
 * unaffected.
 *
 * Generated ids change whenever the page is updated, so anything that targets one has to
 * read it in the same operation that uses it. They are not cacheable between calls.
 */
export function pageContentUrl(pageId: string): string {
  return (
    `${GRAPH_ROOT}/me/onenote/pages/${encodeURIComponent(pageId)}/content` +
    `?includeIDs=true&includeInkML=true`
  );
}

/**
 * The places `<ink>` might be, in the order they are worth trying.
 *
 * The whole body comes first because a non-multipart response still carries the ink
 * inline, and splitting a body that has no boundary yields nothing to search.
 */
export function inkCandidates(content: RawPageContent): InkCandidate[] {
  return [
    { label: 'body', text: content.raw },
    ...content.parts.map((part, i) => ({ label: `part-${i}`, text: part.body })),
  ];
}

/**
 * Reads page content. Kept apart from GraphStructure because this endpoint returns
 * `multipart/mixed` text rather than the JSON every structure call returns.
 */
export class GraphPageContent {
  readonly #tokens: TokenSource;
  readonly #fetch: FetchLike;
  readonly #gate: RequestGate;

  /**
   * `gate` defaults to UNGATED so a test runs at full speed. `createGraphPageContent`
   * passes the process-wide gate, which is what keeps page reads inside the same
   * concurrency and rate limits the structure calls obey — the limits are per user, so
   * two clients pacing separately would together exceed both.
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
   * One page's content with `includeInkML=true`, split into its multipart parts.
   *
   * @throws {GraphRequestError} on any non-2xx response.
   */
  async fetchRaw(pageId: string): Promise<RawPageContent> {
    const url = pageContentUrl(pageId);
    const accessToken = await this.#tokens.getAccessToken();

    return this.#gate.run(async () => {
      const response = await this.#fetch(url, {
        headers: { Authorization: `Bearer ${accessToken}` },
      });

      const contentType = response.headers.get('content-type');
      const raw = await safeText(response);

      if (!response.ok) {
        // The retry hint goes on the error because the gate is what acts on it; a 429
        // here is retried on the same terms as one from a structure call.
        throw new GraphRequestError(
          url,
          response.status,
          response.statusText,
          raw,
          parseRetryAfter(response.headers.get('retry-after')),
        );
      }

      return { raw, contentType, parts: splitMultipart(raw, contentType) ?? [] };
    });
  }

  /**
   * One page's typed content and handwriting, from a single fetch.
   *
   * @throws {GraphRequestError} on a non-2xx response.
   * @throws {InkRenderError} if rasterisation fails.
   */
  async fetchContent(pageId: string, width: number = DEFAULT_RENDER_WIDTH): Promise<PageContent> {
    const content = await this.fetchRaw(pageId);
    const html = pageHtml(content);
    return { html: html === null ? null : trimPageHtml(html), ink: renderPageInk(content, width) };
  }

  /**
   * The page's handwriting as a PNG, or null when the page has none.
   *
   * Null is the normal answer for a typed page and is not an error.
   *
   * @throws {GraphRequestError} on a non-2xx response.
   * @throws {InkRenderError} if rasterisation fails.
   */
  async fetchInk(pageId: string, width: number = DEFAULT_RENDER_WIDTH): Promise<InkImage | null> {
    return renderPageInk(await this.fetchRaw(pageId), width);
  }
}

/**
 * Render the ink out of an already-fetched response, or null when it holds none.
 *
 * The first candidate that yields strokes wins. Searching on rather than returning the
 * first candidate's answer matters because the whole body is searched first and it
 * contains the InkML part too, so a body-level match and a part-level match are the
 * same strokes drawn twice if both were used.
 */
export function renderPageInk(
  content: RawPageContent,
  width: number = DEFAULT_RENDER_WIDTH,
): InkImage | null {
  for (const candidate of inkCandidates(content)) {
    const image = renderInk(candidate.text, width);
    if (image !== null) return image;
  }
  return null;
}

/**
 * The rendered ink *and* the InkML it came from.
 *
 * `renderPageInk` answers the question the read path asks — "what does this look like" —
 * and throws the source away. The mirror needs the source too: it stores the InkML
 * beside the PNG so a change to MAX_INK_PNG_BYTES can be applied by re-rendering
 * locally instead of re-fetching every inked page from Graph. See the note in
 * ./mirror-blobs.ts.
 */
export function renderPageInkWithSource(
  content: RawPageContent,
  width: number = DEFAULT_RENDER_WIDTH,
): { image: InkImage; inkml: string } | null {
  for (const candidate of inkCandidates(content)) {
    const image = renderInk(candidate.text, width);
    if (image !== null) return { image, inkml: candidate.text };
  }
  return null;
}

/** The HTML part of a content response, or null if the response carries none. */
export function htmlPart(content: RawPageContent): MultipartPart | null {
  return findPart(content.parts, /html/i);
}

/**
 * The page's HTML as Graph sent it, or null when the response carries none.
 *
 * A response that was not multipart at all is the HTML: `includeInkML=true` on a page
 * with no ink comes back as plain `text/html`, and dropping it because there are no
 * parts to search would lose every typed page.
 */
export function pageHtml(content: RawPageContent): string | null {
  const part = htmlPart(content);
  if (part !== null) return part.body;
  return content.parts.length === 0 ? content.raw : null;
}

/** Build the client from the server's Graph auth, sharing the process-wide gate. */
export function createGraphPageContent(tokens: TokenSource): GraphPageContent {
  return new GraphPageContent(tokens, withRequestTimeout(globalThis.fetch), PRODUCTION_GATE);
}

/** An unreadable body must not mask the status that is about to be thrown. */
async function safeText(response: Response): Promise<string> {
  try {
    return await response.text();
  } catch {
    return '';
  }
}
