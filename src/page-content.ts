// Fetching one page's content from Graph, with its handwriting.
//
// `?includeInkML=true` is what makes the difference: without it the response is HTML
// with `<!-- InkNode is not supported -->` where the handwriting was, and the strokes
// are simply gone. With it the response is `multipart/mixed` — one part the same HTML,
// another the InkML that ./ink.ts turns into a picture.
//
// One fetch serves both halves. A caller asking for a page should not pay two round
// trips, and should not have to know that ink is a separate concern; issue #13 builds
// the HTML half on top of this same fetch.
//
// This module holds no HTML trimming. It returns the parts as Graph sent them.

import { GraphRequestError, GRAPH_ROOT, type FetchLike, type TokenSource } from './graph-structure.ts';
import { renderInk, DEFAULT_RENDER_WIDTH, type InkImage } from './ink.ts';
import { splitMultipart, findPart, type MultipartPart } from './multipart.ts';

/** A page's content as Graph returned it, already split into parts. */
export interface RawPageContent {
  /** The whole response body, before any splitting. */
  readonly raw: string;
  /** The response `Content-Type`, which carries the multipart boundary. */
  readonly contentType: string | null;
  /** The parts, or an empty array when the response was not multipart. */
  readonly parts: MultipartPart[];
}

/** Anything in a content response that could hold an `<ink>` root. */
export interface InkCandidate {
  /** Where the text came from: `body`, or `part-<n>`. */
  readonly label: string;
  readonly text: string;
}

/** The URL of one page's content, with the handwriting included. */
export function pageContentUrl(pageId: string): string {
  return `${GRAPH_ROOT}/me/onenote/pages/${encodeURIComponent(pageId)}/content?includeInkML=true`;
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

  constructor(tokens: TokenSource, fetchImpl: FetchLike = globalThis.fetch) {
    this.#tokens = tokens;
    this.#fetch = fetchImpl;
  }

  /**
   * One page's content with `includeInkML=true`, split into its multipart parts.
   *
   * @throws {GraphRequestError} on any non-2xx response.
   */
  async fetchRaw(pageId: string): Promise<RawPageContent> {
    const url = pageContentUrl(pageId);
    const accessToken = await this.#tokens.getAccessToken();
    const response = await this.#fetch(url, { headers: { Authorization: `Bearer ${accessToken}` } });

    const contentType = response.headers.get('content-type');
    const raw = await safeText(response);

    if (!response.ok) {
      throw new GraphRequestError(url, response.status, response.statusText, raw);
    }

    return { raw, contentType, parts: splitMultipart(raw, contentType) ?? [] };
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

/** The HTML part of a content response, or null if the response carries none. */
export function htmlPart(content: RawPageContent): MultipartPart | null {
  return findPart(content.parts, /html/i);
}

/** Build the client from the server's Graph auth. */
export function createGraphPageContent(tokens: TokenSource): GraphPageContent {
  return new GraphPageContent(tokens);
}

/** An unreadable body must not mask the status that is about to be thrown. */
async function safeText(response: Response): Promise<string> {
  try {
    return await response.text();
  } catch {
    return '';
  }
}
