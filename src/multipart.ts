// Splitting a `multipart/mixed` response body into its parts.
//
// This exists for one endpoint: `GET /me/onenote/pages/{id}/content?includeInkML=true`
// answers with `multipart/mixed`, one part being the page HTML and another the InkML
// with the real stroke data. Node's fetch does not split multipart bodies, and pulling
// in a MIME library for a two-part response with no nesting and no transfer encoding is
// not worth the dependency.
//
// What it deliberately does not do: no nested multipart, no `Content-Transfer-Encoding`
// decoding, no binary parts. Graph sends neither here. A part body is returned as the
// text it was given.

/** One part of a multipart body: its raw header block, its type, and its body text. */
export interface MultipartPart {
  /** The part's header block, verbatim and trimmed. */
  readonly headers: string;
  /** The part's `Content-Type` header value, or `unknown` when it has none. */
  readonly contentType: string;
  /** The part's body, trimmed of the newlines that frame it. */
  readonly body: string;
}

/** What a part with no `Content-Type` header reports. */
export const UNKNOWN_CONTENT_TYPE = 'unknown';

/**
 * The boundary token from a `Content-Type` header value, or null if there is none.
 *
 * Graph quotes its boundary; RFC 2046 allows it unquoted, so both are accepted.
 */
export function findBoundary(contentType: string | null | undefined): string | null {
  const match = /boundary="?([^";]+)"?/i.exec(contentType ?? '');
  const boundary = match?.[1]?.trim();
  return boundary === undefined || boundary === '' ? null : boundary;
}

/**
 * Split a multipart body into its parts, or return null if `contentType` names no
 * boundary — which is how a caller learns the response was not multipart at all.
 *
 * The delimiter is `--<boundary>`, the preamble before the first one is dropped, and
 * everything after the closing `--<boundary>--` is dropped. Headers are separated from
 * the body by a blank line, which is `\r\n\r\n` per RFC 2046 but arrives as `\n\n` from
 * anything that has normalised newlines on the way — both are handled, because a part
 * whose header block is mistaken for body text carries `Content-Type:` into the InkML.
 */
export function splitMultipart(
  rawText: string,
  contentType: string | null | undefined,
): MultipartPart[] | null {
  const boundary = findBoundary(contentType);
  if (boundary === null) return null;

  const parts: MultipartPart[] = [];

  // Segment 0 is the preamble before the first delimiter and is never a part.
  for (const segment of rawText.split(`--${boundary}`).slice(1)) {
    // The closing delimiter is `--<boundary>--`, so a segment starting with `--` ends
    // the body; anything after it is the epilogue.
    if (segment.startsWith('--')) break;
    if (segment.trim() === '') continue;
    parts.push(parsePart(segment));
  }

  return parts;
}

/** The first part whose `Content-Type` matches, or null. Order is the body's order. */
export function findPart(parts: MultipartPart[], pattern: RegExp): MultipartPart | null {
  return parts.find((part) => pattern.test(part.contentType)) ?? null;
}

function parsePart(segment: string): MultipartPart {
  const crlf = segment.indexOf('\r\n\r\n');
  const lf = crlf === -1 ? segment.indexOf('\n\n') : -1;
  const bodyStart = crlf !== -1 ? crlf + 4 : lf !== -1 ? lf + 2 : 0;

  const headers = segment.slice(0, bodyStart).trim();

  return {
    headers,
    contentType: /^content-type:\s*([^\r\n]+)/im.exec(headers)?.[1]?.trim() ?? UNKNOWN_CONTENT_TYPE,
    body: segment.slice(bodyStart).trim(),
  };
}
