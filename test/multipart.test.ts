import test from 'node:test';
import assert from 'node:assert/strict';

import { findBoundary, findPart, splitMultipart, UNKNOWN_CONTENT_TYPE } from '../src/multipart.ts';

const BOUNDARY = 'MTk2MjA1Nzg0NzE1NjA2ODU4NA';

/** A response shaped like Graph's: an HTML part, then an InkML part. */
function graphBody(sep: string): string {
  return [
    `--${BOUNDARY}`,
    'Content-Type: text/html',
    '',
    '<html><body><p>typed</p></body></html>',
    `--${BOUNDARY}`,
    'Content-Type: application/inkml+xml',
    '',
    '<inkml:ink />',
    `--${BOUNDARY}--`,
    '',
  ].join(sep);
}

test('the boundary is read from the content type, quoted or not', () => {
  assert.equal(findBoundary(`multipart/mixed; boundary="${BOUNDARY}"`), BOUNDARY);
  assert.equal(findBoundary(`multipart/mixed; boundary=${BOUNDARY}`), BOUNDARY);
  assert.equal(findBoundary(`multipart/mixed; boundary=${BOUNDARY}; charset=utf-8`), BOUNDARY);
  assert.equal(findBoundary('text/html'), null);
  assert.equal(findBoundary(null), null);
  assert.equal(findBoundary(undefined), null);
});

test('a response that is not multipart splits to null, not to an empty list', () => {
  // The distinction is the caller's: null means "search the whole body for ink".
  assert.equal(splitMultipart('<html></html>', 'text/html'), null);
});

for (const [name, sep] of [
  ['CRLF', '\r\n'],
  ['LF', '\n'],
] as const) {
  test(`a ${name}-separated body splits into its parts`, () => {
    const parts = splitMultipart(graphBody(sep), `multipart/mixed; boundary="${BOUNDARY}"`);

    assert.ok(parts !== null);
    assert.equal(parts.length, 2);
    assert.equal(parts[0]?.contentType, 'text/html');
    assert.equal(parts[0]?.body, '<html><body><p>typed</p></body></html>');
    assert.equal(parts[1]?.contentType, 'application/inkml+xml');
    assert.equal(parts[1]?.body, '<inkml:ink />');
  });
}

test('the preamble, the epilogue, and the closing delimiter are not parts', () => {
  const body = [
    'This is a multipart message in MIME format.',
    `--${BOUNDARY}`,
    'Content-Type: text/html',
    '',
    '<p>only part</p>',
    `--${BOUNDARY}--`,
    'trailing noise after the closing delimiter',
  ].join('\r\n');

  const parts = splitMultipart(body, `multipart/mixed; boundary=${BOUNDARY}`);

  assert.equal(parts?.length, 1);
  assert.equal(parts?.[0]?.body, '<p>only part</p>');
});

test('a part with no content-type header reports one, and keeps its whole body', () => {
  const body = `--${BOUNDARY}\r\n\r\njust a body\r\n--${BOUNDARY}--`;
  const parts = splitMultipart(body, `multipart/mixed; boundary=${BOUNDARY}`);

  assert.equal(parts?.[0]?.contentType, UNKNOWN_CONTENT_TYPE);
  assert.equal(parts?.[0]?.body, 'just a body');
});

test('a part with no blank line is all body, so no header text leaks into it', () => {
  // Nothing observed from Graph does this. The check exists because the alternative
  // failure is silent: header text parsed as body would end up inside the InkML.
  const body = `--${BOUNDARY}\r\n<p>headerless</p>\r\n--${BOUNDARY}--`;
  const parts = splitMultipart(body, `multipart/mixed; boundary=${BOUNDARY}`);

  assert.equal(parts?.[0]?.headers, '');
  assert.equal(parts?.[0]?.body, '<p>headerless</p>');
});

test('findPart matches on the part content type and keeps body order', () => {
  const parts = splitMultipart(graphBody('\r\n'), `multipart/mixed; boundary=${BOUNDARY}`) ?? [];

  assert.equal(findPart(parts, /html/i)?.body, '<html><body><p>typed</p></body></html>');
  assert.equal(findPart(parts, /inkml/i)?.body, '<inkml:ink />');
  assert.equal(findPart(parts, /image\/png/i), null);
});
