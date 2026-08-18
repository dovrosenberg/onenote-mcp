import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';

import { GraphRequestError, type FetchLike, type TokenSource } from '../src/graph-structure.ts';
import {
  GraphPageContent,
  htmlPart,
  inkCandidates,
  pageContentUrl,
  pageHtml,
  renderPageInk,
} from '../src/page-content.ts';

const TOKEN = 'fake-access-token';
const PAGE_ID = '1-abc!123';
const BOUNDARY = 'MTk2MjA1Nzg0NzE1NjA2ODU4NA';

const tokens: TokenSource = { getAccessToken: () => Promise.resolve(TOKEN) };

interface Call {
  url: string;
  authorization: string | undefined;
}

/** A fetch keyed by exact URL, so every behavioural assertion also asserts the URL. */
function fakeFetch(routes: Record<string, () => Response>): {
  fetchImpl: FetchLike;
  calls: Call[];
} {
  const calls: Call[] = [];
  const fetchImpl: FetchLike = (url, init) => {
    calls.push({ url, authorization: init.headers['Authorization'] });
    const route = routes[url];
    if (route === undefined) return Promise.reject(new Error(`no route for ${url}`));
    return Promise.resolve(route());
  };
  return { fetchImpl, calls };
}

/** A response shaped like Graph's: the HTML with its InkNode placeholder, then InkML. */
function multipartResponse(inkml: string): Response {
  const body = [
    `--${BOUNDARY}`,
    'Content-Type: text/html',
    '',
    '<html><body><p>typed</p><!-- InkNode is not supported --></body></html>',
    `--${BOUNDARY}`,
    'Content-Type: application/inkml+xml',
    '',
    inkml,
    `--${BOUNDARY}--`,
    '',
  ].join('\r\n');

  return new Response(body, {
    status: 200,
    headers: { 'content-type': `multipart/mixed; boundary="${BOUNDARY}"` },
  });
}

function fixture(name: string): Promise<string> {
  return readFile(path.join(import.meta.dirname, 'fixtures', name), 'utf8');
}

async function caught(promise: Promise<unknown>): Promise<unknown> {
  try {
    await promise;
  } catch (err) {
    return err;
  }
  assert.fail('expected the call to reject');
}

test('the content URL asks for ink and escapes the page id', () => {
  // Without includeInkML the handwriting is gone from the response entirely, so this
  // query parameter is the whole reason this module exists.
  assert.equal(
    pageContentUrl(PAGE_ID),
    'https://graph.microsoft.com/v1.0/me/onenote/pages/1-abc!123/content?includeInkML=true',
  );
  assert.match(pageContentUrl('a/b'), /pages\/a%2Fb\/content/);
});

test('one fetch yields both the HTML part and the InkML part', async () => {
  const url = pageContentUrl(PAGE_ID);
  const { fetchImpl, calls } = fakeFetch({
    [url]: () => multipartResponse('<ink><trace>0 0, 2540 2540</trace></ink>'),
  });

  const content = await new GraphPageContent(tokens, fetchImpl).fetchRaw(PAGE_ID);

  assert.equal(calls.length, 1, 'the caller must not pay two round trips');
  assert.equal(calls[0]?.authorization, `Bearer ${TOKEN}`);
  assert.equal(content.parts.length, 2);
  assert.match(htmlPart(content)?.body ?? '', /<p>typed<\/p>/);
  assert.equal(content.parts[1]?.body, '<ink><trace>0 0, 2540 2540</trace></ink>');
});

test('a non-multipart response still exposes its body as an ink candidate', async () => {
  const url = pageContentUrl(PAGE_ID);
  const { fetchImpl } = fakeFetch({
    [url]: () =>
      new Response('<html><body><p>typed</p></body></html>', {
        status: 200,
        headers: { 'content-type': 'text/html' },
      }),
  });

  const content = await new GraphPageContent(tokens, fetchImpl).fetchRaw(PAGE_ID);

  assert.deepEqual(content.parts, []);
  assert.equal(htmlPart(content), null);
  assert.deepEqual(
    inkCandidates(content).map((candidate) => candidate.label),
    ['body'],
  );
});

test('the ink in a multipart response renders to a PNG', async () => {
  const url = pageContentUrl(PAGE_ID);
  const inkml = await fixture('xyf-himetric.inkml');
  const { fetchImpl } = fakeFetch({ [url]: () => multipartResponse(inkml) });

  const image = await new GraphPageContent(tokens, fetchImpl).fetchInk(PAGE_ID);

  assert.ok(image !== null);
  assert.equal(image.strokeCount, 2);
  assert.equal(image.width, 1400);
});

test('a page with no ink renders to null, which is normal and not an error', async () => {
  const url = pageContentUrl(PAGE_ID);
  const { fetchImpl } = fakeFetch({
    [url]: () => multipartResponse('<p>this part carries no strokes</p>'),
  });

  assert.equal(await new GraphPageContent(tokens, fetchImpl).fetchInk(PAGE_ID), null);
});

test('ink that appears in both the whole body and a part is drawn once', async () => {
  // The whole body is searched first because a non-multipart response has no parts, and
  // the body contains the InkML part as well. Taking every candidate would double every
  // stroke of every multipart page.
  const content = {
    raw: '<ink><trace>0 0, 2540 2540</trace></ink>',
    contentType: 'multipart/mixed',
    parts: [
      {
        headers: 'Content-Type: application/inkml+xml',
        contentType: 'application/inkml+xml',
        body: '<ink><trace>0 0, 2540 2540</trace></ink>',
      },
    ],
  };

  assert.equal(renderPageInk(content)?.strokeCount, 1);
});

test('a non-2xx content response throws GraphRequestError with the body', async () => {
  const url = pageContentUrl(PAGE_ID);
  const { fetchImpl } = fakeFetch({
    [url]: () => new Response('{"error":{"code":"20266"}}', { status: 400, statusText: 'Bad Request' }),
  });

  const err = await caught(new GraphPageContent(tokens, fetchImpl).fetchRaw(PAGE_ID));

  assert.ok(err instanceof GraphRequestError, `expected GraphRequestError, got ${String(err)}`);
  assert.equal(err.status, 400);
  assert.match(err.body, /20266/);
});

test('one fetch yields the trimmed HTML and the ink together', async () => {
  const url = pageContentUrl(PAGE_ID);
  const inkml = await fixture('xyf-himetric.inkml');
  const { fetchImpl, calls } = fakeFetch({ [url]: () => multipartResponse(inkml) });

  const content = await new GraphPageContent(tokens, fetchImpl).fetchContent(PAGE_ID);

  assert.equal(calls.length, 1, 'the caller must not pay two round trips');
  assert.equal(content.html, '<html><body><p>typed</p></body></html>');
  assert.equal(content.ink?.strokeCount, 2);
});

test('a typed page comes back as HTML with no ink, which is normal', async () => {
  const url = pageContentUrl(PAGE_ID);
  const { fetchImpl } = fakeFetch({
    [url]: () =>
      new Response('<html><body><p id="p:1"><span style="color:red">typed</span></p></body></html>', {
        status: 200,
        headers: { 'content-type': 'text/html' },
      }),
  });

  // A response with no ink is not multipart at all, so the HTML is the whole body.
  const content = await new GraphPageContent(tokens, fetchImpl).fetchContent(PAGE_ID);

  assert.equal(content.html, '<html><body><p id="p:1">typed</p></body></html>');
  assert.equal(content.ink, null);
});

test('the HTML half is found in a multipart body and in a plain one', () => {
  const part = {
    headers: 'Content-Type: text/html',
    contentType: 'text/html',
    body: '<p>from the part</p>',
  };

  assert.equal(
    pageHtml({ raw: 'whole body', contentType: 'multipart/mixed', parts: [part] }),
    '<p>from the part</p>',
  );
  assert.equal(pageHtml({ raw: '<p>whole body</p>', contentType: 'text/html', parts: [] }), '<p>whole body</p>');
  assert.equal(
    pageHtml({
      raw: 'x',
      contentType: 'multipart/mixed',
      parts: [{ headers: '', contentType: 'application/inkml+xml', body: '<ink/>' }],
    }),
    null,
  );
});
