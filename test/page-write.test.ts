// The write client, driven through a fake `fetch` whose routes are keyed by method and
// exact URL. An unrouted request fails the test, so every behavioural assertion here is
// also an assertion about the verb and the URL that were built.
//
// What this cannot check is whether Graph accepts what is sent. The change arrays, the
// `text/html` create, and the 204-with-no-body success are copied from the spike recorded
// in `api-overview.md` under **Writing page content**, which ran against the live service
// on 2026-08-18 — except for the create request, which is the documented `text/html`
// shape rather than the `multipart/form-data` one and was not part of that spike. Nothing
// confirms any of it until an operator runs the server against the real tenant.

import test from 'node:test';
import assert from 'node:assert/strict';

import {
  GraphRequestError,
  GraphResponseError,
  type FetchLike,
  type TokenSource,
} from '../src/graph-structure.ts';
import type { RequestGate } from '../src/graph-throttle.ts';
import {
  GraphPageWrite,
  createPageHtml,
  escapeHtml,
  pageContentPatchUrl,
  sectionPagesUrl,
} from '../src/page-write.ts';

const TOKEN = 'fake-access-token';
const PAGE_ID = '1-abc!123';
const SECTION_ID = '1-sec!456';

const tokens: TokenSource = { getAccessToken: () => Promise.resolve(TOKEN) };

interface Call {
  url: string;
  method: string | undefined;
  authorization: string | undefined;
  contentType: string | undefined;
  body: string | undefined;
}

/** A fetch keyed by `METHOD url`, so an unexpected verb is an unrouted request. */
function fakeFetch(routes: Record<string, () => Response>): {
  fetchImpl: FetchLike;
  calls: Call[];
} {
  const calls: Call[] = [];
  const fetchImpl: FetchLike = (url, init) => {
    calls.push({
      url,
      method: init.method,
      authorization: init.headers['Authorization'],
      contentType: init.headers['Content-Type'],
      body: init.body,
    });
    const route = routes[`${init.method ?? 'GET'} ${url}`];
    if (route === undefined) {
      return Promise.reject(new Error(`no route for ${init.method ?? 'GET'} ${url}`));
    }
    return Promise.resolve(route());
  };
  return { fetchImpl, calls };
}

/** 204 with an empty body, which is what a successful PATCH answers. */
function noContent(): Response {
  return new Response(null, { status: 204 });
}

/** The create response, reduced to the members this module reads. */
function createdPageResponse(id: string, title: string): Response {
  return new Response(
    JSON.stringify({
      id,
      title,
      contentUrl: `https://graph.microsoft.com/v1.0/me/onenote/pages/${id}/content`,
      links: {
        oneNoteClientUrl: { href: `onenote:https://example.invalid/${id}` },
        oneNoteWebUrl: { href: `https://example.invalid/${id}` },
      },
    }),
    { status: 201, headers: { 'content-type': 'application/json' } },
  );
}

async function caught(promise: Promise<unknown>): Promise<unknown> {
  try {
    await promise;
  } catch (err) {
    return err;
  }
  assert.fail('expected the call to reject');
}

function parsedBody(call: Call | undefined): unknown {
  assert.ok(call !== undefined, 'expected a request to have been made');
  assert.ok(call.body !== undefined, 'expected the request to carry a body');
  return JSON.parse(call.body);
}

test('append_to_page PATCHes one body-append change to the page content URL', async () => {
  const url = pageContentPatchUrl(PAGE_ID);
  const { fetchImpl, calls } = fakeFetch({ [`PATCH ${url}`]: noContent });

  await new GraphPageWrite(tokens, fetchImpl).appendToPage(PAGE_ID, '<p>appended</p>');

  assert.equal(calls.length, 1);
  assert.equal(calls[0]?.url, 'https://graph.microsoft.com/v1.0/me/onenote/pages/1-abc!123/content');
  assert.equal(calls[0]?.authorization, `Bearer ${TOKEN}`);
  assert.equal(calls[0]?.contentType, 'application/json');
  assert.deepEqual(parsedBody(calls[0]), [
    { target: 'body', action: 'append', content: '<p>appended</p>' },
  ]);
});

test('update_page_title targets title, never #title, and sends the string verbatim', async () => {
  // `#title` is 400 code 20149 and `append` is 400 code 20141, so both halves of this
  // change object are the only ones the service accepts. The content is not parsed:
  // whatever is sent becomes the title character for character.
  const url = pageContentPatchUrl(PAGE_ID);
  const { fetchImpl, calls } = fakeFetch({ [`PATCH ${url}`]: noContent });

  await new GraphPageWrite(tokens, fetchImpl).updatePageTitle(PAGE_ID, 'A & B "d" 5 < 6');

  assert.deepEqual(parsedBody(calls[0]), [
    { target: 'title', action: 'replace', content: 'A & B "d" 5 < 6' },
  ]);
});

test('the page id is escaped into the PATCH URL', () => {
  assert.match(pageContentPatchUrl('a/b'), /pages\/a%2Fb\/content$/);
  assert.match(sectionPagesUrl('a/b'), /sections\/a%2Fb\/pages$/);
});

test('an empty change array is refused without spending a request', async () => {
  // Graph answers 400 code 20125 to it, and a request that can only fail is quota spent
  // to learn nothing.
  const { fetchImpl, calls } = fakeFetch({});

  const err = await caught(new GraphPageWrite(tokens, fetchImpl).patchPage(PAGE_ID, []));

  assert.ok(err instanceof RangeError);
  assert.equal(calls.length, 0);
});

test('a failed PATCH carries the status, the body, and the verb', async () => {
  // The OData code in the body is the only thing that separates one 400 from another —
  // 20120 is a target that cannot be located — so it must survive to the caller.
  const url = pageContentPatchUrl(PAGE_ID);
  const body = JSON.stringify({
    error: { code: '20120', message: 'The PATCH target could not be located.' },
  });
  const { fetchImpl } = fakeFetch({
    [`PATCH ${url}`]: () => new Response(body, { status: 400, statusText: 'Bad Request' }),
  });

  const err = await caught(
    new GraphPageWrite(tokens, fetchImpl).appendToPage(PAGE_ID, '<p>x</p>'),
  );

  assert.ok(err instanceof GraphRequestError);
  assert.equal(err.status, 400);
  assert.equal(err.method, 'PATCH');
  assert.match(err.body, /20120/);
  assert.match(err.message, /^PATCH /, 'a failed write must not report itself as a failed read');
});

test('create_page POSTs one HTML document and returns the id and the links', async () => {
  const url = sectionPagesUrl(SECTION_ID);
  const { fetchImpl, calls } = fakeFetch({
    [`POST ${url}`]: () => createdPageResponse('1-new!789', 'Meeting notes'),
  });

  const page = await new GraphPageWrite(tokens, fetchImpl).createPage(
    SECTION_ID,
    'Meeting notes',
    '<p>body</p>',
  );

  assert.equal(calls.length, 1);
  assert.equal(calls[0]?.contentType, 'text/html; charset=utf-8');
  assert.equal(page.id, '1-new!789');
  assert.equal(page.title, 'Meeting notes');
  assert.equal(page.webUrl, 'https://example.invalid/1-new!789');
  assert.equal(page.clientUrl, 'onenote:https://example.invalid/1-new!789');
});

test('the submitted document sets a real <title> and leaves the body flowing', () => {
  // The title has to be the <title> element: every by-name lookup and every search
  // matches the title Graph stores, and a heading in the body sets nothing.
  //
  // <body> deliberately carries no data-absolute-enabled. Without it Graph wraps the
  // submission in one `<div data-id="_default">`, which is what makes `body` — and so
  // append_to_page — cover the whole page rather than one outline of it.
  const html = createPageHtml('Meeting notes', '<p>body</p>');

  assert.match(html, /<title>Meeting notes<\/title>/);
  assert.match(html, /<body><p>body<\/p><\/body>/);
  assert.equal(/data-absolute-enabled/.test(html), false);
});

test('a title in a created document is escaped, unlike a title in a PATCH', () => {
  // The two paths are opposite and both are correct. Here the title is an element in a
  // document Graph parses, so `<` has to be escaped or it opens a tag; in the PATCH the
  // content is stored as characters and escaping it would put `&amp;` in the title.
  const html = createPageHtml('A & B <c> "d"', '<p>body</p>');

  assert.match(html, /<title>A &amp; B &lt;c&gt; &quot;d&quot;<\/title>/);
  assert.equal(escapeHtml('<script>'), '&lt;script&gt;');
});

test('a create whose response is not JSON is a GraphResponseError', async () => {
  const url = sectionPagesUrl(SECTION_ID);
  const { fetchImpl } = fakeFetch({
    [`POST ${url}`]: () => new Response('<html>gateway</html>', { status: 201 }),
  });

  const err = await caught(
    new GraphPageWrite(tokens, fetchImpl).createPage(SECTION_ID, 'T', '<p>b</p>'),
  );

  assert.ok(err instanceof GraphResponseError);
});

test('a create whose response carries no id is a GraphResponseError', async () => {
  // The page may well exist, but nothing can be said about it, and a caller told
  // "created" with no id has nothing to read back or append to.
  const url = sectionPagesUrl(SECTION_ID);
  const { fetchImpl } = fakeFetch({
    [`POST ${url}`]: () =>
      new Response(JSON.stringify({ title: 'T' }), {
        status: 201,
        headers: { 'content-type': 'application/json' },
      }),
  });

  const err = await caught(
    new GraphPageWrite(tokens, fetchImpl).createPage(SECTION_ID, 'T', '<p>b</p>'),
  );

  assert.ok(err instanceof GraphResponseError);
  assert.match((err as Error).message, /no page id/);
});

test('a create response with no links yields nulls rather than failing', async () => {
  // The page exists. Failing the call over a missing convenience link would report a
  // failure the caller cannot act on and cannot undo.
  const url = sectionPagesUrl(SECTION_ID);
  const { fetchImpl } = fakeFetch({
    [`POST ${url}`]: () =>
      new Response(JSON.stringify({ id: '1-new!789', title: 'T' }), {
        status: 201,
        headers: { 'content-type': 'application/json' },
      }),
  });

  const page = await new GraphPageWrite(tokens, fetchImpl).createPage(
    SECTION_ID,
    'T',
    '<p>b</p>',
  );

  assert.equal(page.id, '1-new!789');
  assert.equal(page.webUrl, null);
  assert.equal(page.clientUrl, null);
});

test('every write goes through the gate', async () => {
  // Writes count against the same per-user limits as reads, so a write that paced itself
  // separately would push the reads over. The gate is what holds both to one budget.
  let gated = 0;
  const gate: RequestGate = {
    run: (operation) => {
      gated += 1;
      return operation();
    },
  };
  const patchUrl = pageContentPatchUrl(PAGE_ID);
  const postUrl = sectionPagesUrl(SECTION_ID);
  const { fetchImpl } = fakeFetch({
    [`PATCH ${patchUrl}`]: noContent,
    [`POST ${postUrl}`]: () => createdPageResponse('1-new!789', 'T'),
  });

  const write = new GraphPageWrite(tokens, fetchImpl, gate);
  await write.appendToPage(PAGE_ID, '<p>x</p>');
  await write.updatePageTitle(PAGE_ID, 'T');
  await write.createPage(SECTION_ID, 'T', '<p>b</p>');

  assert.equal(gated, 3);
});
