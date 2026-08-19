// append_to_page, create_page and update_page_title, driven through their own `handle`
// with a fake write client. The URLs and the request bodies are asserted in
// test/page-write.test.ts; what matters here is what reaches the client and what never
// does.
//
// The refusals are the point of this file. A rejected argument must cost no Graph
// request: a write that is going to be wrong is cheaper to refuse than to send, and on a
// create it would otherwise leave a page behind that the caller then has to find and
// delete. Every refusal test asserts the client was not called.
//
// `create_page_by_name` is driven through a fake `LookupStructure` as well, and that fake
// counts its calls: the resolver's own rules are asserted in test/name-lookup.test.ts, so
// what this file checks is that the section id the tool creates in is the one the lookup
// returned, and that a refused argument costs neither the lookup nor the write.

import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';

import type { CallToolResult } from '@modelcontextprotocol/sdk/types.js';

import type { ExpandedNotebook, PageSummary } from '../src/graph-structure.ts';
import { GraphRequestError } from '../src/graph-structure.ts';
import { NameLookupError } from '../src/name-lookup.ts';
import { ToolInputError, indexTools, type ToolDefinition } from '../src/mcp-tools.ts';
import type { RawPageContent } from '../src/page-content.ts';
import { CLEARANCE_ID_PREFIX } from '../src/page-layout.ts';
import type { CreatedPage } from '../src/page-write.ts';
import {
  createWriteTools,
  type PageLayoutReader,
  type PageWriteClient,
  type WriteLookupStructure,
} from '../src/write-tools.ts';

const PAGE_ID = '1-abc!123';
const SECTION_ID = '1-sec!456';

/** An outline in the shape Graph emits one, with no handwriting anywhere near it. */
const TYPED_PAGE =
  '<html><head><title>t</title></head><body data-absolute-enabled="true">' +
  '<div id="div:{aaa}{32}" data-id="_default" style="position:absolute;left:48px;top:120px;width:624px">' +
  '<p>Typed text.</p></div></body></html>';

/** The committed InkML whose strokes sit inside that outline's column, ending at 466px. */
const INK_BELOW_TEXT = readFileSync(
  new URL('./fixtures/ink-below-text.inkml', import.meta.url),
  'utf8',
);

const CREATED: CreatedPage = {
  id: '1-new!789',
  title: 'Meeting notes',
  webUrl: 'https://example.invalid/1-new!789',
  clientUrl: 'onenote:https://example.invalid/1-new!789',
};

interface Fake extends PageWriteClient {
  readonly calls: string[][];
}

function fakeWrite(created: CreatedPage = CREATED, fail?: () => never): Fake {
  const calls: string[][] = [];
  return {
    calls,
    appendToPage: (pageId, html) => {
      calls.push(['appendToPage', pageId, html]);
      if (fail !== undefined) fail();
      return Promise.resolve();
    },
    updatePageTitle: (pageId, title) => {
      calls.push(['updatePageTitle', pageId, title]);
      if (fail !== undefined) fail();
      return Promise.resolve();
    },
    createPage: (sectionId, title, bodyHtml) => {
      calls.push(['createPage', sectionId, title, bodyHtml]);
      if (fail !== undefined) fail();
      return Promise.resolve(created);
    },
  };
}

/** A page in the shape Graph returns one, with the ink part optional. */
function pageContent(html: string, inkml = ''): RawPageContent {
  const parts = [
    { headers: 'Content-Type: text/html', contentType: 'text/html', body: html },
    ...(inkml === ''
      ? []
      : [{ headers: 'Content-Type: application/inkml+xml', contentType: 'application/inkml+xml', body: inkml }]),
  ];
  return { raw: [html, inkml].join('\n'), contentType: 'multipart/mixed; boundary=b', parts };
}

interface FakeReader extends PageLayoutReader {
  readonly reads: string[];
}

function fakeLayout(content: RawPageContent = pageContent(TYPED_PAGE)): FakeReader {
  const reads: string[] = [];
  return {
    reads,
    fetchRaw: (pageId) => {
      reads.push(pageId);
      return Promise.resolve(content);
    },
  };
}

/**
 * A notebook holding one section directly and one inside a section group whose name
 * carries an ordering prefix. The prefix is what `create_page_by_name` has to see
 * through when a caller names the month rather than the number.
 */
const EXPANDED: ExpandedNotebook[] = [
  {
    id: 'nb-2026',
    displayName: 'Bullet Journal - 2026',
    sections: [{ id: 'sec-inbox', displayName: 'Inbox' }],
    sectionGroups: [
      {
        id: 'grp-feb',
        displayName: '062 - February',
        sections: [{ id: 'sec-log', displayName: 'Monthly Log' }],
      },
    ],
  },
];

/** The pages `findPagesByTitle` answers with, keyed the way Graph filters them: by title. */
const PAGES: Record<string, PageSummary[]> = {
  'monthly log': [
    { id: '1-page!log', title: 'Monthly Log', lastModifiedDateTime: '2026-08-19T10:00:00Z' },
  ],
  standup: [
    { id: '1-page!a', title: 'Standup', lastModifiedDateTime: '2026-08-19T10:00:00Z' },
    { id: '1-page!b', title: 'standup', lastModifiedDateTime: '2026-08-18T10:00:00Z' },
  ],
};

interface FakeLookup extends WriteLookupStructure {
  readonly lookups: { tree: number; byName: string[]; titles: { sectionId: string; title: string }[] };
}

function fakeLookup(tree: ExpandedNotebook[] = EXPANDED): FakeLookup {
  const lookups = {
    tree: 0,
    byName: [] as string[],
    titles: [] as { sectionId: string; title: string }[],
  };
  return {
    lookups,
    getExpandedTree: () => {
      lookups.tree += 1;
      return Promise.resolve(tree);
    },
    findSectionsByName: (displayName: string) => {
      lookups.byName.push(displayName);
      return Promise.resolve([]);
    },
    // Graph does this comparison in full and case-insensitively, so the fake does too.
    findPagesByTitle: (sectionId: string, title: string) => {
      lookups.titles.push({ sectionId, title });
      return Promise.resolve(PAGES[title.trim().toLowerCase()] ?? []);
    },
  };
}

function tool(
  name: string,
  client: PageWriteClient,
  layout: PageLayoutReader = fakeLayout(),
  lookup: WriteLookupStructure = fakeLookup(),
): ToolDefinition {
  const found = indexTools(createWriteTools(client, layout, lookup)).get(name);
  assert.ok(found !== undefined, `${name} must be registered`);
  return found;
}

function payload(result: CallToolResult): Record<string, unknown> {
  const block = result.content[0];
  assert.ok(block !== undefined && block.type === 'text');
  return JSON.parse(block.text) as Record<string, unknown>;
}

async function caught(promise: Promise<unknown>): Promise<unknown> {
  try {
    await promise;
  } catch (err) {
    return err;
  }
  assert.fail('expected the call to reject');
}

test('the writing tools are registered and none claims to be read-only', () => {
  const tools = createWriteTools(fakeWrite(), fakeLayout(), fakeLookup());

  assert.deepEqual(
    tools.map((t) => t.name),
    [
      'append_to_page',
      'append_to_page_by_name',
      'create_page',
      'create_page_by_name',
      'update_page_title',
    ],
  );
  for (const t of tools) {
    assert.equal(t.annotations?.readOnlyHint, false, `${t.name} writes`);
  }
});

test('append_to_page passes the fragment through and says where it landed', async () => {
  // The note is not decoration. `body` is the first outline on a page a person wrote in
  // the OneNote client, so a caller told only "appended" would believe it reached the
  // bottom of the page.
  const client = fakeWrite();

  const result = await tool('append_to_page', client).handle({
    pageId: PAGE_ID,
    htmlFragment: '<p>appended</p>',
  });

  assert.deepEqual(client.calls, [['appendToPage', PAGE_ID, '<p>appended</p>']]);
  const body = payload(result);
  assert.equal(body['pageId'], PAGE_ID);
  assert.equal(body['appended'], true);
  assert.match(String(body['note']), /first outline/);
});

test('an append to a page with ink below the text is pushed below the strokes', async () => {
  // The whole point of reading before writing. OneNote fixes ink in place and no write
  // can move it, so text appended into an outline the strokes overlap renders on top of
  // the handwriting. The fixture's ink ends at 466px, the outline starts at 120px.
  const client = fakeWrite();
  const layout = fakeLayout(pageContent(TYPED_PAGE, INK_BELOW_TEXT));

  const result = await tool('append_to_page', client, layout).handle({
    pageId: PAGE_ID,
    htmlFragment: '<p>appended</p>',
  });

  assert.deepEqual(layout.reads, [PAGE_ID]);
  const sent = client.calls[0]?.[2] ?? '';
  assert.equal((sent.match(/<br \/>/g) ?? []).length, 18);
  assert.ok(
    sent.endsWith(`<div data-id="${CLEARANCE_ID_PREFIX}478"><p>appended</p></div>`),
    "the caller's fragment goes last, unchanged, inside the marker",
  );

  const clearance = payload(result)['inkClearance'] as Record<string, unknown>;
  assert.deepEqual(clearance, {
    blankLines: 18,
    inkBottomPx: 466,
    // The outline holds one paragraph, so the text is estimated to end 16px below its
    // top and the padding is measured from there rather than from the top itself.
    measuredFromPx: 136,
    contentStartsAtPx: 478,
    clearsTheInk: true,
  });
});

test('a page with no ink is appended to exactly as asked', async () => {
  const client = fakeWrite();
  const layout = fakeLayout();

  const result = await tool('append_to_page', client, layout).handle({
    pageId: PAGE_ID,
    htmlFragment: '<p>appended</p>',
  });

  assert.deepEqual(client.calls, [['appendToPage', PAGE_ID, '<p>appended</p>']]);
  assert.equal(payload(result)['inkClearance'], null);
});

test('a second append does not stack another block of blank lines', async () => {
  // The marker left by the first append records the ink it cleared. Without reading it
  // back, a page written to three times would carry three stacks of padding.
  const cleared =
    TYPED_PAGE.replace(
      '<p>Typed text.</p>',
      `<p>Typed text.</p><p><br /></p><div data-id="${CLEARANCE_ID_PREFIX}490"><p>already below the ink</p></div>`,
    );
  const client = fakeWrite();

  await tool('append_to_page', client, fakeLayout(pageContent(cleared, INK_BELOW_TEXT))).handle({
    pageId: PAGE_ID,
    htmlFragment: '<p>second</p>',
  });

  assert.deepEqual(client.calls, [['appendToPage', PAGE_ID, '<p>second</p>']]);
});

test('a refused fragment costs no read and no write', async () => {
  const client = fakeWrite();
  const layout = fakeLayout();

  await caught(
    tool('append_to_page', client, layout).handle({
      pageId: PAGE_ID,
      htmlFragment: '<body><p>x</p></body>',
    }),
  );

  assert.deepEqual(layout.reads, [], 'the argument check comes before the read');
  assert.deepEqual(client.calls, []);
});

test('an unclosed tag is appended rather than refused', async () => {
  // Measured on the live service: `<p>unclosed` returned 204 and OneNote closed the tag.
  // A strict check here would refuse content that works.
  const client = fakeWrite();

  await tool('append_to_page', client).handle({
    pageId: PAGE_ID,
    htmlFragment: '<p>unclosed',
  });

  assert.equal(client.calls.length, 1);
});

test('a fragment carrying a document element is refused, and costs no request', async () => {
  for (const fragment of [
    '<html><body><p>x</p></body></html>',
    '<body><p>x</p></body>',
    '<title>not the page title</title>',
    '<script>alert(1)</script>',
    '<p>x</p><style>p { color: red }</style>',
  ]) {
    const client = fakeWrite();
    const err = await caught(
      tool('append_to_page', client).handle({ pageId: PAGE_ID, htmlFragment: fragment }),
    );

    assert.ok(err instanceof ToolInputError, `${fragment} must be refused`);
    assert.equal(err.argument, 'htmlFragment');
    assert.equal(client.calls.length, 0, 'a refused fragment must not reach Graph');
  }
});

test('create_page sends the title and the body, and answers with the new id', async () => {
  const client = fakeWrite();

  const result = await tool('create_page', client).handle({
    sectionId: SECTION_ID,
    title: 'Meeting notes',
    htmlFragment: '<p>body</p>',
  });

  assert.deepEqual(client.calls, [['createPage', SECTION_ID, 'Meeting notes', '<p>body</p>']]);
  const body = payload(result);
  assert.equal(body['pageId'], '1-new!789');
  assert.equal(body['title'], 'Meeting notes');
  assert.equal(body['webUrl'], 'https://example.invalid/1-new!789');
});

test('create_page reports the requested title when Graph echoes an empty one', async () => {
  // Page metadata lags: two pages created during the #17 spike answered `""` for their
  // title within seconds of being made. Reporting that empty string would tell the caller
  // its title was lost when it was not.
  const client = fakeWrite({ ...CREATED, title: '' });

  const result = await tool('create_page', client).handle({
    sectionId: SECTION_ID,
    title: 'Meeting notes',
    htmlFragment: '<p>body</p>',
  });

  assert.equal(payload(result)['title'], 'Meeting notes');
});

test('append_to_page_by_name resolves the path and appends to the page it found', async () => {
  const client = fakeWrite();
  const lookup = fakeLookup();

  const result = await tool('append_to_page_by_name', client, fakeLayout(), lookup).handle({
    notebookName: 'Bullet Journal - 2026',
    sectionGroupName: 'February',
    sectionName: 'Monthly Log',
    pageTitle: 'monthly log',
    htmlFragment: '<p>appended</p>',
  });

  assert.deepEqual(client.calls, [['appendToPage', '1-page!log', '<p>appended</p>']]);
  assert.equal(lookup.lookups.tree, 1);
  assert.deepEqual(lookup.lookups.titles, [{ sectionId: 'sec-log', title: 'monthly log' }]);

  const body = payload(result);
  assert.equal(body['pageId'], '1-page!log');
  // The title Graph holds, not the one the caller typed: the caller's differed in case.
  assert.deepEqual(body['page'], { id: '1-page!log', title: 'Monthly Log' });
  assert.deepEqual(body['section'], { id: 'sec-log', displayName: 'Monthly Log' });
  assert.equal(body['appended'], true);
  assert.match(String(body['note']), /first outline/);
});

test('an append by name to a page with ink is padded the same way as one by id', async () => {
  // The clearance is shared code, and this is what says so. A second copy of the append
  // that skipped the read would write over the handwriting and report success.
  const client = fakeWrite();
  const layout = fakeLayout(pageContent(TYPED_PAGE, INK_BELOW_TEXT));

  const result = await tool('append_to_page_by_name', client, layout).handle({
    notebookName: 'Bullet Journal - 2026',
    sectionGroupName: 'February',
    sectionName: 'Monthly Log',
    pageTitle: 'Monthly Log',
    htmlFragment: '<p>appended</p>',
  });

  assert.deepEqual(layout.reads, ['1-page!log']);
  const sent = client.calls[0]?.[2] ?? '';
  assert.equal((sent.match(/<br \/>/g) ?? []).length, 18);
  assert.ok(sent.endsWith(`<div data-id="${CLEARANCE_ID_PREFIX}478"><p>appended</p></div>`));
  assert.equal(
    (payload(result)['inkClearance'] as Record<string, unknown>)['contentStartsAtPx'],
    478,
  );
});

test('a page title matching nothing, or more than one page, writes nothing', async () => {
  // A page title is matched in full and nothing else, so both of these are the caller
  // being wrong about which page it meant. Appending to a guess would put content on
  // someone's page with nothing to say it happened.
  for (const [pageTitle, kind] of [
    ['No Such Page', 'not-found'],
    ['Standup', 'ambiguous'],
  ] as const) {
    const client = fakeWrite();
    const layout = fakeLayout();

    const err = await caught(
      tool('append_to_page_by_name', client, layout).handle({
        notebookName: 'Bullet Journal - 2026',
        sectionGroupName: 'February',
        sectionName: 'Monthly Log',
        pageTitle,
        htmlFragment: '<p>appended</p>',
      }),
    );

    assert.ok(err instanceof NameLookupError, `${pageTitle} must not be resolved to a page`);
    assert.equal(err.kind, kind);
    assert.equal(err.argument, 'pageTitle');
    // The container ladder does not apply to a page title, and the message must not tell
    // the caller to drop a leading number from one.
    assert.doesNotMatch(err.message, /leading number removed/);
    assert.match(err.message, /list_pages_by_name/);
    assert.equal(client.calls.length, 0, 'nothing may be written');
    assert.deepEqual(layout.reads, [], 'and no page may be read');
  }
});

test('append_to_page_by_name refuses a bad fragment before it spends the lookup', async () => {
  const client = fakeWrite();
  const lookup = fakeLookup();

  const err = await caught(
    tool('append_to_page_by_name', client, fakeLayout(), lookup).handle({
      notebookName: 'Bullet Journal - 2026',
      sectionGroupName: 'February',
      sectionName: 'Monthly Log',
      pageTitle: 'Monthly Log',
      htmlFragment: '<script>x</script>',
    }),
  );

  assert.ok(err instanceof ToolInputError);
  assert.equal(err.argument, 'htmlFragment');
  assert.equal(lookup.lookups.tree, 0, 'a refused fragment must not cost a Graph request');
  assert.deepEqual(lookup.lookups.titles, []);
  assert.equal(client.calls.length, 0);
});

test('create_page_by_name resolves the names and creates the page in that section', async () => {
  // The whole point of the tool: notebook, section group and section names in, one page
  // created, no list_notebooks -> list_sections walk in between. 'February' has to reach
  // '062 - February', because a caller knows the month and not the number.
  const client = fakeWrite();
  const lookup = fakeLookup();

  const result = await tool('create_page_by_name', client, fakeLayout(), lookup).handle({
    notebookName: 'bullet journal - 2026',
    sectionGroupName: 'February',
    sectionName: 'Monthly Log',
    title: 'Meeting notes',
    htmlFragment: '<p>body</p>',
  });

  assert.deepEqual(client.calls, [['createPage', 'sec-log', 'Meeting notes', '<p>body</p>']]);
  assert.equal(lookup.lookups.tree, 1, 'the common path is one expanded-tree request');
  assert.deepEqual(lookup.lookups.byName, [], 'and no account-wide fallback');

  const body = payload(result);
  assert.equal(body['pageId'], '1-new!789');
  assert.deepEqual(body['section'], { id: 'sec-log', displayName: 'Monthly Log' });
  assert.deepEqual(body['sectionGroup'], { id: 'grp-feb', displayName: '062 - February' });
  // A caller that asked for 'February' can see what it actually got, and why.
  assert.deepEqual(body['matchedBy'], {
    notebook: 'exact',
    sectionGroup: 'without-prefix',
    section: 'exact',
  });
  assert.equal(body['deepSearchUsed'], false);
});

test('create_page_by_name reaches a section sitting directly in the notebook', async () => {
  const client = fakeWrite();

  await tool('create_page_by_name', client).handle({
    notebookName: 'Bullet Journal - 2026',
    sectionName: 'Inbox',
    title: 'Meeting notes',
    htmlFragment: '<p>body</p>',
  });

  assert.deepEqual(client.calls, [['createPage', 'sec-inbox', 'Meeting notes', '<p>body</p>']]);
});

test('a name that matches nothing creates no page', async () => {
  // The failure has to happen before the write. A tool that guessed a section would put
  // someone's meeting notes in the wrong notebook, and nothing would say so.
  const client = fakeWrite();

  const err = await caught(
    tool('create_page_by_name', client).handle({
      notebookName: 'Bullet Journal - 2026',
      sectionName: 'No Such Section',
      title: 'Meeting notes',
      htmlFragment: '<p>body</p>',
    }),
  );

  assert.ok(err instanceof NameLookupError);
  assert.equal(err.argument, 'sectionName');
  assert.equal(client.calls.length, 0, 'a page must not be created under a guessed name');
});

test('create_page_by_name refuses a bad fragment before it spends the lookup', async () => {
  const client = fakeWrite();
  const lookup = fakeLookup();

  const err = await caught(
    tool('create_page_by_name', client, fakeLayout(), lookup).handle({
      notebookName: 'Bullet Journal - 2026',
      sectionName: 'Inbox',
      title: 'Meeting notes',
      htmlFragment: '<html><body><p>x</p></body></html>',
    }),
  );

  assert.ok(err instanceof ToolInputError);
  assert.equal(err.argument, 'htmlFragment');
  assert.equal(lookup.lookups.tree, 0, 'a refused fragment must not cost a Graph request');
  assert.equal(client.calls.length, 0);
});

test('a title holding a tag is refused on both tools, and costs no request', async () => {
  // A title is stored verbatim, so `<p>x</p>` produces a page literally called `<p>x</p>`.
  // Graph answers 204 to that, which means nothing but this check can catch it.
  for (const [name, args] of [
    ['create_page', { sectionId: SECTION_ID, htmlFragment: '<p>b</p>' }],
    ['update_page_title', { pageId: PAGE_ID }],
  ] as const) {
    const field = name === 'create_page' ? 'title' : 'newTitle';
    const client = fakeWrite();

    const err = await caught(
      tool(name, client).handle({ ...args, [field]: '<p>marked up</p>' }),
    );

    assert.ok(err instanceof ToolInputError, `${name} must refuse a title holding a tag`);
    assert.equal(err.argument, field);
    assert.equal(client.calls.length, 0);
  }
});

test('a title holding a bare < is allowed; this tool is the only way to set one', async () => {
  const client = fakeWrite();

  await tool('update_page_title', client).handle({ pageId: PAGE_ID, newTitle: 'if x <y then' });

  assert.deepEqual(client.calls, [['updatePageTitle', PAGE_ID, 'if x <y then']]);
});

test('update_page_title sends the new title verbatim and reports it back', async () => {
  const client = fakeWrite();

  const result = await tool('update_page_title', client).handle({
    pageId: PAGE_ID,
    newTitle: 'A & B "d" 5 < 6',
  });

  assert.deepEqual(client.calls, [['updatePageTitle', PAGE_ID, 'A & B "d" 5 < 6']]);
  assert.equal(payload(result)['title'], 'A & B "d" 5 < 6');
});

test('a missing argument is a ToolInputError naming it', async () => {
  const client = fakeWrite();

  const err = await caught(tool('append_to_page', client).handle({ pageId: PAGE_ID }));

  assert.ok(err instanceof ToolInputError);
  assert.equal(err.argument, 'htmlFragment');
  assert.equal(client.calls.length, 0);
});

test('a Graph failure propagates for mcp-tools to turn into an isError result', async () => {
  // Nothing here catches it. `toolErrorResult` reduces a GraphRequestError to its status
  // and OData code in one place, and duplicating that here would produce two wordings for
  // the same failure.
  const client = fakeWrite(CREATED, () => {
    throw new GraphRequestError(
      'https://graph.microsoft.com/v1.0/me/onenote/pages/1-abc!123/content',
      404,
      'Not Found',
      '{"error":{"code":"20112","message":"invalid entity id"}}',
      undefined,
      'PATCH',
    );
  });

  const err = await caught(
    tool('append_to_page', client).handle({ pageId: PAGE_ID, htmlFragment: '<p>x</p>' }),
  );

  assert.ok(err instanceof GraphRequestError);
  assert.equal(err.status, 404);
});
