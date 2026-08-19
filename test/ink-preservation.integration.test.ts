// The one test that has to run against the real account: does a write strip handwriting
// off a page that has some?
//
// Nothing in this repository can answer that with a fixture. Graph decides what a PATCH
// does to the ink it is not told about, and the failure mode is silent — a page comes
// back with its typed text intact and its strokes gone, with no error anywhere. So this
// file drives the shipped `append_to_page` and `update_page_title` tools against a live
// page, renders its ink before and after, and compares the stroke count and the PNG
// bytes. Issue #19.
//
// It is skipped unless the environment names a page, so `npm test` and CI run everything
// else without a credential. To run it:
//
//   1. In the OneNote client, make a throwaway page in a scratch section. Put fake typed
//      text on it and write on it by hand from a tablet — the handwriting cannot be
//      created through Graph, which is why this step needs a person. Do not use a page
//      that holds anything you would mind losing: every assertion here is about the ink
//      surviving, and the run itself is what finds out.
//   2. Export the same variables the server reads — ONENOTE_CLIENT_ID,
//      ONENOTE_AUTHORITY, GOOGLE_CLOUD_PROJECT, FIRESTORE_CACHE_DOC — and have a seeded
//      token cache (`npm run bootstrap`).
//   3. Name the page, either directly or by the names you gave it:
//        ONENOTE_INK_TEST_PAGE_ID=1-abc!123
//      or
//        ONENOTE_INK_TEST_NOTEBOOK='Scratch' ONENOTE_INK_TEST_SECTION='Ink tests' \
//        ONENOTE_INK_TEST_PAGE_TITLE='Ink write check'
//      with ONENOTE_INK_TEST_SECTION_GROUP as well when the section sits in one.
//   4. `node --test test/ink-preservation.integration.test.ts`
//
// The run leaves the page changed: `append_to_page` adds a marker paragraph and never
// removes it, and the rename is put back only when the original title could be read. Do
// not paste this test's output into an issue or a log — an assertion failure prints
// stroke counts and hashes, which are harmless, but the surrounding run may name the
// notebook and section you set above.

import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { before, describe, it } from 'node:test';
import { setTimeout as delay } from 'node:timers/promises';

import { loadConfig } from '../src/config.ts';
import { createGraphAuth, type GraphAuth } from '../src/graph-auth.ts';
import { GRAPH_ROOT, createGraphStructure, graphGet } from '../src/graph-structure.ts';
import type { GraphStructure } from '../src/graph-structure.ts';
import { DEFAULT_RENDER_WIDTH } from '../src/ink.ts';
import { indexTools, type ToolDefinition } from '../src/mcp-tools.ts';
import { resolveSection } from '../src/name-lookup.ts';
import { createGraphPageContent, pageHtml, renderPageInk } from '../src/page-content.ts';
import { createGraphPageWrite, type GraphPageWrite } from '../src/page-write.ts';
import { createWriteTools } from '../src/write-tools.ts';

const PAGE_ID_VAR = 'ONENOTE_INK_TEST_PAGE_ID';
const NOTEBOOK_VAR = 'ONENOTE_INK_TEST_NOTEBOOK';
const SECTION_GROUP_VAR = 'ONENOTE_INK_TEST_SECTION_GROUP';
const SECTION_VAR = 'ONENOTE_INK_TEST_SECTION';
const TITLE_VAR = 'ONENOTE_INK_TEST_PAGE_TITLE';

/**
 * How long a write is given to become visible.
 *
 * Page content is readable back quickly, but page *metadata* — which is where a title
 * lives — lags by seconds: the spike in issue #17 saw a create whose response carried a
 * title that a `$select=title` read still answered `""` to. The polling is here so a lag
 * reads as a lag rather than as a failed write.
 */
const POLL_ATTEMPTS = 8;
const POLL_INTERVAL_MS = 2_000;

/** One reading of the page: its ink, and enough of its text to prove a write landed. */
interface PageState {
  readonly strokeCount: number;
  /** null when the page carried no ink at all — which this test treats as a failure. */
  readonly pngSha256: string | null;
  readonly pngBytes: number;
  readonly html: string | null;
  /** The `<title>` of the content document, when Graph sent one. */
  readonly contentTitle: string | null;
}

function env(name: string): string | null {
  const value = process.env[name];
  return value === undefined || value.trim() === '' ? null : value.trim();
}

/** Either a page id, or a notebook, a section and a page title, or the suite is skipped. */
function skipReason(): string | false {
  if (env(PAGE_ID_VAR) !== null) return false;
  if (env(NOTEBOOK_VAR) !== null && env(SECTION_VAR) !== null && env(TITLE_VAR) !== null) {
    return false;
  }
  return (
    `live-account test: set ${PAGE_ID_VAR}, or ${NOTEBOOK_VAR} + ${SECTION_VAR} + ` +
    `${TITLE_VAR} (and ${SECTION_GROUP_VAR} when the section is in one), against a ` +
    'throwaway page carrying handwriting. See the header of this file.'
  );
}

describe('writes preserve existing ink (live account)', { skip: skipReason() }, () => {
  let auth: GraphAuth;
  let structure: GraphStructure;
  let write: GraphPageWrite;
  let tools: Map<string, ToolDefinition>;
  let pageId: string;
  let baseline: PageState;

  /** One reading of the page, from the same single fetch the server's read tool uses. */
  async function readPage(): Promise<PageState> {
    const content = createGraphPageContent(auth);
    const raw = await content.fetchRaw(pageId);
    const ink = renderPageInk(raw, DEFAULT_RENDER_WIDTH);
    const html = pageHtml(raw);

    return {
      strokeCount: ink === null ? 0 : ink.strokeCount,
      pngSha256: ink === null ? null : createHash('sha256').update(ink.png).digest('hex'),
      pngBytes: ink === null ? 0 : ink.png.byteLength,
      html,
      contentTitle: html === null ? null : titleOf(html),
    };
  }

  /** The title Graph holds as page metadata, which is what a listing and a search show. */
  async function metadataTitle(): Promise<string | null> {
    const url =
      `${GRAPH_ROOT}/me/onenote/pages/${encodeURIComponent(pageId)}?$select=id,title`;
    const body = await graphGet(url, await auth.getAccessToken(), globalThis.fetch);
    const title = body['title'];
    return typeof title === 'string' && title !== '' ? title : null;
  }

  /** Re-read until `ready`, or until the attempts run out; the caller asserts. */
  async function pollUntil<T>(read: () => Promise<T>, ready: (value: T) => boolean): Promise<T> {
    let value = await read();
    for (let attempt = 1; attempt < POLL_ATTEMPTS && !ready(value); attempt += 1) {
      await delay(POLL_INTERVAL_MS);
      value = await read();
    }
    return value;
  }

  /** Call a shipped tool the way the MCP layer does, and fail on an `isError` result. */
  async function callTool(name: string, args: Record<string, unknown>): Promise<void> {
    const tool = tools.get(name);
    assert.ok(tool !== undefined, `internal: no tool named ${name}`);
    const result = await tool.handle(args);
    assert.notEqual(result.isError, true, `${name} failed: ${resultText(result)}`);
  }

  /** Both invariants in one place, so each failure says what a wrong answer means. */
  function assertInkUnchanged(after: PageState, operation: string): void {
    assert.equal(
      after.strokeCount,
      baseline.strokeCount,
      `${operation} changed the stroke count: ${baseline.strokeCount} before, ` +
        `${after.strokeCount} after. The write tools do not ship.`,
    );
    assert.equal(
      after.pngSha256,
      baseline.pngSha256,
      `${operation} changed the rendered ink: ${baseline.pngBytes} bytes before, ` +
        `${after.pngBytes} after, and the hashes differ. The stroke count alone can ` +
        'match while the strokes themselves moved. The write tools do not ship.',
    );
  }

  before(async () => {
    const config = loadConfig(['graph', 'firestore']);
    if (config.graph === undefined || config.firestore === undefined) {
      throw new Error('internal: loadConfig returned neither graph nor firestore');
    }

    auth = createGraphAuth(config.graph, config.firestore);
    structure = createGraphStructure(auth);
    write = createGraphPageWrite(auth);
    tools = indexTools(createWriteTools(write, createGraphPageContent(auth)));
    pageId = await resolvePageId(structure);

    baseline = await readPage();
    assert.ok(
      baseline.pngSha256 !== null,
      'the test page carries no handwriting, so this test would pass without testing ' +
        'anything. Write on the page by hand from a tablet and run this again.',
    );
  });

  it('append_to_page leaves the handwriting untouched', async () => {
    const marker = `ink-check-append-${Date.now().toString(36)}`;

    await callTool('append_to_page', { pageId, htmlFragment: `<p>${marker}</p>` });

    const after = await pollUntil(readPage, (state) => (state.html ?? '').includes(marker));

    // Without this the ink assertions below are vacuous: an append that silently did
    // nothing would leave the strokes identical and the test green.
    assert.ok(
      (after.html ?? '').includes(marker),
      `the appended paragraph never appeared in the page HTML after ` +
        `${POLL_ATTEMPTS * POLL_INTERVAL_MS} ms, so nothing here tested a write.`,
    );

    assertInkUnchanged(after, 'append_to_page');
  });

  it('update_page_title leaves the handwriting untouched', async (t) => {
    const marker = `ink-check-title-${Date.now().toString(36)}`;
    const original = (await metadataTitle()) ?? baseline.contentTitle;
    const newTitle = `ink write check ${marker}`;

    try {
      await callTool('update_page_title', { pageId, newTitle });

      const seen = await pollUntil(
        async () => (await metadataTitle()) ?? (await readPage()).contentTitle,
        (title) => title === newTitle,
      );

      assert.equal(
        seen,
        newTitle,
        `the rename never became visible after ${POLL_ATTEMPTS * POLL_INTERVAL_MS} ms, ` +
          'so nothing here tested a write.',
      );

      assertInkUnchanged(await readPage(), 'update_page_title');
    } finally {
      if (original !== null && original !== '') {
        await write.updatePageTitle(pageId, original);
      } else {
        t.diagnostic(
          'the original page title could not be read, so the page keeps the test title.',
        );
      }
    }
  });
});

/** The page named by the environment: an id outright, or the names resolved to one. */
async function resolvePageId(structure: GraphStructure): Promise<string> {
  const explicit = env(PAGE_ID_VAR);
  if (explicit !== null) return explicit;

  const notebookName = env(NOTEBOOK_VAR);
  const sectionName = env(SECTION_VAR);
  const title = env(TITLE_VAR);
  if (notebookName === null || sectionName === null || title === null) {
    throw new Error('internal: the suite should have been skipped');
  }

  const resolved = await resolveSection(structure, {
    notebookName,
    sectionGroupName: env(SECTION_GROUP_VAR) ?? undefined,
    sectionName,
  });

  const pages = await structure.findPagesByTitle(resolved.section.id, title);
  const page = pages[0];
  if (page === undefined || pages.length !== 1) {
    throw new Error(
      `${TITLE_VAR} matched ${pages.length} pages in the named section; it has to match ` +
        `exactly one, or set ${PAGE_ID_VAR} instead.`,
    );
  }
  return page.id;
}

/** The `<title>` of a content document, or null when it carries none. */
function titleOf(html: string): string | null {
  const match = /<title[^>]*>([\s\S]*?)<\/title>/i.exec(html);
  return match?.[1] === undefined ? null : match[1].trim();
}

/** The text blocks of a tool result, for a failure message. */
function resultText(result: { content?: unknown }): string {
  const content = result.content;
  if (!Array.isArray(content)) return '(no content)';
  return content
    .map((block: unknown) =>
      typeof block === 'object' && block !== null && 'text' in block
        ? String((block as { text: unknown }).text)
        : '',
    )
    .join(' ')
    .trim();
}
