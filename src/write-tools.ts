// The writing tools: append_to_page, create_page, update_page_title.
//
// The Graph mechanics are in ./page-write.ts. What lives here is the argument checking
// and the wording, and both exist for the same reason: the caller is a model that gets
// one shot at each call and cannot see the page it is writing to.
//
// Two things are checked before a request is spent:
//
// A fragment must be a fragment. `<html>`, `<head>`, `<body>`, `<title>`, `<meta>`,
// `<script>` and `<style>` are rejected rather than submitted. The PATCH `content` is a
// piece of page content, not a document; a model that sends a whole document gets its
// wrapper silently dropped and its `<title>` ignored, which looks like the tool losing
// the title it asked for.
//
// A title carries no tags. `update_page_title` stores its argument verbatim — the spike
// in issue #17 produced a page actually titled `<p>x</p>` — so a tag in a title becomes
// part of the title text. Rejecting it here is the only place that can tell the caller
// so; Graph answers 204 and the mistake is only visible in the notebook. A bare `<` that
// opens no tag is left alone, because `if x <y then` is a legal title.
//
// What is deliberately *not* validated is well-formedness. `<p>unclosed` returned 204 and
// the service closed the tag, so a strict parser here would refuse content that works.
//
// Every result is one JSON text block, and each one says where the change landed. A
// caller that appended to a client-authored page needs to know it wrote to the first
// outline rather than the bottom of the page, and nothing in a 204 tells it that.

import type { CallToolResult, Tool } from '@modelcontextprotocol/sdk/types.js';

import type { CreatedPage } from './page-write.ts';
import { ToolInputError, requiredString, type ToolDefinition } from './mcp-tools.ts';

/** These tools write. `openWorldHint` because the world is someone's notebook. */
const WRITE: Tool['annotations'] = { readOnlyHint: false, openWorldHint: true };

/** Elements that make a fragment a document. None of them can survive a PATCH. */
const DOCUMENT_ELEMENTS = ['html', 'head', 'body', 'title', 'meta', 'base', 'link'];

/** Elements OneNote will not run or apply, and that a model should not be sending. */
const INERT_ELEMENTS = ['script', 'style'];

const FORBIDDEN_ELEMENT =
  new RegExp(`<\\s*/?\\s*(${[...DOCUMENT_ELEMENTS, ...INERT_ELEMENTS].join('|')})\\b`, 'i');

/**
 * A complete tag: `<b>`, `</b>`, `<p class="x">`.
 *
 * Deliberately not "contains a `<`". A title like `if x <y then` is legal and this tool
 * is the only way to set one, so a rule that rejected every `<` would make a legal title
 * unreachable. A `<` with no closing `>` after it is left alone; a whole tag is the case
 * where the caller meant markup and would get the characters instead.
 */
const LOOKS_LIKE_MARKUP = /<\/?[a-zA-Z][^<>]*>/;

/** The slice of `GraphPageWrite` these tools call, so a test can pass a plain object. */
export interface PageWriteClient {
  appendToPage(pageId: string, html: string): Promise<void>;
  updatePageTitle(pageId: string, title: string): Promise<void>;
  createPage(sectionId: string, title: string, bodyHtml: string): Promise<CreatedPage>;
}

/** Build the three writing tools over one write client. */
export function createWriteTools(write: PageWriteClient): ToolDefinition[] {
  return [
    {
      name: 'append_to_page',
      title: 'Append to a page',
      description:
        'Add HTML to the end of an existing OneNote page. The content goes at the end ' +
        "of the page's body element, which for a page written by a person in the " +
        'OneNote client means the end of its first outline — not the bottom of the ' +
        'page, which may hold other outlines beside it. For a page created by ' +
        'create_page it is the bottom of the page. htmlFragment is page content such ' +
        '<p>, <ul>, <table> or <h1>, not a whole HTML document: a fragment carrying ' +
        '<html>, <body> or <title> is rejected. Existing content is never replaced. ' +
        'pageId comes from list_pages, search_pages or list_pages_by_name.',
      inputSchema: {
        type: 'object',
        properties: {
          pageId: {
            type: 'string',
            description: 'A page id from list_pages, search_pages or list_pages_by_name.',
          },
          htmlFragment: {
            type: 'string',
            description:
              'The HTML to append, as page content rather than a document. Unclosed ' +
              'tags are tolerated; OneNote closes them.',
          },
        },
        required: ['pageId', 'htmlFragment'],
        additionalProperties: false,
      },
      annotations: { ...WRITE, destructiveHint: false, idempotentHint: false },
      handle: async (args) => {
        const pageId = requiredString(args, 'pageId');
        const htmlFragment = fragmentArgument(args, 'htmlFragment');

        await write.appendToPage(pageId, htmlFragment);

        return jsonResult({
          pageId,
          appended: true,
          note:
            'The content was appended to the end of the page body. On a page authored ' +
            'in the OneNote client that is the end of the first outline. Nothing that ' +
            'was on the page was replaced. A read of the page content will show the ' +
            'change; the page metadata a listing returns can lag behind it by a few ' +
            'seconds.',
        });
      },
    },
    {
      name: 'create_page',
      title: 'Create a page',
      description:
        'Create a new OneNote page in a section, with a real page title and a body of ' +
        'HTML. The title becomes the page title OneNote shows and the one every ' +
        'search and by-name lookup matches on, so it is worth making it specific. ' +
        'htmlFragment is the body content — <p>, <ul>, <table>, <h1> and so on — not a ' +
        'whole HTML document, and it must not carry its own <title>: this tool builds ' +
        'the document around it. Returns the new page id and links to open the page. ' +
        'sectionId comes from list_sections or find_page_by_name.',
      inputSchema: {
        type: 'object',
        properties: {
          sectionId: {
            type: 'string',
            description: 'A section id from list_sections or one of the by-name lookups.',
          },
          title: {
            type: 'string',
            description: 'The page title. Plain text; markup in it is not rendered.',
          },
          htmlFragment: {
            type: 'string',
            description: 'The body content, as page HTML rather than a whole document.',
          },
        },
        required: ['sectionId', 'title', 'htmlFragment'],
        additionalProperties: false,
      },
      annotations: { ...WRITE, destructiveHint: false, idempotentHint: false },
      handle: async (args) => {
        const sectionId = requiredString(args, 'sectionId');
        const title = titleArgument(args, 'title');
        const htmlFragment = fragmentArgument(args, 'htmlFragment');

        const page = await write.createPage(sectionId, title, htmlFragment);

        return jsonResult({
          pageId: page.id,
          title: page.title === '' ? title : page.title,
          webUrl: page.webUrl,
          clientUrl: page.clientUrl,
          note:
            'The page was created. Use pageId with get_page_content to read it back or ' +
            'append_to_page to add to it. Do not confirm the page by looking for its ' +
            'title in a listing straight away: page metadata can take a few seconds to ' +
            'catch up, and the title above is the one the service accepted.',
        });
      },
    },
    {
      name: 'update_page_title',
      title: 'Rename a page',
      description:
        "Replace an existing page's title. The new title is stored exactly as given: " +
        'nothing in it is parsed, so markup would appear literally in the title and is ' +
        'rejected. The page content is untouched. This overwrites the old title, which ' +
        'cannot be recovered through this server. pageId comes from list_pages, ' +
        'search_pages or list_pages_by_name.',
      inputSchema: {
        type: 'object',
        properties: {
          pageId: {
            type: 'string',
            description: 'A page id from list_pages, search_pages or list_pages_by_name.',
          },
          newTitle: {
            type: 'string',
            description: 'The replacement title, as plain text.',
          },
        },
        required: ['pageId', 'newTitle'],
        additionalProperties: false,
      },
      // The old title is gone, so this is the one write here that destroys something.
      // It is idempotent: sending the same title twice leaves the same page.
      annotations: { ...WRITE, destructiveHint: true, idempotentHint: true },
      handle: async (args) => {
        const pageId = requiredString(args, 'pageId');
        const newTitle = titleArgument(args, 'newTitle');

        await write.updatePageTitle(pageId, newTitle);

        return jsonResult({
          pageId,
          title: newTitle,
          note:
            'The page title was replaced and the page content was left alone. A ' +
            'listing may report the old title for a few seconds afterwards.',
        });
      },
    },
  ];
}

/**
 * A fragment of page content.
 *
 * @throws {ToolInputError} if it is missing, empty, or carries a document element.
 */
export function fragmentArgument(args: Readonly<Record<string, unknown>>, name: string): string {
  const value = requiredString(args, name);
  const found = FORBIDDEN_ELEMENT.exec(value);
  if (found !== null) {
    throw new ToolInputError(
      name,
      `must be a fragment of page content, not a whole HTML document — remove the <${(found[1] ?? '').toLowerCase()}> element and pass only the content that goes on the page`,
    );
  }
  return value;
}

/**
 * A page title.
 *
 * @throws {ToolInputError} if it is missing, empty, or looks like markup. A title is
 * stored verbatim, so `<p>x</p>` would become a page literally called `<p>x</p>`.
 */
export function titleArgument(args: Readonly<Record<string, unknown>>, name: string): string {
  const value = requiredString(args, name);
  if (LOOKS_LIKE_MARKUP.test(value)) {
    throw new ToolInputError(
      name,
      'must not contain HTML tags; a title is stored exactly as given, so a tag in it becomes part of the title text rather than formatting it',
    );
  }
  return value;
}

/** One JSON text block, the shape every other tool in this server answers with. */
function jsonResult(payload: unknown): CallToolResult {
  return { content: [{ type: 'text', text: JSON.stringify(payload, null, 2) }] };
}
