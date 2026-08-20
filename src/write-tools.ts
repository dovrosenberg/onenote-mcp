// The writing tools: append_to_page, append_to_page_by_name, create_page,
// create_page_by_name, update_page_title.
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
//
// The two `_by_name` tools take names where the others take ids. They resolve them
// through src/name-lookup.ts, the same resolver the `_by_name` browsing tools use, so a
// caller that knows where the content goes spends one call instead of a walk followed by
// a write. Both check the title and the fragment before resolving anything, so a refusal
// costs neither the lookup nor the write.
//
// `append_to_page_by_name` needs a page as well as a section, and the page title is not
// matched by the container ladder: it is compared in full, ignoring case, by Graph. A
// title matching no page or more than one is an error rather than a choice, because a
// write to a guessed page is invisible until someone opens the notebook.

import type { CallToolResult, Tool } from '@modelcontextprotocol/sdk/types.js';

import type { PageSummary } from './graph-structure.ts';
import { InkParseError, parseInkStrokes, strokeBounds } from './ink.ts';
import { logEvent } from './logging.ts';
import { ToolInputError, optionalString, requiredString, type ToolDefinition } from './mcp-tools.ts';
import {
  NameLookupError,
  resolveSection,
  resolvedPayload,
  type LookupStructure,
} from './name-lookup.ts';
import { pageHtml, type RawPageContent } from './page-content.ts';
import {
  clearanceHtml,
  clearedTo,
  parseOutlines,
  planInkClearance,
  type ClearancePlan,
} from './page-layout.ts';
import type { ResyncOutcome } from './mirror-sync.ts';
import type { CreatedPage } from './page-write.ts';

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

/** The slice of `GraphPageContent` `append_to_page` reads the page's layout through. */
export interface PageLayoutReader {
  fetchRaw(pageId: string): Promise<RawPageContent>;
}

/** The slice of `GraphPageWrite` these tools call, so a test can pass a plain object. */
export interface PageWriteClient {
  appendToPage(pageId: string, html: string): Promise<void>;
  updatePageTitle(pageId: string, title: string): Promise<void>;
  createPage(sectionId: string, title: string, bodyHtml: string): Promise<CreatedPage>;
}

/**
 * The slice of `GraphStructure` the `_by_name` writing tools resolve through: the name
 * resolver's own two calls, plus the title filter that turns a named section and a page
 * title into one page.
 */
export interface WriteLookupStructure extends LookupStructure {
  findPagesByTitle(sectionId: string, title: string): Promise<PageSummary[]>;
}

/** The name arguments both `_by_name` writing tools share. Same wording as the browsing ones. */
const NAME_PATH_PROPERTIES = {
  notebookName: {
    type: 'string',
    description: 'The notebook name, matched in full and case-insensitively.',
  },
  sectionGroupName: {
    type: 'string',
    description:
      'The section group holding the section, if it is in one. Omit when the section ' +
      'sits directly in the notebook; omitting it does not search inside groups.',
  },
  sectionName: {
    type: 'string',
    description: 'The section name, matched in full and case-insensitively.',
  },
} as const;

/**
 * What a successful write tells the mirror.
 *
 * `resyncPage` re-reads the page from Graph and stores it, so the very next read answers
 * from the mirror with what was just written rather than falling through to Graph until
 * the next sync run. `markPageStale` is the fallback for when that fails: a stale marker
 * makes the next read a miss, which is correct but slower.
 *
 * An absent object means no mirror is configured and there is nothing to tell.
 */
export interface MirrorWriteSync {
  resyncPage(
    pageId: string,
    hint: { title?: string; sectionId?: string },
  ): Promise<ResyncOutcome>;
  markPageStale(pageId: string): Promise<void>;
  /** The section a mirrored page sits in, or null when the mirror does not hold it. */
  sectionOfPage(pageId: string): Promise<string | null>;
  holdSectionListing(sectionId: string): Promise<void>;
  releaseSectionListing(sectionId: string): Promise<void>;
}

/**
 * Build the writing tools over a write client, the reader the appends use, and the
 * structure client the two `_by_name` tools resolve names through.
 */
export function createWriteTools(
  write: PageWriteClient,
  layout: PageLayoutReader,
  lookup: WriteLookupStructure,
  mirror?: MirrorWriteSync,
): ToolDefinition[] {
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
        '<html>, <body> or <title> is rejected. Existing content is never replaced. If ' +
        'that outline also holds handwriting, blank lines are added ahead of the ' +
        'content so it lands below the lowest stroke instead of on top of it — OneNote ' +
        'fixes ink in place and no write can move it. pageId comes from list_pages, ' +
        'search_pages or list_pages_by_name.',
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

        return jsonResult(await append(write, layout, pageId, htmlFragment, mirror));
      },
    },
    {
      name: 'append_to_page_by_name',
      title: 'Append to a page in a named section',
      description:
        'Add HTML to the end of a page you can name, in one call: the notebook name, ' +
        'the section group name if the section is in one, the section name, and the ' +
        'page title. This is append_to_page without the walk it would take to turn ' +
        'those names into a page id. Container names are matched in full and ' +
        "case-insensitively first, then against the name with any leading number " +
        "stripped, so 'February' finds a section group named '062 - February'; a " +
        'substring is tried last, and matchedBy in the result says which of those ' +
        'matched. The page title is different: it is matched in full ignoring case and ' +
        'nothing else, and a title matching no page or more than one is an error that ' +
        'writes nothing — use list_pages_by_name to see the titles, or search_pages to ' +
        'match part of one. Omit sectionGroupName when the section sits directly in the ' +
        'notebook; it is not a wildcard. Everything append_to_page does to the content ' +
        'applies here: it goes at the end of the page body, which on a page authored in ' +
        'the OneNote client is the end of its first outline; existing content is never ' +
        'replaced; and blank lines are added ahead of it if handwriting would otherwise ' +
        'be written over. htmlFragment is page content, not a whole HTML document.',
      inputSchema: {
        type: 'object',
        properties: {
          ...NAME_PATH_PROPERTIES,
          pageTitle: {
            type: 'string',
            description:
              'The page title, matched in full and case-insensitively. Not a substring ' +
              'and not a prefix rule: the whole title, or it is an error.',
          },
          htmlFragment: {
            type: 'string',
            description:
              'The HTML to append, as page content rather than a document. Unclosed ' +
              'tags are tolerated; OneNote closes them.',
          },
        },
        required: ['notebookName', 'sectionName', 'pageTitle', 'htmlFragment'],
        additionalProperties: false,
      },
      annotations: { ...WRITE, destructiveHint: false, idempotentHint: false },
      handle: async (args) => {
        // The fragment is checked before anything is resolved, so a refusal costs
        // neither of the two reads this tool would otherwise spend finding the page.
        const path = namePath(args);
        const pageTitle = requiredString(args, 'pageTitle');
        const htmlFragment = fragmentArgument(args, 'htmlFragment');

        const resolved = await resolveSection(lookup, path);
        const page = await onePageByTitle(lookup, resolved.section.id, pageTitle);

        return jsonResult({
          ...resolvedPayload(resolved),
          page: { id: page.id, title: page.title },
          ...(await append(write, layout, page.id, htmlFragment, mirror)),
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

        // The section's page listing is held back first: a create has no page to mark
        // stale, and the listing is the thing it makes wrong. See `beginWrite`.
        await beginWrite(mirror, { sectionId });
        try {
          const page = await write.createPage(sectionId, title, htmlFragment);
          // sectionId is passed because the page is not in the mirror yet, so there is no
          // stored placement to read it from. The title comes from the create response,
          // which carries the right one — a metadata read here would not.
          await resync(
            mirror,
            page.id,
            { title: page.title === '' ? title : page.title, sectionId },
            true,
          );

          return jsonResult({
            pageId: page.id,
            title: page.title === '' ? title : page.title,
            webUrl: page.webUrl,
            clientUrl: page.clientUrl,
            note:
              'The page was created. Use pageId with get_page_content to read it back ' +
              'or append_to_page to add to it. Do not confirm the page by looking for ' +
              'its title in a listing straight away: page metadata can take a few ' +
              'seconds to catch up, and the title above is the one the service accepted.',
          });
        } finally {
          await endWrite(mirror, { sectionId });
        }
      },
    },
    {
      name: 'create_page_by_name',
      title: 'Create a page in a named section',
      description:
        'Create a new OneNote page in a section you can name, in one call: the notebook ' +
        'name, the section group name if the section is in one, the section name, the ' +
        'page title, and the body HTML. This is create_page without the list_notebooks ' +
        '-> list_sections walk it would take to turn those names into a section id. ' +
        'Container names are matched in full and case-insensitively first, then against ' +
        "the name with any leading number stripped, so 'February' finds a section group " +
        "named '062 - February'; a substring is tried last, and matchedBy in the result " +
        'says which of those matched. Omit sectionGroupName when the section sits ' +
        'directly in the notebook; it is not a wildcard. A name that matches nothing, or ' +
        'matches more than one thing, is an error listing the candidates and no page is ' +
        'created. The title becomes the page title OneNote shows and the one every ' +
        'search and by-name lookup matches on. htmlFragment is the body content — <p>, ' +
        '<ul>, <table>, <h1> and so on — not a whole HTML document, and it must not ' +
        'carry its own <title>: this tool builds the document around it. Returns the new ' +
        'page id and links to open the page.',
      inputSchema: {
        type: 'object',
        properties: {
          ...NAME_PATH_PROPERTIES,
          title: {
            type: 'string',
            description: 'The page title. Plain text; markup in it is not rendered.',
          },
          htmlFragment: {
            type: 'string',
            description: 'The body content, as page HTML rather than a whole document.',
          },
        },
        required: ['notebookName', 'sectionName', 'title', 'htmlFragment'],
        additionalProperties: false,
      },
      annotations: { ...WRITE, destructiveHint: false, idempotentHint: false },
      handle: async (args) => {
        // Every argument is checked before the lookup runs, so a fragment that was
        // going to be refused costs no Graph request at all — not even the one that
        // turns the names into a section id.
        const path = namePath(args);
        const title = titleArgument(args, 'title');
        const htmlFragment = fragmentArgument(args, 'htmlFragment');

        const resolved = await resolveSection(lookup, path);
        const sectionId = resolved.section.id;

        await beginWrite(mirror, { sectionId });
        try {
          const page = await write.createPage(sectionId, title, htmlFragment);
          await resync(
            mirror,
            page.id,
            { title: page.title === '' ? title : page.title, sectionId },
            true,
          );

          return jsonResult({
            ...resolvedPayload(resolved),
            pageId: page.id,
            title: page.title === '' ? title : page.title,
            webUrl: page.webUrl,
            clientUrl: page.clientUrl,
            note:
              'The page was created in the section named above. Check section and ' +
              'matchedBy if the names you gave were approximate. Use pageId with ' +
              'get_page_content to read it back or append_to_page to add to it. Do not ' +
              'confirm the page by looking for its title in a listing straight away: ' +
              'page metadata can take a few seconds to catch up, and the title above ' +
              'is the one the service accepted.' +
              (resolved.deepSearchUsed
                ? ' The section sits below the section group named, and was found by ' +
                  'an account-wide search on its name.'
                : ''),
          });
        } finally {
          await endWrite(mirror, { sectionId });
        }
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

        // A rename is the one write that changes what a *listing* says, so it holds the
        // section's page listing back as well as marking the page stale. The section is
        // read from the mirror rather than taken as an argument: this tool names only a
        // page, and a page the mirror does not hold appears in no listing to be wrong.
        const sectionId = await sectionOfPage(mirror, pageId);

        await beginWrite(mirror, { pageId, ...(sectionId === null ? {} : { sectionId }) });
        try {
          await write.updatePageTitle(pageId, newTitle);
          // The title comes from here rather than from a read-back: measured 2026-08-19,
          // page metadata read immediately after a write can come back empty.
          await resync(mirror, pageId, { title: newTitle }, false);

          return jsonResult({
            pageId,
            title: newTitle,
            note:
              'The page title was replaced and the page content was left alone. A ' +
              'listing may report the old title for a few seconds afterwards.',
          });
        } finally {
          await endWrite(mirror, { ...(sectionId === null ? {} : { sectionId }) });
        }
      },
    },
  ];
}

/**
 * Append to one page and describe what happened, for both tools that append.
 *
 * The ink clearance is the reason this is shared rather than repeated: a second copy
 * that forgot to read the page first would write over someone's handwriting and report
 * success, and the failure would only be visible in OneNote.
 */
async function append(
  write: PageWriteClient,
  layout: PageLayoutReader,
  pageId: string,
  htmlFragment: string,
  mirror?: MirrorWriteSync,
): Promise<Record<string, unknown>> {
  const clearance = await inkClearance(layout, pageId);
  const content = clearance === null ? htmlFragment : clearanceHtml(clearance, htmlFragment);

  await beginWrite(mirror, { pageId });
  await write.appendToPage(pageId, content);
  // No title hint: an append cannot change one.
  await resync(mirror, pageId, {}, true);

  return {
    pageId,
    appended: true,
    inkClearance:
      clearance === null
        ? null
        : {
            blankLines: clearance.breaks,
            inkBottomPx: Math.round(clearance.inkBottom),
            measuredFromPx: Math.round(clearance.clearedFrom),
            contentStartsAtPx: clearance.clearedTo,
            clearsTheInk: !clearance.truncated,
          },
    note:
      'The content was appended to the end of the page body. On a page authored ' +
      'in the OneNote client that is the end of the first outline. Nothing that ' +
      'was on the page was replaced. A read of the page content will show the ' +
      'change; the page metadata a listing returns can lag behind it by a few ' +
      'seconds.' +
      (clearance === null
        ? ''
        : ' This outline also holds handwriting, which OneNote keeps at a fixed ' +
          'position that no write can move, so the content was preceded by blank ' +
          'lines to bring it below the lowest stroke. The gap above it is ' +
          'deliberate.'),
  };
}

/**
 * The one page in `sectionId` titled `pageTitle`.
 *
 * Graph does the comparison, case-insensitively and across the whole section, so no
 * bound on how many pages the section holds can hide a match. Zero and more than one are
 * both errors: this is a write, and appending to a guessed page puts someone's content
 * on a page they did not name with nothing to say so.
 *
 * @throws {NameLookupError} when the title matches no page or more than one.
 */
async function onePageByTitle(
  lookup: WriteLookupStructure,
  sectionId: string,
  pageTitle: string,
): Promise<PageSummary> {
  const matches = await lookup.findPagesByTitle(sectionId, pageTitle);
  const first = matches[0];
  if (matches.length === 1 && first !== undefined) return first;

  throw new NameLookupError(
    matches.length === 0 ? 'not-found' : 'ambiguous',
    'pageTitle',
    pageTitle,
    matches.map((page) => ({ id: page.id, displayName: page.title })),
    'page-title',
  );
}

/** The three name arguments, read the same way by both `_by_name` writing tools. */
function namePath(args: Readonly<Record<string, unknown>>): {
  notebookName: string;
  sectionGroupName: string | undefined;
  sectionName: string;
} {
  return {
    notebookName: requiredString(args, 'notebookName'),
    sectionGroupName: optionalString(args, 'sectionGroupName'),
    sectionName: requiredString(args, 'sectionName'),
  };
}

/**
 * What this append has to do about the handwriting already on the page, or null.
 *
 * This costs one extra Graph request per append, which is the price of not writing text
 * across someone's handwriting: no endpoint reports an outline's rendered height or the
 * ink's position, so both have to be read.
 *
 * An unreadable ink document is not a reason to refuse the write. `parseInkStrokes`
 * throws only when trace groups nest past its bound, and the caller's content is good
 * either way — it goes on the page unpadded.
 */
async function inkClearance(
  layout: PageLayoutReader,
  pageId: string,
): Promise<ClearancePlan | null> {
  const raw = await layout.fetchRaw(pageId);
  const html = pageHtml(raw);
  if (html === null) return null;

  // `target: "body"` is the first top-level div, so that outline is the one being
  // appended to and the only one whose column matters.
  const outline = parseOutlines(html)[0];
  if (outline === undefined) return null;

  let ink;
  try {
    ink = strokeBounds(parseInkStrokes(raw.raw));
  } catch (err) {
    if (!(err instanceof InkParseError)) throw err;
    return null;
  }

  return planInkClearance(outline, ink, clearedTo(html));
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

/**
 * Invalidate what this write is about to make wrong, *before* the Graph write.
 *
 * Every one of the five writing tools calls this first, then writes to OneNote, then
 * resyncs. That ordering is what makes the window between the two safe, and it covers a
 * case the fallback inside `resync` cannot. Between the PATCH succeeding and the resync
 * completing, the mirror holds a copy that OneNote no longer agrees with. If the process
 * merely errors, `resync` catches it and marks the page stale. If the process *stops* —
 * Cloud Run cutting the request at 300 seconds, or the instance being reclaimed after an
 * idle period — nothing catches anything: the resync is sitting in the request gate's
 * queue, the queue is wiped with the instance, and no `catch` ever runs because nothing
 * threw. The mirror would then keep serving superseded data as `present`, reporting
 * `source: "mirror"` and a recent `mirroredAt`, until the next scheduled sync noticed the
 * page's `lastModifiedDateTime` had moved. Marking first makes the whole window
 * pessimistic instead: a death anywhere in it leaves a miss, and a miss goes to Graph,
 * which cannot be wrong.
 *
 * **There are two things to invalidate, and which ones apply is per tool.**
 *
 * `pageId` marks the page's stored *content* stale, so `get_page_content` misses. Every
 * write that names an existing page passes it: both appends and the rename.
 *
 * `sectionId` holds the section's *page listing* back, so `list_pages`,
 * `list_pages_by_name`, `find_page_by_name` and `search_pages` miss. `create_page` passes
 * it because there is no page document to mark and the listing is exactly what a create
 * makes wrong — without this, a create whose resync never ran leaves `list_pages`
 * answering from the mirror with a section that does not contain the page, which reads to
 * a model as "the page was not created". `update_page_title` passes it because the title
 * is what every listing and by-name lookup matches on. An append passes neither a section
 * nor anything else: it changes content, which `pageId` covers, and
 * `lastModifiedDateTime`, which only reorders a listing.
 *
 * The cost is one Firestore write per tool write, and — if the Graph write then fails —
 * a discarded cached copy that the next read re-fetches. Both are cheap against serving
 * data the user can see is wrong.
 */
async function beginWrite(
  mirror: MirrorWriteSync | undefined,
  target: { pageId?: string; sectionId?: string },
): Promise<void> {
  if (mirror === undefined) return;

  const { pageId, sectionId } = target;

  if (pageId !== undefined) {
    try {
      await mirror.markPageStale(pageId);
    } catch (err) {
      // Never fails the write. This is a narrowing of an existing window, not a
      // precondition for writing.
      logEvent('mirror-invalidate-failed', { pageId, reason: reasonOf(err) });
    }
  }

  if (sectionId !== undefined) {
    try {
      await mirror.holdSectionListing(sectionId);
    } catch (err) {
      logEvent('mirror-listing-hold-failed', { sectionId, reason: reasonOf(err) });
    }
  }
}

/**
 * Release the listing hold `beginWrite` took, after the resync has landed.
 *
 * Called from a `finally`, so a Graph write that failed releases too — the listing it
 * was going to invalidate is still correct. A hold that outlives its process is not
 * leaked forever: `listingIsHeld` in ./mirror-schema.ts expires one on age.
 */
async function endWrite(
  mirror: MirrorWriteSync | undefined,
  target: { sectionId?: string },
): Promise<void> {
  if (mirror === undefined || target.sectionId === undefined) return;

  try {
    await mirror.releaseSectionListing(target.sectionId);
  } catch (err) {
    logEvent('mirror-listing-release-failed', {
      sectionId: target.sectionId,
      reason: reasonOf(err),
    });
  }
}

/**
 * Where the mirror thinks this page lives, or null.
 *
 * Null covers both "the mirror does not hold it" and "Firestore did not answer", and both
 * mean the same thing to the caller: there is no stored listing to hold back. A failure
 * here must not fail the write, which has not happened yet and is the thing the caller
 * asked for.
 */
async function sectionOfPage(
  mirror: MirrorWriteSync | undefined,
  pageId: string,
): Promise<string | null> {
  if (mirror === undefined) return null;

  try {
    return await mirror.sectionOfPage(pageId);
  } catch (err) {
    logEvent('mirror-invalidate-failed', { pageId, reason: reasonOf(err) });
    return null;
  }
}

/**
 * Bring the mirror's copy of this page up to date, right now.
 *
 * Called after a successful write and never before one: a page whose write failed still
 * matches what the mirror holds.
 *
 * The resync costs one Graph request. What it buys is that a `get_page_content` straight
 * after an `append_to_page` answers from the mirror with the appended text — without it,
 * the page is marked stale and every read falls through to Graph until the next sync run,
 * which on a fifteen-minute schedule is a long window in the middle of a conversation.
 *
 * **Two failure levels, and neither fails the write.** A resync that throws falls back to
 * marking the page stale, which makes the next read a miss — correct, just slower. If
 * that fails too, the event is logged and nothing else happens. The write is the thing
 * that mattered and it has already happened; turning a successful append into a reported
 * error would send the caller to retry a change that is already made. It is also
 * self-healing: the write moved the page's `lastModifiedDateTime`, so the next
 * incremental run repairs whatever this could not.
 */
async function resync(
  mirror: MirrorWriteSync | undefined,
  pageId: string,
  hint: { title?: string; sectionId?: string },
  expectContentChange: boolean,
): Promise<void> {
  if (mirror === undefined) return;

  try {
    const outcome = await mirror.resyncPage(pageId, hint);

    // An append or a create always changes the page's content, so a resync that found
    // nothing to write did not read what was just written. Measured 2026-08-19, a PATCH
    // is visible to the next content read at 3.7 seconds — but that is one observation,
    // and if the read ever does lose the race the stored copy is pre-write content
    // marked `present`, which the read path would serve as current with nothing saying
    // so. Marking it stale is the safe direction: the next read misses and goes to
    // Graph, which cannot be wrong.
    //
    // A rename is the opposite case and must not fall through here: it changes no
    // content by design, so `unchanged` would be the normal answer — except that the
    // title comparison in `writePageFromRaw` makes it `updated`, which is why this is
    // reached only when something really did not take.
    if (outcome !== 'unchanged' || !expectContentChange) return;

    logEvent('mirror-resync-stale-read', { pageId });
  } catch (err) {
    logEvent('mirror-resync-failed', { pageId, reason: reasonOf(err) });
  }

  try {
    await mirror.markPageStale(pageId);
  } catch (err) {
    logEvent('mirror-invalidate-failed', { pageId, reason: reasonOf(err) });
  }
}

/** A reason string for a log line. Never a message, which can carry a request body. */
function reasonOf(err: unknown): string {
  return err instanceof Error ? err.name : 'unknown';
}
