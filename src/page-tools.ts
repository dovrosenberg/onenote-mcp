// The reading tool: get_page_content.
//
// One call answers with both halves of a page, because one Graph fetch produces both —
// see the comment at the top of ./page-content.ts for why `includeInkML=true` is the
// only way to see handwriting at all.
//
// The answer is two content blocks, not one. The typed content and the metadata go in a
// JSON text block; the handwriting goes in an MCP **image** block, so the calling model
// reads it with its own vision. A base64 blob inside the JSON would reach the model as
// text it cannot see, and a file path would point at a container filesystem no client
// can read. There is no OCR service anywhere in this path.
//
// All of a page's ink is one cropped image. Ink and typed content are independent on
// these pages — the handwriting is a note beside the typing, not interleaved with it —
// so nothing is lost by dropping the spatial relationship, and reconstructing it would
// mean splicing ink fragments into the HTML.
//
// A page with no ink yields no image block. That is the normal answer for a typed page
// and is not an error; the JSON says `inkImage: null` so the model is not left guessing
// whether an image was meant to arrive.

import type { CallToolResult, Tool } from '@modelcontextprotocol/sdk/types.js';

import {
  DEFAULT_RENDER_WIDTH,
  MAX_INK_PNG_BYTES,
  MIN_RENDER_WIDTH,
  fitInkToByteBudget,
  type InkImage,
} from './ink.ts';
import type { PageContent } from './page-content.ts';
import { requiredString, type ToolDefinition } from './mcp-tools.ts';

/** This tool reads; it does not write. */
const READ_ONLY: Tool['annotations'] = { readOnlyHint: true, openWorldHint: true };

/** The slice of `GraphPageContent` this tool calls, so a test can pass a plain object. */
export interface PageContentClient {
  fetchContent(pageId: string): Promise<PageContent>;
}

/** How the image was sized, and what the JSON block reports about it. */
export interface InkImageSummary {
  readonly strokeCount: number;
  readonly width: number;
  readonly height: number;
  readonly bytes: number;
  /** True when the byte budget forced a render narrower than DEFAULT_RENDER_WIDTH. */
  readonly downscaled: boolean;
}

/**
 * Build the reading tool over one page-content client.
 *
 * `maxBytes` exists so a test can force the downscale path without a fixture large
 * enough to blow the real budget; production passes nothing and gets MAX_INK_PNG_BYTES.
 */
export function createPageTools(
  content: PageContentClient,
  maxBytes: number = MAX_INK_PNG_BYTES,
): ToolDefinition[] {
  return [
    {
      name: 'get_page_content',
      title: 'Read a page',
      description:
        'Read one OneNote page. Returns two things: a JSON block holding the page HTML ' +
        'with the styling noise stripped out, and — when the page has handwriting — a ' +
        'PNG image block of that handwriting for you to read directly. There is no OCR ' +
        'behind this; the image is the ink, and reading it is your job. A page that was ' +
        'typed rather than written comes back with inkImage null and no image block, ' +
        'which is normal and not an error. All of the page\'s ink is cropped into one ' +
        'image, so the strokes keep no positional relationship to the HTML. pageId ' +
        'comes from list_pages or search_pages.',
      inputSchema: {
        type: 'object',
        properties: {
          pageId: {
            type: 'string',
            description: 'A page id from list_pages or search_pages.',
          },
        },
        required: ['pageId'],
        additionalProperties: false,
      },
      annotations: READ_ONLY,
      handle: async (args) => {
        const pageId = requiredString(args, 'pageId');
        const page = await content.fetchContent(pageId);
        const ink = page.ink === null ? null : fitInkToByteBudget(page.ink, maxBytes);
        return pageResult(pageId, page.html, ink);
      },
    },
  ];
}

/** The JSON block, then the image block when there is one. */
export function pageResult(
  pageId: string,
  html: string | null,
  ink: InkImage | null,
): CallToolResult {
  const summary = ink === null ? null : summarise(ink);

  const content: CallToolResult['content'] = [
    {
      type: 'text',
      text: JSON.stringify(
        { pageId, html, inkImage: summary, note: note(html, summary) },
        null,
        2,
      ),
    },
  ];

  if (ink !== null) {
    content.push({
      type: 'image',
      // Buffer rather than btoa: the PNG is bytes, and btoa needs a binary string built
      // one character at a time, which for a megabyte of image is a real cost.
      data: Buffer.from(ink.png).toString('base64'),
      mimeType: 'image/png',
    });
  }

  return { content };
}

function summarise(ink: InkImage): InkImageSummary {
  return {
    strokeCount: ink.strokeCount,
    width: ink.width,
    height: ink.height,
    bytes: ink.png.byteLength,
    downscaled: ink.width < DEFAULT_RENDER_WIDTH,
  };
}

/**
 * One sentence about what did and did not arrive.
 *
 * A model that receives no image block has no way to tell "this page was typed" from
 * "the image was dropped", and a model reading a shrunken render has no way to tell
 * "this handwriting is illegible" from "this render is too small to read". Both cases
 * are stated rather than inferred.
 */
function note(html: string | null, ink: InkImageSummary | null): string {
  const parts: string[] = [];

  if (html === null) {
    parts.push('Graph returned no HTML for this page.');
  }

  if (ink === null) {
    parts.push('This page has no handwriting, so no image block follows. That is normal.');
  } else {
    parts.push(
      `The image block that follows is ${ink.strokeCount} handwritten strokes cropped to ` +
        `their bounding box, rendered ${ink.width}x${ink.height}. Read it directly; it is ` +
        'not transcribed anywhere.',
    );
    if (ink.downscaled) {
      parts.push(
        `It was rendered narrower than the usual ${DEFAULT_RENDER_WIDTH}px to fit the ` +
          `response size budget (floor ${MIN_RENDER_WIDTH}px), so it may be harder to read.`,
      );
    }
  }

  return parts.join(' ');
}
