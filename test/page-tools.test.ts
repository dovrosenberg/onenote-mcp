// get_page_content, driven through its own `handle` with a fake page-content client.
//
// The ink in these results is rendered from the committed InkML fixture rather than
// stubbed, so the assertions about the image block are assertions about real PNG bytes:
// the block's base64 decodes to a PNG signature, and the width in the JSON matches the
// width in the file's own IHDR. What no test covers is whether a client displays the
// block to its model — nothing confirms that until an operator points one at the
// deployed URL.

import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';

import type { CallToolResult } from '@modelcontextprotocol/sdk/types.js';

import { DEFAULT_RENDER_WIDTH, MIN_RENDER_WIDTH, renderInk, type InkImage } from '../src/ink.ts';
import { GraphRequestError } from '../src/graph-structure.ts';
import { ToolInputError, indexTools, type ToolDefinition } from '../src/mcp-tools.ts';
import type { PageContent } from '../src/page-content.ts';
import { createPageTools, type PageContentClient } from '../src/page-tools.ts';

const FIXTURES = path.join(import.meta.dirname, 'fixtures');

const HTML = '<div><p>typed content</p></div>';

/** The fixture's ink, rendered once: rasterising is the slowest thing in this file. */
const INK: InkImage = await (async () => {
  const image = renderInk(await readFile(path.join(FIXTURES, 'xyf-himetric.inkml'), 'utf8'));
  assert.ok(image !== null, 'the fixture must yield ink for these tests to mean anything');
  return image;
})();

interface Fake extends PageContentClient {
  readonly calls: string[];
}

function fakeContent(page: PageContent | (() => never)): Fake {
  const calls: string[] = [];
  return {
    calls,
    fetchContent: (pageId: string) => {
      calls.push(pageId);
      if (typeof page === 'function') page();
      return Promise.resolve(page as PageContent);
    },
  };
}

function tool(client: PageContentClient, maxBytes?: number): ToolDefinition {
  const tools = createPageTools(client, maxBytes);
  const found = indexTools(tools).get('get_page_content');
  assert.ok(found !== undefined);
  return found;
}

interface Payload {
  pageId: string;
  html: string | null;
  inkImage: { strokeCount: number; width: number; height: number; downscaled: boolean } | null;
  note: string;
}

function payload(result: CallToolResult): Payload {
  const first = result.content[0];
  assert.equal(first?.type, 'text');
  return JSON.parse((first as { text: string }).text) as Payload;
}

/** Width and height out of a PNG's IHDR, so the assertion is about the file itself. */
function pngSize(png: Buffer): { width: number; height: number } {
  assert.deepEqual(
    [...png.subarray(0, 8)],
    [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a],
    'expected a PNG signature',
  );
  return { width: png.readUInt32BE(16), height: png.readUInt32BE(20) };
}

test('a page with ink answers with the HTML and a real PNG image block', async () => {
  const client = fakeContent({ html: HTML, ink: INK });
  const result = await tool(client).handle({ pageId: 'p-1' });

  assert.deepEqual(client.calls, ['p-1']);
  assert.equal(result.isError, undefined);

  const body = payload(result);
  assert.equal(body.pageId, 'p-1');
  assert.equal(body.html, HTML);
  assert.equal(body.inkImage?.strokeCount, INK.strokeCount);
  assert.equal(body.inkImage?.downscaled, false);

  // The image is an MCP image block, not a path and not base64 inside the text block.
  assert.equal(result.content.length, 2);
  const image = result.content[1] as { type: string; data: string; mimeType: string };
  assert.equal(image.type, 'image');
  assert.equal(image.mimeType, 'image/png');

  const decoded = Buffer.from(image.data, 'base64');
  assert.deepEqual(pngSize(decoded), { width: body.inkImage?.width, height: body.inkImage?.height });
  assert.equal(decoded.byteLength, INK.png.byteLength);
});

test('the JSON block never carries the image bytes', async () => {
  const result = await tool(fakeContent({ html: HTML, ink: INK })).handle({ pageId: 'p-1' });
  const text = (result.content[0] as { text: string }).text;
  // Base64 of a PNG starts iVBORw0KGgo. Finding it in the text block would mean the
  // image reached the model as characters it cannot see.
  assert.equal(text.includes('iVBORw0KGgo'), false);
});

test('a typed page answers with html, inkImage null, and no image block', async () => {
  const result = await tool(fakeContent({ html: HTML, ink: null })).handle({ pageId: 'p-2' });

  assert.equal(result.isError, undefined, 'a page with no ink is normal, not an error');
  assert.equal(result.content.length, 1);

  const body = payload(result);
  assert.equal(body.inkImage, null);
  assert.match(body.note, /no handwriting/i);
});

test('a response that carried no HTML says so rather than reporting an empty page', async () => {
  const result = await tool(fakeContent({ html: null, ink: null })).handle({ pageId: 'p-3' });
  const body = payload(result);
  assert.equal(body.html, null);
  assert.match(body.note, /no HTML/i);
});

test('an oversized image is downscaled and the result says it was', async () => {
  // One byte of budget forces every shrink step, so the floor is what is reached. The
  // fixture is far too small to blow the real budget, which is why maxBytes is injected.
  const result = await tool(fakeContent({ html: HTML, ink: INK }), 1).handle({ pageId: 'p-1' });

  const body = payload(result);
  assert.equal(body.inkImage?.downscaled, true);
  assert.equal(body.inkImage?.width, MIN_RENDER_WIDTH);
  assert.ok((body.inkImage?.width ?? 0) < DEFAULT_RENDER_WIDTH);
  assert.match(body.note, /narrower/i);

  const image = result.content[1] as { data: string };
  assert.equal(pngSize(Buffer.from(image.data, 'base64')).width, MIN_RENDER_WIDTH);
});

test('a missing pageId is a ToolInputError naming the argument', async () => {
  await assert.rejects(
    () => tool(fakeContent({ html: HTML, ink: null })).handle({}),
    (err: unknown) => err instanceof ToolInputError && err.argument === 'pageId',
  );
});

test('a Graph failure propagates for the caller to map, rather than being swallowed', async () => {
  const failing = fakeContent(() => {
    throw new GraphRequestError('https://graph.microsoft.com/v1.0/x', 404, 'Not Found', '{}');
  });
  await assert.rejects(
    () => tool(failing).handle({ pageId: 'gone' }),
    (err: unknown) => err instanceof GraphRequestError,
  );
});

test('the declared schema takes exactly one required argument', () => {
  const definition = tool(fakeContent({ html: HTML, ink: null }));
  assert.deepEqual(definition.inputSchema.required, ['pageId']);
  assert.equal(definition.inputSchema.additionalProperties, false);
  assert.equal(definition.annotations?.readOnlyHint, true);
});
