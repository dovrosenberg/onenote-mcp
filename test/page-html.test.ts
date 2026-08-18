import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';

import { trimPageHtml } from '../src/page-html.ts';

function fixture(name: string): Promise<string> {
  return readFile(path.join(import.meta.dirname, 'fixtures', name), 'utf8');
}

/** The visible words of a document, in order. Tags out, whitespace collapsed. */
function words(html: string): string[] {
  return html
    .replaceAll(/<[^>]*>/g, ' ')
    .split(/\s+/)
    .filter((word) => word !== '');
}

test('every word of a heavily styled page survives the trim, in order', async () => {
  const html = await fixture('styled-page.html');

  // The acceptance criterion of issue #13: materially smaller, and no text lost.
  assert.deepEqual(words(trimPageHtml(html)), words(html));
});

test('a heavily styled page trims to a fraction of its size', async () => {
  const html = await fixture('styled-page.html');
  const trimmed = trimPageHtml(html);

  assert.ok(
    trimmed.length < html.length / 2,
    `expected less than half of ${html.length} bytes, got ${trimmed.length}`,
  );
});

test('ids survive, because the PATCH write model targets elements by id', async () => {
  const trimmed = trimPageHtml(await fixture('styled-page.html'));

  assert.match(trimmed, /id="p:\{11111111-1111-1111-1111-111111111111\}\{12\}"/);
  assert.match(trimmed, /data-id="notes-block"/);
});

test('structure survives: headings, lists, tables, links, and their attributes', async () => {
  const trimmed = trimPageHtml(await fixture('styled-page.html'));

  assert.match(trimmed, /<h1[^>]*>/);
  assert.match(trimmed, /<ul>\s*<li/);
  assert.match(trimmed, /<table>/);
  assert.match(trimmed, /<a href="https:\/\/example\.com\/catalogue">/);
  assert.match(trimmed, /data-tag="to-do:completed"/);
  assert.equal((trimmed.match(/<td>/g) ?? []).length, 4, 'no cell may be dropped');
});

test('presentational attributes go, including the ones that are not `style`', async () => {
  const trimmed = trimPageHtml(await fixture('styled-page.html'));

  assert.doesNotMatch(trimmed, /font-family|font-size|color:|background-color|border|list-style/);
  assert.doesNotMatch(trimmed, /cellpadding|cellspacing|vertical-align|margin-top/);
});

test('layout coordinates survive, because ink is rendered into that same space', async () => {
  const trimmed = trimPageHtml(await fixture('styled-page.html'));

  // himetric -> px at 96 dpi in src/ink.ts puts strokes in this coordinate space. The
  // information cannot be recovered once it is thrown away.
  assert.match(trimmed, /style="position:absolute;left:48px;top:120px;width:624px"/);
  assert.match(trimmed, /<body data-absolute-enabled="true">/);
});

test('the InkNode placeholder comments are gone; the ink arrives as a PNG', async () => {
  const trimmed = trimPageHtml(await fixture('styled-page.html'));

  assert.doesNotMatch(trimmed, /InkNode/);
  assert.doesNotMatch(trimmed, /<!--/);
});

test('styling wrappers with nothing left to say are unwrapped or dropped', () => {
  assert.equal(
    trimPageHtml('<p id="a"><span style="font-weight:bold">text</span></p>'),
    '<p id="a">text</p>',
  );
  assert.equal(trimPageHtml('<div><span style="color:red"></span></div>'), '');
  assert.equal(trimPageHtml('<p id="a"><span style="color:red"></span></p>'), '<p id="a"></p>');
});

test('an empty table cell is kept, because dropping it shifts the column', () => {
  assert.equal(
    trimPageHtml('<table><tr><td></td><td style="width:2px">b</td></tr></table>'),
    '<table><tr><td></td><td style="width:2px">b</td></tr></table>',
  );
});

test('an image keeps its source and its size, and loses its border', () => {
  assert.equal(
    trimPageHtml('<img src="s" alt="a" width="10" height="20" style="border:1px solid #000" class="x" />'),
    '<img src="s" alt="a" width="10" height="20" />',
  );
});

test('text is copied through verbatim, so no entity is re-encoded', () => {
  assert.equal(
    trimPageHtml('<p id="a"><span style="color:red">&amp; &lt; &#160; caf&eacute;</span></p>'),
    '<p id="a">&amp; &lt; &#160; caf&eacute;</p>',
  );
});

test('damaged markup is trimmed rather than rejected', () => {
  // fast-xml-parser is not involved here, and neither is any strictness: a page whose
  // text is readable must not become a failed request over a stray tag.
  assert.equal(trimPageHtml('<p id="a">one</span> two'), '<p id="a">one two</p>');
  assert.equal(trimPageHtml('<p id="a">unclosed'), '<p id="a">unclosed</p>');
  assert.equal(trimPageHtml('a < b'), 'a < b');
});

test('an attribute value holding a quote or an angle bracket is not mistaken for a tag', () => {
  assert.equal(trimPageHtml(`<p id='a>b'>t</p>`), '<p id="a>b">t</p>');
  assert.equal(trimPageHtml(`<p id='say "hi"'>t</p>`), '<p id="say &quot;hi&quot;">t</p>');
});
