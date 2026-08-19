// The arithmetic that keeps an append off the handwriting.
//
// Everything here is pure, so the whole decision is testable without a Graph call. What
// it cannot test is the one number that was chosen rather than measured: LINE_HEIGHT_PX.
// No endpoint reports an outline's rendered height or its line height, so whether 19px
// per `<br>` puts the text just below the strokes or well below them is only visible by
// looking at the page in OneNote.
//
// The HTML here is written in the shape Graph emits — absolutely positioned top-level
// divs carrying `left`, `top` and `width` in px — because that shape is the whole input
// to `parseOutlines`.

import test from 'node:test';
import assert from 'node:assert/strict';

import type { InkBox } from '../src/ink.ts';
import {
  CLEARANCE_ID_PREFIX,
  DEFAULT_OUTLINE_WIDTH_PX,
  INK_CLEARANCE_MARGIN_PX,
  LINE_HEIGHT_PX,
  MAX_CLEARANCE_BREAKS,
  clearanceHtml,
  clearedTo,
  estimateContentHeight,
  parseOutlines,
  planInkClearance,
  type Outline,
} from '../src/page-layout.ts';

const PAGE = `<html><head><title>t</title></head>
<body data-absolute-enabled="true" style="font-family:Calibri;font-size:11pt">
  <div id="div:{aaa}{32}" data-id="_default" style="position:absolute;left:48px;top:120px;width:624px">
    <h1>Heading</h1>
    <p>Typed text.</p>
    <div data-id="kept-nested">
      <p>A nested div survives when it carries a data-id, and is not an outline.</p>
    </div>
  </div>
  <!-- InkNode is not supported -->
  <div id="div:{bbb}{113}" style="position:absolute;left:723px;top:143px;width:624px">
    <p>A second outline, off to the right.</p>
  </div>
</body></html>`;

const OUTLINE_CONTENT =
  '\n    <h1>Heading</h1>\n    <p>Typed text.</p>\n    <div data-id="kept-nested">\n' +
  '      <p>A nested div survives when it carries a data-id, and is not an outline.</p>\n' +
  '    </div>\n  ';

const OUTLINE: Outline = {
  id: 'div:{aaa}{32}',
  dataId: '_default',
  left: 48,
  top: 120,
  width: 624,
  content: OUTLINE_CONTENT,
};

/** The same outline with nothing in it, for the cases about geometry rather than height. */
const EMPTY: Outline = { ...OUTLINE, content: '' };

function box(left: number, top: number, right: number, bottom: number): InkBox {
  return { left, top, right, bottom };
}

test('parseOutlines returns the top-level divs, in order, with their geometry', () => {
  const outlines = parseOutlines(PAGE);

  assert.deepEqual(outlines, [
    OUTLINE,
    {
      id: 'div:{bbb}{113}',
      dataId: null,
      left: 723,
      top: 143,
      width: 624,
      content: '\n    <p>A second outline, off to the right.</p>\n  ',
    },
  ]);
});

test('a nested div is not an outline, however it is attributed', () => {
  // `target: "body"` addresses the first *top-level* div. A walk that counted the nested
  // one would target a div the service does not treat as an outline at all.
  assert.equal(
    parseOutlines(PAGE).some((outline) => outline.dataId === 'kept-nested'),
    false,
  );
});

test('a page with no body element is scanned whole', () => {
  const outlines = parseOutlines('<div style="position:absolute;left:10px;top:20px"><p>x</p></div>');
  assert.deepEqual(outlines, [
    { id: '', dataId: null, left: 10, top: 20, width: null, content: '<p>x</p>' },
  ]);
});

test('void and self-closing tags do not swallow the outlines after them', () => {
  const html =
    '<body><div style="left:48px;top:120px"><p>one<br>two<img src="x"></p></div>' +
    '<div style="left:48px;top:400px"><p>two</p></div></body>';
  assert.equal(parseOutlines(html).length, 2);
});

test('an unclosed element does not hide the rest of the page', () => {
  // Refusing to write because a stray tag confused a parser is a worse failure than an
  // imperfect guess at the layout, so the scan has to keep going.
  const html =
    '<body><div style="left:48px;top:120px"><p>unclosed</div>' +
    '<div style="left:48px;top:400px"><p>still found</p></div></body>';
  assert.equal(parseOutlines(html).length, 2);
});

test('no ink means no padding', () => {
  assert.equal(planInkClearance(EMPTY, null, null), null);
});

test('ink beside the outline column is not in the way', () => {
  // The second outline on the fixture starts at 723px. Strokes over there cannot be
  // reached by text that wraps inside a 624px column starting at 48px.
  assert.equal(planInkClearance(EMPTY, box(700, 30, 900, 500), null), null);
});

test('ink above the outline is not in the way', () => {
  assert.equal(planInkClearance(EMPTY, box(50, 10, 500, 100), null), null);
});

test('ink hanging below the outline top is padded from that top', () => {
  const plan = planInkClearance(EMPTY, box(45, 30, 533, 466), null);

  assert.ok(plan !== null);
  assert.equal(plan.clearedFrom, 120);
  assert.equal(plan.inkBottom, 466);
  assert.equal(plan.truncated, false);
  // Measured from the outline's top because nothing reports where its text ends: the
  // padding clears the ink whatever the text height turns out to be.
  assert.equal(plan.breaks, Math.ceil((466 + INK_CLEARANCE_MARGIN_PX - 120) / LINE_HEIGHT_PX));
});

test('an outline with no declared width is assumed to be Graph’s default width', () => {
  const noWidth: Outline = { ...EMPTY, width: null };
  const justInside = box(OUTLINE.left + DEFAULT_OUTLINE_WIDTH_PX - 1, 30, 900, 466);

  assert.notEqual(planInkClearance(noWidth, justInside, null), null);
  assert.equal(planInkClearance(noWidth, box(OUTLINE.left + DEFAULT_OUTLINE_WIDTH_PX, 30, 900, 466), null), null);
});

test('a previous clearance is not repeated', () => {
  // Without this, every append would pad again from the outline's top and a page written
  // to three times would carry three stacks of blank lines. The marker is where the text
  // was put, so ink ending just above it needs nothing.
  assert.equal(planInkClearance(EMPTY, box(45, 30, 533, 466.02), 490), null);
});

test('ink added below an earlier clearance is padded only for the difference', () => {
  const plan = planInkClearance(EMPTY, box(45, 30, 533, 600), 490);

  assert.ok(plan !== null);
  assert.equal(plan.clearedFrom, 490);
  assert.equal(plan.breaks, Math.ceil((600 + INK_CLEARANCE_MARGIN_PX - 490) / LINE_HEIGHT_PX));
  assert.equal(plan.clearedTo, 600 + INK_CLEARANCE_MARGIN_PX);
});

test('padding is bounded, and says when it stopped short', () => {
  const plan = planInkClearance(EMPTY, box(45, 30, 533, 100_000), null);

  assert.ok(plan !== null);
  assert.equal(plan.breaks, MAX_CLEARANCE_BREAKS);
  assert.equal(plan.truncated, true);
  // The marker records where the padding actually reached, not the ink it fell short of,
  // so the next append picks up from there rather than repeating the whole 120 lines.
  assert.equal(plan.clearedTo, OUTLINE.top + MAX_CLEARANCE_BREAKS * LINE_HEIGHT_PX);
});

test('clearedTo reads the deepest marker on the page', () => {
  const html = `<div data-id="${CLEARANCE_ID_PREFIX}490">a</div><div data-id="${CLEARANCE_ID_PREFIX}600">b</div>`;
  assert.equal(clearedTo(html), 600);
  assert.equal(clearedTo('<p>no marker here</p>'), null);
});

test('the padding is breaks, then the fragment wrapped in the marker', () => {
  const plan = planInkClearance(EMPTY, box(45, 30, 533, 466), null);
  assert.ok(plan !== null);

  const html = clearanceHtml(plan, '<p>appended</p>');

  assert.equal((html.match(/<br \/>/g) ?? []).length, plan.breaks);
  assert.equal(clearedTo(html), plan.clearedTo);
  assert.match(html, /<div data-id="ink-clearance-478"><p>appended<\/p><\/div>$/);
});

test('the marker is on the content, not on the padding', () => {
  // Measured on the live service: the service discards a data-id from an element that
  // holds only line breaks, so a marker sitting on the padding is gone by the next read
  // and the following append pads all over again.
  const plan = planInkClearance(EMPTY, box(45, 30, 533, 466), null);
  assert.ok(plan !== null);

  const padding = clearanceHtml(plan, '<p>appended</p>').split('</p>')[0] ?? '';
  assert.equal(padding.includes(CLEARANCE_ID_PREFIX), false);
});

test('the text already in the outline is estimated, so the gap is not the whole outline', () => {
  // Measuring from the outline's top is what made the first live run leave most of a
  // screen of blank space above the appended text. The estimate takes the content off
  // that distance; it is deliberately low, so what is left is a gap rather than an
  // overlap.
  const withText = planInkClearance(OUTLINE, box(45, 30, 533, 466), null);
  const withoutText = planInkClearance(EMPTY, box(45, 30, 533, 466), null);

  assert.ok(withText !== null && withoutText !== null);
  assert.ok(
    withText.breaks < withoutText.breaks,
    `${withText.breaks} breaks over content should be fewer than ${withoutText.breaks} over an empty outline`,
  );
  assert.equal(withText.clearedFrom, 120 + estimateContentHeight(OUTLINE_CONTENT));
});

test('estimateContentHeight counts blocks and line breaks, and nothing else', () => {
  assert.equal(estimateContentHeight(''), 0);
  assert.equal(estimateContentHeight('<p>one</p><p>two</p>'), 32);
  assert.equal(estimateContentHeight('<p><br /><br /></p>'), 48);
  // Spans, links and the styling wrappers Graph adds are not lines.
  assert.equal(estimateContentHeight('<p><span style="x">t</span><a href="#">l</a></p>'), 16);
  // A nested div is a wrapper, not a line; its paragraphs are the lines.
  assert.equal(estimateContentHeight('<div data-id="k"><p>a</p></div>'), 16);
});

test('a page already padded once is not padded again by the estimate alone', () => {
  // Twenty breaks and a paragraph, on an outline at 120px: the estimate alone puts the
  // text past the ink, so the second append adds nothing even before the marker is read.
  const padded: Outline = {
    ...OUTLINE,
    content: `<p>${'<br />'.repeat(20)}</p><p>already below the ink</p>`,
  };

  assert.equal(planInkClearance(padded, box(45, 30, 533, 466), null), null);
});

test('an image counts as its declared height, not as nothing', () => {
  // The case the block table gets worst. Graph returns width and height on an img in the
  // output HTML, and an image is hundreds of px where a paragraph is sixteen; counting it
  // as nothing would put the padding a screen too long.
  assert.equal(estimateContentHeight('<img width="400" height="248" src="x" />'), 248);
  // No height attribute: an assumption, and a low one.
  assert.equal(estimateContentHeight('<img src="x" />'), 100);
});
