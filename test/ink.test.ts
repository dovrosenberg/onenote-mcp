import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';

import {
  DEFAULT_RENDER_WIDTH,
  InkParseError,
  InkRenderError,
  himetricToPx,
  parseInkStrokes,
  rasterizeSvg,
  renderInk,
  strokesToSvg,
  type InkStroke,
} from '../src/ink.ts';

const FIXTURES = path.join(import.meta.dirname, 'fixtures');

function fixture(name: string): Promise<string> {
  return readFile(path.join(FIXTURES, name), 'utf8');
}

/** Width and height out of a PNG's IHDR, so the assertion is about the file itself. */
function pngSize(png: Uint8Array): { width: number; height: number } {
  const view = new DataView(png.buffer, png.byteOffset, png.byteLength);
  assert.deepEqual(
    [...png.slice(0, 8)],
    [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a],
    'expected a PNG signature',
  );
  return { width: view.getUint32(16), height: view.getUint32(20) };
}

function extent(strokes: readonly InkStroke[]): {
  minX: number;
  minY: number;
  maxX: number;
  maxY: number;
} {
  const xs = strokes.flatMap((stroke) => stroke.map((point) => point.x));
  const ys = strokes.flatMap((stroke) => stroke.map((point) => point.y));
  return {
    minX: Math.min(...xs),
    minY: Math.min(...ys),
    maxX: Math.max(...xs),
    maxY: Math.max(...ys),
  };
}

test('himetric converts to px at 96 dpi', () => {
  assert.equal(himetricToPx(2540), 96); // 2540 himetric = 1 inch
  assert.equal(himetricToPx(0), 0);
  assert.equal(himetricToPx(1270), 48);
});

test('text with no ink yields no strokes and renders to null', () => {
  // The normal answer for a page that was typed rather than written.
  assert.deepEqual(parseInkStrokes('<html><body><p>no ink here</p></body></html>'), []);
  assert.equal(renderInk('<html><body><p>no ink here</p></body></html>'), null);
  assert.equal(renderInk(''), null);
});

test('the X,Y,F fixture parses as X and Y, never as pressure', async () => {
  const strokes = parseInkStrokes(await fixture('xyf-himetric.inkml'));

  // Two traces of three points each, in px: a 96px square drawn as two open corners.
  assert.equal(strokes.length, 2);
  assert.deepEqual(strokes[0], [
    { x: 0, y: 0 },
    { x: 96, y: 0 },
    { x: 96, y: 96 },
  ]);
  assert.deepEqual(strokes[1], [
    { x: 96, y: 96 },
    { x: 0, y: 96 },
    { x: 0, y: 0 },
  ]);

  // The F values in the fixture are 8000-9200. Reading the first two numbers of each
  // point instead of the declared channels would put pressure on the Y axis, and the
  // extent would run to hundreds of px rather than 96.
  assert.deepEqual(extent(strokes), { minX: 0, minY: 0, maxX: 96, maxY: 96 });
});

test('traces are collected from every ink root and every nesting level', async () => {
  const strokes = parseInkStrokes(await fixture('two-roots-nested.inkml'));

  // Two loose traces plus three at three depths in the first root, one in the second.
  assert.equal(strokes.length, 6);

  // Each trace starts 100 himetric further along X than the last, so the set of first
  // points identifies which traces were found. A walk that stopped at the first root's
  // direct children would return only the first two.
  assert.deepEqual(
    strokes.map((stroke) => Math.round(stroke[0]?.x ?? NaN)),
    [0, 4, 8, 11, 15, 19],
  );
});

test('a document with no traceFormat falls back to positional X,Y', () => {
  const strokes = parseInkStrokes(
    '<ink xmlns="http://www.w3.org/2003/InkML"><trace>0 0, 2540 2540</trace></ink>',
  );

  assert.deepEqual(strokes, [
    [
      { x: 0, y: 0 },
      { x: 96, y: 96 },
    ],
  ]);
});

test('non-himetric units are taken as px and left alone', () => {
  const strokes = parseInkStrokes(
    [
      '<ink xmlns="http://www.w3.org/2003/InkML"><definitions><context><inkSource><traceFormat>',
      '<channel name="X" units="px" /><channel name="Y" units="px" />',
      '</traceFormat></inkSource></definitions></context><trace>0 0, 10 20</trace></ink>',
    ].join(''),
  );

  assert.deepEqual(strokes, [
    [
      { x: 0, y: 0 },
      { x: 10, y: 20 },
    ],
  ]);
});

test('single-point traces, empty traces, and unparsable points are dropped', () => {
  const strokes = parseInkStrokes(
    [
      '<ink xmlns="http://www.w3.org/2003/InkML">',
      '<trace>0 0</trace>', // one point is not a path
      '<trace></trace>',
      '<trace>0 0, 2540 2540, x y, 2540</trace>', // a bad point and a short point
      '</ink>',
    ].join(''),
  );

  assert.equal(strokes.length, 1);
  assert.deepEqual(strokes[0], [
    { x: 0, y: 0 },
    { x: 96, y: 96 },
  ]);
});

test('an ink element with no usable traces reads as no ink, not as a failure', () => {
  // fast-xml-parser accepts unclosed and mismatched tags, so a damaged root parses to
  // an empty object rather than throwing. Nothing here can tell that apart from an
  // <ink> element that simply holds no traces, and the issue's rule is that no parsable
  // ink is normal.
  assert.deepEqual(parseInkStrokes('<ink><definitions><unclosed></ink>'), []);
  assert.equal(renderInk('<ink />'), null);
});

test('a comment mentioning ink does not swallow the root that follows it', () => {
  // The HTML half of a page carries `<!-- InkNode is not supported -->`, and the
  // fixtures explain themselves in comments. A match that started inside a comment
  // would run to the first real closing tag and take a whole root with it.
  const strokes = parseInkStrokes(
    '<!-- an <ink> root follows --><ink><trace>0 0, 2540 2540</trace></ink>',
  );

  assert.equal(strokes.length, 1);
});

test('trace groups nested past the bound end as a named error', () => {
  const depth = 60;
  const nested = `${'<traceGroup>'.repeat(depth)}<trace>0 0, 100 100</trace>${'</traceGroup>'.repeat(depth)}`;

  assert.throws(
    () => parseInkStrokes(`<ink>${nested}</ink>`),
    (err: unknown) => err instanceof InkParseError && /nested deeper/.test(err.message),
  );
});

test('the SVG is sized to the stroke bounding box plus a margin', () => {
  const strokes: InkStroke[] = [
    [
      { x: 0, y: 0 },
      { x: 96, y: 96 },
    ],
  ];
  const { svg, width, height } = strokesToSvg(strokes);

  // margin = max(96 * 0.05, 96 * 0.05, 5) = 5, so 96 + 2 * 5 on each axis.
  assert.equal(width, 106);
  assert.equal(height, 106);
  assert.match(svg, /viewBox="-5\.00 -5\.00 106\.00 106\.00"/);
  assert.match(svg, /<rect [^>]*fill="white"/);
  assert.match(svg, /<path d="M0\.00,0\.00 L96\.00,96\.00"/);
});

test('a bounding box with no extent still gets a canvas', () => {
  // Two points at the same place: the margin floor is what stops a zero-width SVG.
  const { width, height } = strokesToSvg([
    [
      { x: 3, y: 3 },
      { x: 3, y: 3 },
    ],
  ]);

  assert.equal(width, 10);
  assert.equal(height, 10);
});

test('strokesToSvg refuses an empty stroke list', () => {
  assert.throws(() => strokesToSvg([]), RangeError);
});

test('the X,Y,F fixture renders to a PNG of the expected size', async () => {
  const image = renderInk(await fixture('xyf-himetric.inkml'));

  assert.ok(image !== null);
  assert.equal(image.strokeCount, 2);
  // The SVG is 106 x 106, so a raster fitted to 1400 wide is 1400 square.
  assert.equal(image.width, DEFAULT_RENDER_WIDTH);
  assert.equal(image.height, DEFAULT_RENDER_WIDTH);
  assert.deepEqual(pngSize(image.png), { width: 1400, height: 1400 });
});

test('the render width is honoured and the aspect ratio follows the strokes', () => {
  const image = renderInk(
    '<ink><trace>0 0, 20320 2540</trace></ink>', // 768 x 96 px before the margin
    400,
  );

  assert.ok(image !== null);
  // 768 + 2 * 38.4 = 844.8 -> 845 wide, 96 + 2 * 38.4 = 172.8 -> 173 high.
  assert.equal(pngSize(image.png).width, 400);
  assert.equal(pngSize(image.png).height, Math.round((400 * 173) / 845));
});

test('rasterizeSvg rejects a width that is not a positive integer', () => {
  const { svg } = strokesToSvg([
    [
      { x: 0, y: 0 },
      { x: 1, y: 1 },
    ],
  ]);

  assert.throws(() => rasterizeSvg(svg, 0), RangeError);
  assert.throws(() => rasterizeSvg(svg, 1.5), RangeError);
});

test('a document resvg cannot render raises InkRenderError, not a raw resvg error', () => {
  assert.throws(() => rasterizeSvg('<svg><this is not markup'), InkRenderError);
});
