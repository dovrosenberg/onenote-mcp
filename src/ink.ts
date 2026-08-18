// The ink pipeline: InkML stroke data to an SVG to a PNG.
//
// Graph's normal page-content endpoint drops handwriting and leaves
// `<!-- InkNode is not supported -->` in its place, and Graph cannot export a page as an
// image or a PDF. The only way to see what was written is to fetch the raw strokes with
// `includeInkML=true` and draw them here. The PNG then goes to the calling model as an
// image, which reads the handwriting with its own vision — no OCR service is involved.
//
// This is a port of the validated recon script in Appendix A of project-spec.md. The
// three things that make it work are not obvious and must not be "simplified" away:
//
//   1. Namespace prefixes are stripped. Graph emits `inkml:ink`, `inkml:trace`, and so
//      on, and every lookup here uses the bare name.
//   2. Channel order comes from `<traceFormat>`, never from position. This account's
//      points are X, Y, F — F is pen pressure — so taking the first two numbers of each
//      point draws pressure as a coordinate.
//   3. Coordinates are himetric (1/100 mm) and are converted to px at 96 dpi. That is
//      the same coordinate space the page HTML positions typed content in, so ink and
//      typed content could later be registered against each other by arithmetic.
//
// Nothing here prints stroke data or any part of the source document in an error. The
// coordinates are the user's handwriting, and this repository's output can reach a
// public log.

import { Resvg } from '@resvg/resvg-js';
import { XMLParser } from 'fast-xml-parser';

/** himetric is 1/100 mm; the page HTML is px at 96 dpi. */
const HIMETRIC_PER_INCH = 2540;
const PX_PER_INCH = 96;

/** The raster width Appendix A used, and the width the fixture renders at. */
export const DEFAULT_RENDER_WIDTH = 1400;

/** The channel order assumed when a document declares no `<traceFormat>`. */
const DEFAULT_CHANNEL_ORDER: readonly string[] = ['X', 'Y'];

/** The units assumed when a channel declares none. */
const DEFAULT_UNITS = 'himetric';

/** How deep `<traceGroup>` elements may nest before the walk is abandoned. */
const MAX_TRACE_GROUP_DEPTH = 50;

/** Bounding-box padding: this fraction of each side, never less than MIN_MARGIN_PX. */
const MARGIN_FRACTION = 0.05;
const MIN_MARGIN_PX = 5;

/** A point on a stroke, in px at 96 dpi. */
export interface InkPoint {
  readonly x: number;
  readonly y: number;
}

/** One pen-down-to-pen-up stroke. Strokes with fewer than two points are dropped. */
export type InkStroke = readonly InkPoint[];

/** An SVG document plus the size its own attributes declare, in px. */
export interface InkSvg {
  readonly svg: string;
  readonly width: number;
  readonly height: number;
}

/** A rasterised image and its pixel dimensions. */
export interface RasterImage {
  readonly png: Uint8Array;
  readonly width: number;
  readonly height: number;
}

/** A rendered page of ink. `width` and `height` are the PNG's pixel dimensions. */
export interface InkImage extends RasterImage {
  readonly svg: string;
  readonly strokeCount: number;
}

/**
 * A document could not be walked to the end — it nests trace groups past the bound.
 *
 * A page with no readable ink is not this. That is an empty stroke list, and `null` from
 * renderInk, and it is the normal answer for a typed page.
 */
export class InkParseError extends Error {
  constructor(message: string, options: { cause?: unknown } = {}) {
    super(message);
    this.name = 'InkParseError';
    if ('cause' in options) this.cause = options.cause;
  }
}

/** Rasterisation failed. Separate from InkParseError: the strokes were fine. */
export class InkRenderError extends Error {
  constructor(message: string, options: { cause?: unknown } = {}) {
    super(message);
    this.name = 'InkRenderError';
    if ('cause' in options) this.cause = options.cause;
  }
}

/**
 * `removeNSPrefix` is the setting that makes every lookup below work: Graph returns
 * `inkml:ink`, `inkml:trace`, `inkml:traceFormat`. `parseTagValue: false` keeps a trace
 * body a string — a single-number trace would otherwise arrive as a number.
 */
const xmlParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  removeNSPrefix: true,
  textNodeName: '#text',
  parseTagValue: false,
});

/** Convert one himetric value to px at 96 dpi. */
export function himetricToPx(value: number): number {
  return (value * PX_PER_INCH) / HIMETRIC_PER_INCH;
}

/**
 * Every stroke in a document, from every `<ink>` root in it.
 *
 * The text does not have to be a bare InkML document: `<ink>` roots are located by
 * pattern and parsed one at a time, because Graph delivers them inside a multipart body
 * and a page can carry more than one root.
 *
 * A root that yields nothing is skipped rather than raised on. fast-xml-parser accepts
 * unclosed and mismatched tags without complaint, so a damaged root normally parses to
 * an empty object rather than throwing, and "damaged" cannot be told apart from "an
 * `<ink>` element that holds no traces". Reporting the difference would mean guessing,
 * and guessing wrong turns an ordinary typed page into a failed request.
 *
 * @returns the strokes found, which is an empty array when the text carries no ink.
 * @throws {InkParseError} if trace groups nest past MAX_TRACE_GROUP_DEPTH.
 */
export function parseInkStrokes(text: string): InkStroke[] {
  // Comments are stripped first. The HTML half of a page carries
  // `<!-- InkNode is not supported -->`, and any comment that mentions `<ink>` would
  // otherwise start a match that runs into the following real root and swallows it.
  const withoutComments = text.replace(/<!--[\s\S]*?-->/g, '');
  const roots = withoutComments.matchAll(/<(\w+:)?ink[\s>][\s\S]*?<\/(\w+:)?ink>/gi);

  const strokes: InkStroke[] = [];

  for (const root of roots) {
    let parsed: unknown;
    try {
      parsed = xmlParser.parse(root[0]);
    } catch {
      // The thrown error is not kept: fast-xml-parser puts the offending markup in its
      // message, and that markup is the user's handwriting.
      continue;
    }

    const record = asRecord(asRecord(parsed)?.['ink'] ?? firstValue(parsed));
    if (record === null) continue;

    strokes.push(...strokesFromInkRoot(record));
  }

  return strokes;
}

/**
 * Draw strokes as SVG paths on a white ground, sized to their bounding box.
 *
 * The box is padded so a stroke drawn at the very edge is not clipped by the stroke
 * width itself, and the padding has a floor so a single dot still produces a canvas.
 */
export function strokesToSvg(strokes: readonly InkStroke[]): InkSvg {
  if (strokes.length === 0) {
    throw new RangeError('strokesToSvg needs at least one stroke');
  }

  let minX = Infinity;
  let minY = Infinity;
  let maxX = -Infinity;
  let maxY = -Infinity;

  for (const stroke of strokes) {
    for (const point of stroke) {
      if (point.x < minX) minX = point.x;
      if (point.y < minY) minY = point.y;
      if (point.x > maxX) maxX = point.x;
      if (point.y > maxY) maxY = point.y;
    }
  }

  const margin = Math.max(
    (maxX - minX) * MARGIN_FRACTION,
    (maxY - minY) * MARGIN_FRACTION,
    MIN_MARGIN_PX,
  );
  const originX = minX - margin;
  const originY = minY - margin;
  const width = maxX - minX + margin * 2;
  const height = maxY - minY + margin * 2;

  const paths = strokes
    .map((stroke) => {
      const d = stroke
        .map((point, i) => `${i === 0 ? 'M' : 'L'}${point.x.toFixed(2)},${point.y.toFixed(2)}`)
        .join(' ');
      return `  <path d="${d}" fill="none" stroke="black" stroke-width="1" stroke-linecap="round" stroke-linejoin="round" />`;
    })
    .join('\n');

  // The width and height attributes are whole pixels because they are what the
  // rasteriser scales from; the viewBox keeps the real extent.
  const declaredWidth = Math.max(1, Math.round(width));
  const declaredHeight = Math.max(1, Math.round(height));

  const svg =
    `<svg xmlns="http://www.w3.org/2000/svg" viewBox="${originX.toFixed(2)} ${originY.toFixed(2)} ${width.toFixed(2)} ${height.toFixed(2)}" width="${declaredWidth}" height="${declaredHeight}">\n` +
    `  <rect x="${originX.toFixed(2)}" y="${originY.toFixed(2)}" width="${width.toFixed(2)}" height="${height.toFixed(2)}" fill="white" />\n` +
    `${paths}\n` +
    `</svg>\n`;

  return { svg, width: declaredWidth, height: declaredHeight };
}

/**
 * Rasterise an SVG to a PNG of exactly `width` px, height following the aspect ratio.
 *
 * @throws {InkRenderError} if resvg rejects the document.
 */
export function rasterizeSvg(svg: string, width: number = DEFAULT_RENDER_WIDTH): RasterImage {
  if (!Number.isInteger(width) || width < 1) {
    throw new RangeError(`width must be a positive integer, got ${String(width)}`);
  }

  try {
    const rendered = new Resvg(svg, { fitTo: { mode: 'width', value: width } }).render();
    return { png: rendered.asPng(), width: rendered.width, height: rendered.height };
  } catch (err) {
    throw new InkRenderError(`Rasterising the reconstructed ink failed: ${describe(err)}`, {
      cause: err,
    });
  }
}

/**
 * The whole pipeline: document text to a PNG, or null when the text carries no ink.
 *
 * Null is the normal answer for a page that was typed rather than written. Only a page
 * whose ink is present and unreadable raises.
 *
 * @throws {InkParseError} if `<ink>` is present but yields no strokes.
 * @throws {InkRenderError} if rasterisation fails.
 */
export function renderInk(text: string, width: number = DEFAULT_RENDER_WIDTH): InkImage | null {
  const strokes = parseInkStrokes(text);
  if (strokes.length === 0) return null;

  const { svg } = strokesToSvg(strokes);
  return { ...rasterizeSvg(svg, width), svg, strokeCount: strokes.length };
}

/** Strokes from one already-parsed `<ink>` element. */
function strokesFromInkRoot(inkRoot: Record<string, unknown>): InkStroke[] {
  const { order, units } = readChannels(inkRoot);

  // A document that declares channels but not X or Y is treated as positional, which is
  // what Appendix A's default order amounts to.
  const xIndex = order.indexOf('X') === -1 ? 0 : order.indexOf('X');
  const yIndex = order.indexOf('Y') === -1 ? 1 : order.indexOf('Y');
  const toPx = units === DEFAULT_UNITS ? himetricToPx : (value: number): number => value;

  const strokes: InkStroke[] = [];

  for (const trace of collectTraces(inkRoot, 0)) {
    const raw = typeof trace === 'string' ? trace : asRecord(trace)?.['#text'];
    if (typeof raw !== 'string' || raw.trim() === '') continue;

    const points: InkPoint[] = [];
    for (const pointText of raw.trim().split(',')) {
      const numbers = pointText.trim().split(/\s+/).map(Number);
      // Both bounds and NaN are checked here: a point short of the declared channel
      // count yields `undefined`, which Number.isNaN would let through as a coordinate.
      const x = numbers[xIndex] ?? NaN;
      const y = numbers[yIndex] ?? NaN;
      if (!Number.isFinite(x) || !Number.isFinite(y)) continue;
      points.push({ x: toPx(x), y: toPx(y) });
    }

    // A one-point stroke has no path to draw; Appendix A dropped it and so does this.
    if (points.length > 1) strokes.push(points);
  }

  return strokes;
}

/**
 * Channel order and units from `<definitions><context><inkSource><traceFormat>`.
 *
 * Every context is examined rather than only the first, because a document with more
 * than one ink source declares more than one context and the one carrying the trace
 * format need not come first. Order comes from the first context that declares one.
 */
function readChannels(inkRoot: Record<string, unknown>): {
  order: readonly string[];
  units: string;
} {
  const definitions = asRecord(inkRoot['definitions']);
  for (const context of toArray(definitions?.['context'])) {
    const traceFormat = asRecord(asRecord(asRecord(context)?.['inkSource'])?.['traceFormat']);
    const channels = toArray(traceFormat?.['channel'])
      .map((channel) => asRecord(channel)?.['@_name'])
      .filter((name): name is string => typeof name === 'string' && name !== '');

    if (channels.length === 0) continue;

    const xChannel = toArray(traceFormat?.['channel']).find(
      (channel) => asRecord(channel)?.['@_name'] === 'X',
    );
    const units = asRecord(xChannel)?.['@_units'];

    return { order: channels, units: typeof units === 'string' ? units : DEFAULT_UNITS };
  }

  return { order: DEFAULT_CHANNEL_ORDER, units: DEFAULT_UNITS };
}

/**
 * Every `<trace>` under a node, including those inside nested `<traceGroup>` elements.
 *
 * @throws {InkParseError} past MAX_TRACE_GROUP_DEPTH levels of nesting, which no real
 * document produces; the bound exists so a pathological one ends as a named error.
 */
function collectTraces(node: Record<string, unknown>, depth: number): unknown[] {
  if (depth > MAX_TRACE_GROUP_DEPTH) {
    throw new InkParseError(
      `Ink trace groups nested deeper than ${MAX_TRACE_GROUP_DEPTH} levels, so the walk was abandoned.`,
    );
  }

  const traces = [...toArray(node['trace'])];
  for (const group of toArray(node['traceGroup'])) {
    const record = asRecord(group);
    if (record !== null) traces.push(...collectTraces(record, depth + 1));
  }
  return traces;
}

/** fast-xml-parser gives a single child as a value and repeats as an array. */
function toArray(value: unknown): unknown[] {
  if (value === undefined || value === null) return [];
  return Array.isArray(value) ? value : [value];
}

function asRecord(value: unknown): Record<string, unknown> | null {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) return null;
  return value as Record<string, unknown>;
}

/** The first child of a parsed document, for a root whose name is not `ink`. */
function firstValue(parsed: unknown): unknown {
  const record = asRecord(parsed);
  return record === null ? null : Object.values(record)[0];
}

function describe(err: unknown): string {
  return err instanceof Error ? `${err.name}: ${err.message}` : String(err);
}
