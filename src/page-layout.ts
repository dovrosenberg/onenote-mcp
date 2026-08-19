// Where an append lands relative to the handwriting already on the page.
//
// A OneNote page positions its text in absolutely positioned top-level divs — outlines —
// and keeps its ink in a separate layer of the same coordinate space. An outline grows
// downwards as content is added to it, so appending to a page whose strokes sit below the
// text runs the text over the handwriting. Nothing is lost, but the page becomes
// unreadable at exactly the place someone wrote by hand.
//
// The obvious fix — put the new content in its own outline, positioned below the ink — is
// not available. Measured against the live service on 2026-08-19, and matching the tables
// in `api-overview.md`:
//
//   - `target: "body"` is the first top-level div, not the `<body>` element, so anything
//     appended is a child of that outline rather than a child of the body.
//   - An absolutely positioned `<div>` sent as content is flattened into the outline and
//     its `position`, `top` and `left` are discarded. That holds on a page whose body
//     really is `data-absolute-enabled="true"` with several outlines already on it.
//   - `insert` beside an outline is 400 code 20135, and `replace` on one is 20134 by
//     generated id or 20141 by `data-id`. So no outline can be created, moved or rewritten.
//   - Vertical space cannot be faked with spacing either: `margin-top:278pt` came back as
//     `5.5pt` and an empty `<p>&nbsp;</p>` was deleted outright.
//   - `<br>` survives verbatim — eight sent, eight read back.
//   - A `data-id` survives only on an element that carries content. Sent in one PATCH:
//     `<p data-id="m1-text">text</p>` and `<div data-id="m3-div"><p>text</p></div>` both
//     came back with their ids; `<p data-id="m2-breaks"><br /><br /></p>` and the same
//     wrapped in a div came back with the attribute gone. So the marker that records what
//     was cleared cannot ride on the padding itself — it goes on the caller's content.
//
// So the only lever is line breaks inside the outline being appended to, and this module
// is the arithmetic for how many. It is deliberately pure: it takes the page HTML and the
// ink's bounding box and returns a plan, so the whole decision is testable without a
// Graph call.

import type { InkBox } from './ink.ts';

/**
 * The height one `<br>` adds, in px.
 *
 * Chosen, not measured: Graph reports an outline's `left`, `top` and `width` but never its
 * height or its line height, so nothing in the API can confirm this. 19px is a line of the
 * default 11pt Calibri at 96 dpi. Being wrong here makes the gap above the appended text
 * too large or too small; it cannot put the text back on top of the ink unless it is out
 * by more than the margin below.
 */
export const LINE_HEIGHT_PX = 19;

/** How far below the lowest stroke the appended text starts. */
export const INK_CLEARANCE_MARGIN_PX = 24;

/** The width Graph gives an outline when the input HTML names none. */
export const DEFAULT_OUTLINE_WIDTH_PX = 624;

/**
 * The most breaks one append may add.
 *
 * A page is not unbounded, and a stroke at an absurd coordinate — or a coordinate this
 * repository misread — must not produce a thousand blank lines that a person then has to
 * delete by hand. 120 breaks is about 2280px, well past the bottom of a printed page.
 */
export const MAX_CLEARANCE_BREAKS = 120;

/** The `data-id` prefix that marks content this server pushed down, and how far. */
export const CLEARANCE_ID_PREFIX = 'ink-clearance-';

/** One top-level div: an outline, in the page's own px coordinate space. */
export interface Outline {
  /** The generated id, or the empty string when the page was read without `includeIDs`. */
  readonly id: string;
  /** `_default` on a page this server created, absent on one the client authored. */
  readonly dataId: string | null;
  readonly left: number;
  readonly top: number;
  /** Null when the div declares none; Graph's own default applies then. */
  readonly width: number | null;
}

/** What an append has to do about the ink, and why. */
export interface ClearancePlan {
  /** How many `<br>` elements to put in front of the caller's fragment. */
  readonly breaks: number;
  /** The lowest stroke in the outline's column, in px. Reported, not stored. */
  readonly inkBottom: number;
  /** Where the text was already known to reach: a previous marker, or the outline top. */
  readonly clearedFrom: number;
  /**
   * The page position this padding brings the content down to, in px. This is what the
   * marker records — where the text now starts, not where the ink ends. Recording the ink
   * bottom instead makes the next append pad again over rounding: a stroke measured at
   * 466.02px against a marker of 466 reads as 0.02px of new ink.
   */
  readonly clearedTo: number;
  /** True when `breaks` hit MAX_CLEARANCE_BREAKS and the padding is short of the ink. */
  readonly truncated: boolean;
}

/** Elements that never open a level, so a walk that counted them would nest forever. */
const VOID_ELEMENTS = new Set([
  'area',
  'base',
  'br',
  'col',
  'embed',
  'hr',
  'img',
  'input',
  'link',
  'meta',
  'source',
  'track',
  'wbr',
]);

/** A tag or a comment. Comments matter only in that they must not be read as tags. */
const TAG_OR_COMMENT = /<!--[\s\S]*?-->|<\/?([a-zA-Z][^\s/>]*)((?:"[^"]*"|'[^']*'|[^>"'])*)>/g;

/**
 * The page's outlines, in document order. The first is what `target: "body"` addresses.
 *
 * The scan is tolerant in the same way ./page-html.ts is: an unmatched close tag is
 * ignored and an unclosed element is closed by its parent. A page whose markup is untidy
 * still has to be appendable — refusing to write because a stray tag confused a parser
 * would be a worse failure than an imperfect guess about the layout.
 */
export function parseOutlines(html: string): Outline[] {
  const body = /<body\b[^>]*>([\s\S]*?)<\/body>/i.exec(html)?.[1] ?? html;
  const outlines: Outline[] = [];
  // A stack of open element names rather than a depth count: an unclosed `<p>` inside an
  // outline must not push everything after it a level deeper, which would hide the next
  // outline entirely. A close tag pops to the matching name, and one that matches nothing
  // open is ignored.
  const open: string[] = [];

  for (const match of body.matchAll(TAG_OR_COMMENT)) {
    const tag = match[0];
    if (tag.startsWith('<!--')) continue;

    const name = (match[1] ?? '').toLowerCase();
    if (tag.startsWith('</')) {
      const at = open.lastIndexOf(name);
      if (at !== -1) open.length = at;
      continue;
    }

    const attributes = match[2] ?? '';
    if (open.length === 0 && name === 'div') {
      const style = attributeValue(attributes, 'style') ?? '';
      outlines.push({
        id: attributeValue(attributes, 'id') ?? '',
        dataId: attributeValue(attributes, 'data-id'),
        left: pixels(style, 'left') ?? 0,
        top: pixels(style, 'top') ?? 0,
        width: pixels(style, 'width'),
      });
    }

    if (!VOID_ELEMENTS.has(name) && !tag.endsWith('/>')) open.push(name);
  }

  return outlines;
}

/**
 * How far down the page a previous append already cleared, from the markers on it.
 *
 * Without this every append would pad again from the outline's top, and a page written to
 * three times would carry three stacks of blank lines. A later append pads only for ink
 * that reaches past the marker — and not at all when the handwriting has not changed.
 */
export function clearedTo(html: string): number | null {
  let deepest: number | null = null;

  for (const match of html.matchAll(/data-id="ink-clearance-(\d+)"/g)) {
    const value = Number(match[1]);
    if (Number.isFinite(value) && (deepest === null || value > deepest)) deepest = value;
  }

  return deepest;
}

/**
 * How much padding an append to `outline` needs, or null when it needs none.
 *
 * Null is the normal answer: a page with no ink, ink that sits beside the outline's
 * column rather than in it, or ink that a previous append already cleared.
 */
export function planInkClearance(
  outline: Outline,
  ink: InkBox | null,
  alreadyCleared: number | null,
): ClearancePlan | null {
  if (ink === null) return null;

  // Ink in another column is not in the way. An outline's width is the column: text wraps
  // inside it and cannot spill sideways into the strokes.
  const right = outline.left + (outline.width ?? DEFAULT_OUTLINE_WIDTH_PX);
  if (ink.right <= outline.left || ink.left >= right) return null;

  // With no marker the only thing known about the text is where the outline starts, so
  // the padding is measured from there. That clears the ink whatever the text height is,
  // at the cost of a gap the size of the text already in the outline.
  const clearedFrom = alreadyCleared ?? outline.top;
  if (ink.bottom <= clearedFrom) return null;

  const wanted = Math.ceil((ink.bottom + INK_CLEARANCE_MARGIN_PX - clearedFrom) / LINE_HEIGHT_PX);
  const breaks = Math.min(wanted, MAX_CLEARANCE_BREAKS);

  return {
    breaks,
    inkBottom: ink.bottom,
    clearedFrom,
    // Where the padding actually reaches, which is short of the ink when it was capped.
    clearedTo: Math.round(
      Math.min(ink.bottom + INK_CLEARANCE_MARGIN_PX, clearedFrom + breaks * LINE_HEIGHT_PX),
    ),
    truncated: breaks < wanted,
  };
}

/**
 * The content to send: the padding, then the caller's fragment carrying the marker.
 *
 * The marker is a `data-id` rather than a comment because a comment does not survive the
 * round trip and `data-id` does. It wraps the fragment rather than sitting on the padding
 * because the service discards the attribute from an element that holds nothing but line
 * breaks — measured, see the note at the top of this file. Wrapping the fragment also
 * puts the marker exactly where the meaning is: this content is the content that was
 * moved below the ink.
 */
export function clearanceHtml(plan: ClearancePlan, fragment: string): string {
  const id = `${CLEARANCE_ID_PREFIX}${plan.clearedTo}`;
  return `<p>${'<br />'.repeat(plan.breaks)}</p><div data-id="${id}">${fragment}</div>`;
}

/** One attribute out of a start tag, single or double quoted. */
function attributeValue(attributes: string, name: string): string | null {
  const pattern = new RegExp(`\\b${name}\\s*=\\s*("([^"]*)"|'([^']*)')`, 'i');
  const match = pattern.exec(attributes);
  if (match === null) return null;
  return match[2] ?? match[3] ?? null;
}

/** `left:48px` out of a style attribute, as a number. Graph writes px and nothing else. */
function pixels(style: string, property: string): number | null {
  const match = new RegExp(`(?:^|;)\\s*${property}\\s*:\\s*(-?[0-9.]+)px`, 'i').exec(style);
  if (match === null) return null;
  const value = Number(match[1]);
  return Number.isFinite(value) ? value : null;
}
