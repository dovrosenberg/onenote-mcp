// Trimming the HTML half of a page-content response down to readable structure.
//
// Graph returns page markup that is mostly presentation. Every text run is wrapped in a
// `<span style="font-family:Calibri;font-size:11pt;color:#252423">`, every paragraph
// carries margin declarations, and the handwriting is a `<!-- InkNode is not supported -->`
// comment. A model reading the page pays for all of it and learns nothing from it.
//
// Four things here are deliberate:
//
//   1. `id` and `data-id` attributes survive. The OneNote write model is a PATCH whose
//      actions target an element by id, so stripping them would break issue #18.
//   2. Position declarations survive. Page content is laid out absolutely in px at 96 dpi
//      under `<body data-absolute-enabled="true">`, which is the same coordinate space
//      ./ink.ts renders strokes into. Nothing needs that today; a tool registering ink
//      against typed content would, and the information cannot be recovered later.
//   3. Text nodes are copied out verbatim, never re-encoded. The trimmer only ever
//      rewrites tags, so no entity and no character of the user's text can be lost to a
//      round trip through a parser's escaping rules.
//   4. Comments are dropped wholesale, which is what removes the InkNode placeholders.
//      The ink is delivered as a PNG instead, by ./ink.ts.
//
// The parser is deliberately small and tolerant: no entity decoding, no implied end tags,
// no `<pre>`/`<script>`/`<style>` raw-text handling. Graph emits well-formed XHTML-shaped
// markup, and none of those cases appear in it. An unmatched close tag is ignored and an
// unclosed element is closed when its parent ends, so damaged markup produces trimmed
// output rather than an exception — a page whose text is readable must not become a
// failed request over a stray tag.

/** Attributes kept as they are. Anything matching `data-*` is kept as well. */
const KEPT_ATTRIBUTES: ReadonlySet<string> = new Set([
  'id',
  'href',
  'src',
  'alt',
  'lang',
  'title',
  'colspan',
  'rowspan',
  'start',
  'name',
  'content',
  'value',
  'type',
]);

/** The style declarations that carry layout coordinates rather than presentation. */
const KEPT_STYLE_PROPERTIES: ReadonlySet<string> = new Set([
  'position',
  'left',
  'top',
  'width',
  'height',
]);

/** `width` and `height` are coordinates on an image and presentation everywhere else. */
const SIZED_ELEMENTS: ReadonlySet<string> = new Set(['img', 'object', 'iframe']);

/** Elements that never have an end tag. Emitted as `<name />`. */
const VOID_ELEMENTS: ReadonlySet<string> = new Set([
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
  'param',
  'source',
  'track',
  'wbr',
]);

/**
 * Elements that count as content even when they hold no text.
 *
 * Table structure is in here because dropping an empty `<td>` shifts every cell after it
 * into the wrong column, which changes what the table says.
 */
const CONTENT_ELEMENTS: ReadonlySet<string> = new Set([
  'html',
  'head',
  'body',
  'title',
  'meta',
  'img',
  'br',
  'hr',
  'object',
  'iframe',
  'input',
  'table',
  'thead',
  'tbody',
  'tfoot',
  'tr',
  'td',
  'th',
]);

/** Elements that exist only to carry styling, and are dropped once it is gone. */
const UNWRAPPABLE_ELEMENTS: ReadonlySet<string> = new Set(['span', 'font', 'div']);

/** The name given to the parse tree's root, which is never emitted. */
const ROOT_NAME = '#root';

interface Attribute {
  readonly name: string;
  readonly value: string | null;
}

interface ElementNode {
  readonly kind: 'element';
  readonly name: string;
  attrs: Attribute[];
  children: HtmlNode[];
}

interface TextNode {
  readonly kind: 'text';
  readonly text: string;
}

type HtmlNode = ElementNode | TextNode;

/**
 * Trim one page's HTML to readable structure.
 *
 * Every piece of text survives. What is removed is presentational attributes, comments,
 * and wrapper elements left with nothing to say.
 */
export function trimPageHtml(html: string): string {
  const root = parseHtml(html);
  transform(root);
  return collapseBlankLines(serializeChildren(root));
}

/**
 * Close the gaps a removed element leaves behind.
 *
 * This is the only rule that touches text, and it only ever replaces a run of blank
 * lines with one newline. HTML collapses whitespace when it is rendered, so no blank
 * line in this markup carries meaning — Graph emits no `<pre>`, and the parser above
 * does not treat it as raw text in any case.
 */
function collapseBlankLines(html: string): string {
  return html.replaceAll(/[ \t]*\n(?:[ \t]*\n)+/g, '\n').trim();
}

/** Parse into a tolerant tree. Comments and declarations are dropped as they are met. */
function parseHtml(html: string): ElementNode {
  const root: ElementNode = { kind: 'element', name: ROOT_NAME, attrs: [], children: [] };
  const stack: ElementNode[] = [root];
  let index = 0;

  const addText = (text: string): void => {
    if (text !== '') stack[stack.length - 1]?.children.push({ kind: 'text', text });
  };

  while (index < html.length) {
    const open = html.indexOf('<', index);
    if (open === -1) {
      addText(html.slice(index));
      break;
    }
    addText(html.slice(index, open));

    if (html.startsWith('<!--', open)) {
      const end = html.indexOf('-->', open + 4);
      index = end === -1 ? html.length : end + 3;
      continue;
    }
    if (html.startsWith('<!', open) || html.startsWith('<?', open)) {
      const end = html.indexOf('>', open);
      index = end === -1 ? html.length : end + 1;
      continue;
    }

    const close = findTagEnd(html, open);
    if (close === -1) {
      // An unterminated `<` is text, not a tag.
      addText(html.slice(open));
      break;
    }

    const inner = html.slice(open + 1, close);
    index = close + 1;

    if (inner.startsWith('/')) {
      closeElement(stack, inner.slice(1).trim().toLowerCase());
      continue;
    }

    const name = /^[^\s/>]+/.exec(inner)?.[0]?.toLowerCase();
    if (name === undefined) {
      // `< ` is not a tag open in any markup Graph produces; treat it as text.
      addText(html.slice(open, close + 1));
      continue;
    }

    const element: ElementNode = {
      kind: 'element',
      name,
      attrs: parseAttributes(inner.slice(name.length)),
      children: [],
    };
    stack[stack.length - 1]?.children.push(element);
    if (!VOID_ELEMENTS.has(name) && !inner.trimEnd().endsWith('/')) stack.push(element);
  }

  return root;
}

/** The index of the `>` that ends a tag, skipping any inside a quoted value. */
function findTagEnd(html: string, open: number): number {
  let quote: string | null = null;
  for (let i = open + 1; i < html.length; i++) {
    const char = html[i];
    if (quote !== null) {
      if (char === quote) quote = null;
    } else if (char === '"' || char === "'") {
      quote = char;
    } else if (char === '>') {
      return i;
    }
  }
  return -1;
}

/**
 * Pop to the nearest matching open element.
 *
 * A close tag with nothing to match is ignored rather than unwinding the stack, because
 * one stray `</div>` would otherwise reparent the whole rest of the page.
 */
function closeElement(stack: ElementNode[], name: string): void {
  for (let i = stack.length - 1; i > 0; i--) {
    if (stack[i]?.name === name) {
      stack.length = i;
      return;
    }
  }
}

function parseAttributes(text: string): Attribute[] {
  const attrs: Attribute[] = [];
  const pattern = /([^\s=/>]+)(?:\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s>]+)))?/g;

  for (const match of text.matchAll(pattern)) {
    const name = match[1]?.toLowerCase();
    if (name === undefined || name === '/') continue;
    attrs.push({ name, value: match[2] ?? match[3] ?? match[4] ?? null });
  }

  return attrs;
}

/** Filter attributes, drop what says nothing, and unwrap what is left holding nothing. */
function transform(element: ElementNode): void {
  element.attrs = keepAttributes(element);

  const children: HtmlNode[] = [];
  for (const child of element.children) {
    if (child.kind === 'text') {
      children.push(child);
      continue;
    }
    transform(child);
    if (!isWorthKeeping(child)) continue;
    if (isUnwrappable(child)) children.push(...child.children);
    else children.push(child);
  }
  element.children = children;
}

function keepAttributes(element: ElementNode): Attribute[] {
  const kept: Attribute[] = [];

  for (const attr of element.attrs) {
    if (attr.name === 'style') {
      const style = keepStyle(attr.value ?? '');
      if (style !== '') kept.push({ name: 'style', value: style });
      continue;
    }
    if (
      attr.name.startsWith('data-') ||
      KEPT_ATTRIBUTES.has(attr.name) ||
      ((attr.name === 'width' || attr.name === 'height') && SIZED_ELEMENTS.has(element.name))
    ) {
      kept.push(attr);
    }
  }

  return kept;
}

/** The layout declarations of a `style` value, in source order, or `''` if it has none. */
function keepStyle(value: string): string {
  const kept: string[] = [];

  for (const declaration of value.split(';')) {
    const colon = declaration.indexOf(':');
    if (colon === -1) continue;
    const property = declaration.slice(0, colon).trim().toLowerCase();
    const propertyValue = declaration.slice(colon + 1).trim();
    if (propertyValue !== '' && KEPT_STYLE_PROPERTIES.has(property)) {
      kept.push(`${property}:${propertyValue}`);
    }
  }

  return kept.join(';');
}

/**
 * Whether an element survives the trim.
 *
 * An element carrying an id is kept whatever it holds: the ids are what a PATCH action
 * targets, and an empty paragraph is a blank line the author put there. What this drops
 * is the styling wrapper — a `<span>` whose only attribute was `style` and whose only
 * content was the styling itself.
 */
function isWorthKeeping(element: ElementNode): boolean {
  if (CONTENT_ELEMENTS.has(element.name)) return true;
  if (element.attrs.some((attr) => attr.name === 'id' || attr.name === 'data-id')) return true;
  return element.children.some(hasContent);
}

function hasContent(node: HtmlNode): boolean {
  if (node.kind === 'text') return node.text.trim() !== '';
  if (CONTENT_ELEMENTS.has(node.name)) return true;
  return node.children.some(hasContent);
}

/** A styling wrapper with no styling left is one tag pair of pure noise. */
function isUnwrappable(element: ElementNode): boolean {
  return UNWRAPPABLE_ELEMENTS.has(element.name) && element.attrs.length === 0;
}

function serializeChildren(element: ElementNode): string {
  return element.children.map(serialize).join('');
}

function serialize(node: HtmlNode): string {
  if (node.kind === 'text') return node.text;

  const attrs = node.attrs.map(serializeAttribute).join('');
  if (VOID_ELEMENTS.has(node.name)) return `<${node.name}${attrs} />`;
  return `<${node.name}${attrs}>${serializeChildren(node)}</${node.name}>`;
}

function serializeAttribute(attr: Attribute): string {
  if (attr.value === null) return ` ${attr.name}`;
  // Source values may have been single-quoted, so a `"` inside one is possible.
  return ` ${attr.name}="${attr.value.replaceAll('"', '&quot;')}"`;
}
