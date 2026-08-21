# Page content conventions

Read before changing ink rendering, the HTML trimmer, page reads or writes, the layout
padding, or the name-lookup ladder.

## Ink and HTML

**The ink pipeline is a port, not a design.** `src/ink.ts` reproduces the recon script in
Appendix A of `project-spec.md`, which is validated against a real 440-stroke page. Four
things in it look arbitrary and are not: namespaces are stripped because Graph emits
`inkml:ink`; channel order is read from `<traceFormat>` because this account's points are
X, Y, F and F is pressure; coordinates are himetric and become px at 96 dpi; and traces
are collected from every `<ink>` root and every nesting level of `<traceGroup>`. Changing
any of the four produces a picture that is wrong rather than an error.

**A page with no ink is `null`, not an error.** Most pages are typed. `renderInk` returns
null whenever no `<ink>` root yields a stroke, including when a root parses to nothing:
`fast-xml-parser` accepts unclosed and mismatched tags without throwing, so "damaged" and
"holds no traces" cannot be told apart, and guessing would turn ordinary pages into failed
requests. The two errors that do exist are `InkParseError` for trace groups nested past 50
levels and `InkRenderError` for a document resvg rejects.

**No ink error message quotes the document.** Stroke coordinates are the user's
handwriting and this repository's output can reach a public log. The parse failure path
discards the thrown parser error for that reason: `fast-xml-parser` puts the offending
markup in its message.

**One fetch serves both halves of a page.** `GraphPageContent.fetchRaw` is the only
`includeInkML=true` call, and it returns the split parts. `fetchContent` turns that one
response into `{ html, ink }` — the trimmed HTML and the rendered PNG — and is what the
`get_page_content` tool wraps. A page with no ink comes back as plain
`text/html` rather than as `multipart/mixed`, so `pageHtml` falls back to the whole body
when there are no parts; dropping it because there are no parts to search would lose the
text of every typed page.

**The content URL asks for `includeIDs=true` as well.** Graph emits the generated `id`
attributes only when asked, so without it a page authored in the OneNote client comes
back with no id and no `data-id` anywhere and nothing on it can be targeted by a PATCH.
The two parameters compose — measured 2026-08-18: `includeIDs=true&includeInkML=true` is
200, `multipart/mixed`, ids present in the HTML part — so `src/multipart.ts` is
unaffected. Generated ids change whenever the page is updated, so anything that targets
one has to read it in the same operation that uses it; they cannot be cached between
calls.

**The HTML trimmer rewrites tags and never touches text.** `trimPageHtml` in
`src/page-html.ts` parses to a small tree, filters attributes, drops comments, and
removes or unwraps elements. Text nodes are copied to the output verbatim, so no entity
and no character of the user's writing can be lost to a parser's escaping rules. The one
exception is `collapseBlankLines`, which closes the gaps removed elements leave behind.
`test/page-html.test.ts` asserts the word sequence of the fixture is identical before and
after.

**Three things survive the trim that look like noise.** `id` and `data-id`, because the
PATCH write model in issue #18 targets an element by id. The `position`, `left`, `top`,
`width`, and `height` style declarations, because page content is laid out absolutely in
px at 96 dpi and that is the coordinate space `src/ink.ts` renders strokes into — nothing
needs it today and it cannot be recovered later. Empty `<td>` elements, because dropping
one shifts every cell after it into the wrong column. An element carrying an id is kept
whatever it holds; what gets dropped is the wrapper whose only attribute was `style`.

**The trimmer's parser is tolerant, not strict.** An unmatched close tag is ignored, an
unclosed element is closed when its parent ends, and an unterminated `<` is text. A page
whose text is readable must not become a failed request over a stray tag. It also does
not decode entities, imply end tags, or treat `<pre>`, `<script>`, and `<style>` as raw
text; Graph emits none of those, and adding an HTML library for this would be a
dependency for one endpoint.

## Tool results, writing, and name lookup

**A tool answers with one JSON text block, and every count it reports is real.**
`list_pages` says `moreAvailable` when Graph returned exactly `top` items, and
`search_pages` says how many sections it searched out of how many it found. A model
cannot tell a truncated search that matched nothing from a complete one, so a bounded
result that omitted the bound would be read as "no such page". `get_page_content` is the
one tool that answers with a second block, and it is an image block for the reason below.

**`get_page_content` returns the ink as an MCP image block, and that is the point.** The
calling model reads the handwriting with its own vision; there is no OCR service and no
handwriting API anywhere in this repository. Base64 inside the JSON text block would
reach the model as characters it cannot see, and a file path would name a container
filesystem no client can read — `test/page-tools.test.ts` asserts the PNG's base64 prefix
does not appear in the text block. A page with no ink yields one block and `inkImage:
null`; that is the normal answer for a typed page and not an error.

**All of a page's ink is one cropped image.** Ink and typed content are independent on
these pages, so the bounding box in `strokesToSvg` loses nothing worth keeping, and
nothing here splices ink fragments back into the HTML at the positions they came from.

**The ink PNG is shrunk by measurement, not by arithmetic.** `fitInkToByteBudget` in
`src/ink.ts` re-rasterises the same SVG at 0.75× the width until the PNG fits
`MAX_INK_PNG_BYTES`, because PNG size depends on how dense the strokes are and cannot be
predicted from the pixel count. It stops at `MIN_RENDER_WIDTH` even when that still does
not fit: an image too small to read is a better answer than a failed request, and the
result says which width it got so the model can tell "illegible handwriting" from
"rendered too small".

**Writing is a PATCH of change objects, and the shapes are measured, not guessed.**
`src/page-write.ts` sends `[{target, action, content}]` to
`PATCH /me/onenote/pages/{id}/content` and reads a 204 with no body. The three shapes and
every error code they can produce were measured against the live service on 2026-08-18 by
the spike in issue #17 and are recorded in `api-overview.md` under **Writing page
content**. The array is applied as a unit: one change naming a missing target fails the
whole request with 400 code 20120 and applies none of the others.

**`target: "body"` is the first outline, not the page, and `append_to_page` says so.** A
page authored in the OneNote client has sibling top-level divs, and an append lands at the
end of the first of them. Reaching another one means reading `?includeIDs=true` and
targeting that div's generated id, which is issue #27. The tool description states where
the content went, because a 204 does not.

**An append is padded off the handwriting, because nothing else can move it.** OneNote
fixes ink in a page-level layer and an outline grows downwards, so appending to a page
whose strokes hang below its text renders the text over the handwriting. Putting the new
content in its own outline below the ink is not possible: measured 2026-08-19, an
absolutely positioned div sent to `target: "body"` is flattened into the first outline
with its position dropped, `insert` beside an outline is 400 code 20135, and `replace` on
one is 20134 or 20141. Margins are normalised away and an empty paragraph is deleted, so
the only lever left is `<br>`, which survives verbatim. `src/page-layout.ts` reads the
outline's `top` and the ink's bounding box and decides how many breaks go in front of the
caller's fragment; `append_to_page` pays one extra Graph read per call for it, and the
result JSON reports the padding rather than hiding it.

**The padding is measured from an estimate of where the text ends, and marked so it
happens once.** No endpoint reports an outline's rendered height, so `estimateContentHeight`
counts block elements at deliberately low heights — `p`, `li` and `br` at 16px, `h1` at
24px, `tr` at 18px — plus the declared `height` of any `img`, which is the case a block
count gets worst. Every constant is chosen low on purpose: an estimate that is too small
pads a little too much, and one that is too large puts the new text back on the
handwriting. Wrapping is not modelled, which errs the same safe way. Measured against the
live page on 2026-08-19: 136px estimated where the real content was about 174px, so the
padding ran two blank lines long. The deeper of the estimate and the marker wins, so a
page whose content wraps is not padded twice. The fragment is wrapped in `<div data-id="ink-clearance-{px}">`, recording the page
position the content was brought down to — not the ink bottom, which reads as new ink on
the next call over rounding, and not on the padding itself, because the service discards a
`data-id` from an element holding only line breaks. `LINE_HEIGHT_PX = 19` and
`INK_CLEARANCE_MARGIN_PX = 12` are chosen numbers: nothing in the API reports a line
height, and only looking at the page in OneNote can confirm them.

**A page created here has one outline, deliberately.** `createPageHtml` omits
`data-absolute-enabled` from `<body>`, so Graph wraps the submission in a single
`<div data-id="_default">` and `body` then covers the whole page. That is what makes
`append_to_page` reach the bottom of a page this server created. Setting the attribute
would produce sibling outlines like a client-authored page, and appends would land in the
first one.

**The title is escaped on create and verbatim on rename.** `createPage` puts it in a
`<title>` element inside a document Graph parses, so it is escaped; `updatePageTitle`
sends it as PATCH `content`, which is stored character for character — the spike produced
a page actually titled `<p>x</p>`. Escaping the second would put `&amp;` in a title, and
not escaping the first would let a `<` open a tag.

**The write tools refuse before they spend a request.** `fragmentArgument` rejects a
fragment carrying `<html>`, `<head>`, `<body>`, `<title>`, `<meta>`, `<base>`, `<link>`,
`<script>` or `<style>`, and `titleArgument` rejects a title holding a complete tag. Both
throw `ToolInputError` and neither reaches Graph — on `create_page` a wrong request that
went through would leave a page behind for someone to find and delete. What is *not*
checked is well-formedness: `<p>unclosed` returns 204 and the service closes the tag, so a
strict parser here would refuse content that works. A bare `<` that opens no tag is left
alone in a title, because `if x <y then` is a legal title and this tool is the only way to
set one.

**Container names are matched by a ladder, and the result says which rung answered.**
`matchNodes` in `src/name-lookup.ts` tries exact and case-insensitive, then the same
comparison against the candidate with a leading ordering prefix removed, then a
case-insensitive substring. The middle rung is the one this account needs: its section
groups are named `062 - February` and a caller knows the month, not the number. A rung is
only tried when the one above it matched nothing, so an exact match can never lose to a
looser one — the test for that is two groups where only one strips to `February`.
`matchedBy` in every `_by_name` result names the rung, so a caller that asked for
`February` can see it got `062 - February` and why. Page titles do not use the ladder;
they are matched in full, ignoring case, and `search_pages` is the substring tool.

**A `_by_name` tool refuses rather than guesses.** A name matching nothing is a
`NameLookupError` listing the sibling names, never an empty result: a caller cannot tell
an empty section from a section that does not exist. A name matching more than once on
the same rung is a `NameLookupError` carrying the candidates. `sectionGroupName` omitted
means the section is a direct child of the notebook, not "search everywhere".

**`list_pages_by_name` is the answer to "which page is it?".** It returns every title in
a named section with its page id, so a model reads the titles, picks one, and passes that
id straight to `get_page_content`. Nothing else is needed in between, and the tool
description says so — a caller that re-resolved the name would pay for the lookup twice.

**The two `_by_name` writing tools resolve names after they check the content.**
`create_page_by_name` and `append_to_page_by_name` are `create_page` and `append_to_page`
with the section named rather than identified, over the same `resolveSection` the
browsing `_by_name` tools use, so writing to a known place costs one call instead of a
`list_notebooks` → `list_sections` walk followed by a write. The title and the fragment
are validated before the lookup runs: a refused argument costs neither the Graph request
that resolves the names nor the one that would write. A container name that matches
nothing or matches twice is a `NameLookupError` and nothing is written — a write tool
that guessed a section would put content in the wrong notebook and nothing would say so.
`resolvedPayload` lives in `src/name-lookup.ts` rather than in either tool module,
because the browsing tools and these two answer with it and a second copy would drift.

**A page title is not matched by the container ladder, and `append_to_page_by_name` says
so in its error.** Graph compares it in full and case-insensitively across the section,
which is the same comparison `find_page_by_name` uses. Zero matches and more than one are
both `NameLookupError`, because this is a write: `find_page_by_name` can answer with an
empty list and let the caller decide, and an append cannot. `NameLookupError` takes a
`matching` argument for that — the container message tells the caller a leading number is
stripped, which is true of a section group and false of a page title, so telling an
`append_to_page_by_name` caller to drop one would send it to do something that changes
nothing. The page-title message names `list_pages_by_name` and `search_pages` instead,
and lists no candidates: getting them costs a second request, and the tool that lists them
is one call away.

**Both appends go through one `append` helper, and that is deliberate.** It reads the
page, plans the ink clearance, writes, and builds the result. A second copy of that in
`append_to_page_by_name` that skipped the read would write across someone's handwriting,
return 204, and report success — the failure is only visible in OneNote.

**The fallback below a named section group is one filtered request, not a walk.**
`getExpandedTree` reaches a notebook's sections and one level of section group, because
Graph caps `$expand` nesting at two levels — a third answers 400, and so does
`$levels=max`. A section nested deeper is therefore absent from that response rather than
known to be absent, and `findSectionsByName` settles it in one request at any depth by
asking for sections account-wide with their parents expanded. It runs only when the
expanded tree came back empty-handed, and `deepSearchUsed` in the result says when it did.

**Let Graph do the comparisons it will do, and know which ones it refuses.** Page titles
are matched by `tolower(title) eq '…'` on the section's pages, so nothing bounds how many
pages a section may hold before a match could be missed. Sections account-wide cannot be
matched that way: `tolower(displayName) eq '…'` there answers 500 with code 19999, while
`contains(tolower(displayName), '…')` answers normally, so `findSectionsByName` asks for
the substring and applies the full-name rule itself. The table of what each endpoint
accepts is in `api-overview.md`.

**`list_sections` returns sections and section groups in one tagged list.** Graph exposes
them as two relationships and the caller has to recurse through both — this account is a
notebook per year with a section group per month — so a result that carried only sections
would make a caller guess that the second relationship exists.

**An unscoped `search_pages` walks sections, and the walk is bounded twice.** At most
`MAX_SECTIONS_SEARCHED` sections and `SEARCH_TIME_BUDGET_MS` of wall clock, both in
`src/page-search.ts`. The account-wide page list is not an option — it is the endpoint
that fails with 20266 — so there is nothing cheaper to fall back on. The time budget is
checked before each section is fetched rather than during, so an overrun costs one round
trip. A failure listing one section aborts the whole search: an expired refresh token
fails every section identically, and returning "no matches" for it would be an answer.
