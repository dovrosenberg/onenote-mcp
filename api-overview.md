# OneNote Graph API: what this repository relies on

A summary of <https://learn.microsoft.com/en-us/graph/onenote-get-content>, plus the
things that page does not say and that were found by running against the real account.
The source page is the authority on what the service offers; the **Measured** sections
below are the authority on what it actually did here, because several of them contradict
the documentation.

Read this before adding a Graph call. The throttling limits and the design principles
they force are in the `Graph request budget` section of `CLAUDE.md`.

## Resource paths

All paths hang off `https://graph.microsoft.com/v1.0/me/onenote`.

| Path | Returns | Query options the doc lists |
|---|---|---|
| `/notebooks` | Every notebook | `filter, orderby, select, top, skip, expand, count` |
| `/notebooks/{id}` | One notebook | `select, expand` |
| `/notebooks/{id}/sections` | Sections directly under a notebook | `filter, orderby, select, top, skip, expand, count` |
| `/notebooks/{id}/sectionGroups` | Section groups directly under a notebook | same |
| `/sectionGroups` | Every section group, including nested ones | same |
| `/sectionGroups/{id}` | One section group | `select, expand` |
| `/sectionGroups/{id}/sections` | Sections directly under a section group | `filter, orderby, select, top, skip, expand, count` |
| `/sections` | Every section in the account, including those in nested groups | same |
| `/sections/{id}` | One section | `select, expand` |
| `/sections/{id}/pages` | Pages in one section | `filter, orderby, select, expand, top, skip, count, pagelevel` |
| `/pages` | Page metadata across all notebooks | `filter, orderby, select, expand, top, skip, count` |
| `/pages/{id}` | One page's metadata | `select, expand, pagelevel` |
| `/pages/{id}/content` | The page HTML | `includeIDs` |
| `/pages/{id}/preview` | A text snippet, up to 300 characters | — |
| `/resources/{id}/$value` | The bytes of an image or attachment | — |

**Measured.** Two of these do not work on this account, whatever the doc says:

- `/pages` (account-wide) fails with error `20266`, "maximum sections exceeded". It is
  banned in `src/graph-structure.ts` and a test scans `src/` for it.
- `/sections` (account-wide) with **no** `$filter` answers `500` with code `19999`. With a
  `$filter` it answers normally. `/sectionGroups` (account-wide) answers `500` with
  code `19999` filtered or not, and answered `504` on one unfiltered attempt.

## Query options

| Option | Doc says |
|---|---|
| `$select` | Properties to return. Property names are case-sensitive. |
| `$expand` | Navigation properties to inline. |
| `$filter` | Boolean expression. Operators `eq ne gt ge lt le`, `and or not`; functions `contains endswith startswith length indexof substring tolower toupper trim concat`. |
| `$orderby` | Any property of the entity, `asc` (default) or `desc`. |
| `$top` | Result count, **default 20, maximum 100**. |
| `$skip` | Entries to skip, for paging. |
| `$count` | `count=true` puts the total in `@odata.count`. |
| `pagelevel` | Page indentation level and order within a section. Only on a section's pages, or one page. |

Defaults worth knowing: pages sort `lastModifiedTime desc`; notebooks, section groups and
sections sort `name asc`. A pages request with no `$top` returns 20 and an
`@odata.nextLink`.

**Measured.** `$top=200` on a section's pages is `400` with code `20129`, "the limit of
'100' for the `$top` query has been exceeded" — the 100 ceiling is enforced, not advisory.

## What `$expand` accepts

| Entity | Expandable |
|---|---|
| Notebooks | `sections`, `sectionGroups` |
| Section groups | `sections`, `sectionGroups`, `parentNotebook`, `parentSectionGroup` |
| Sections | `parentNotebook`, `parentSectionGroup` |
| Pages | `parentNotebook`, `parentSection` |

Expanding a parent's children and a child's parents in the same request is a circular
reference and is not supported.

**Measured.** Nesting is capped at two levels. This works:

```
GET /me/onenote/notebooks?$select=id,displayName
    &$expand=sections($select=id,displayName),
             sectionGroups($select=id,displayName;$expand=sections($select=id,displayName))
```

and returns 54 notebooks, their 290 sections, their 43 section groups and those groups'
270 sections in one request, 78 KB, about 3 seconds. A third level of `sectionGroups`
inside that, and the `$levels=max` form the documentation shows, both answer `400`. That
cap is why `getExpandedTree()` reaches one level of section group and `getFullTree()`
still exists.

The separator inside one clause carrying both `$select` and `$expand` is a **semicolon**:
`sectionGroups($select=id,displayName;$expand=sections(...))`.

`$select` inside every expand clause is worth 5.7x — 441 KB without it, 78 KB with it, for
the same tree.

**Measured 2026-08-19.** `lastModifiedDateTime` is returned inside an `$expand` clause,
and at the top level, on every entity asked for:

```
GET /me/onenote/notebooks?$select=id,displayName,lastModifiedDateTime
    &$expand=sections($select=id,displayName,lastModifiedDateTime),
             sectionGroups($select=id,displayName;$expand=sections($select=id,displayName,lastModifiedDateTime))
```

55 of 55 notebooks and 568 of 568 sections carried the field. The account has grown since
the 78 KB measurement above; the same tree is 111,615 bytes with the timestamps added and
79,660 bytes without them, so the field costs about 40%.

**Measured 2026-08-19, and this is an availability property, not a syntax one.**
`$expand` on `/notebooks` returned `500` with code `19999` — *"Something failed, the API
cannot share any more information at the time of the request"* — continuously for about
seven minutes, then recovered on its own with no change to the request:

| Request | During the window |
|---|---|
| `/notebooks?$select=id,displayName` | **200** |
| `/notebooks?$select=id,displayName,lastModifiedDateTime` | **200** |
| `/notebooks?$expand=sections($select=id,displayName)` | 500, 19999 |
| `/notebooks?$expand=sectionGroups($select=id,displayName)` | 500, 19999 |
| the full `getExpandedTree()` URL, unchanged | 500, 19999 |

18 attempts spanning 19:34:53Z to 19:40:45Z all failed; 19:41:35Z answered 200. So any
`$expand` on the notebooks collection can be unavailable for minutes at a time while
un-expanded calls on the same collection succeed. `src/graph-throttle.ts` retries only
`429` and `503`, so a `500` here is not retried and aborts whatever asked for it — which
is `search_pages`, `find_page_by_name`, `list_pages_by_name`, `create_page_by_name` and
`append_to_page_by_name`, all of which go through `getExpandedTree()`.

## `lastModifiedDateTime` on a section rolls up from its pages

**Measured 2026-08-19**, against a section in a scratch notebook, reading
`/me/onenote/sections/{id}?$select=id,lastModifiedDateTime` between each step:

| Step | Section `lastModifiedDateTime` |
|---|---|
| before | `2026-08-19T14:38:36Z` |
| after `POST /sections/{id}/pages` (201) | `2026-08-19T19:32:39Z` |
| after `PATCH /pages/{id}/content` (204) | `2026-08-19T19:32:48Z` |
| after `DELETE /pages/{id}` (204) | `2026-08-19T19:32:57Z` |

So creating, editing **and deleting** a page all move the parent section's timestamp.
Microsoft documents the field on `onenoteSection` without saying it rolls up; it does.

The control matters, because the three deltas above track wall clock and would look the
same if the field simply reported "now". It does not: three reads of the same section 20
seconds apart with no write of any kind in between all returned `2026-08-19T19:32:57Z`
unchanged. The field moves on a write and only on a write.

A re-read of the deleted page answered `404`.

## A page rename does not move the page's `lastModifiedDateTime`

**Measured 2026-08-21**, against a scratch page, reading
`/me/onenote/sections/{id}/pages?$select=id,title,lastModifiedDateTime` between each step:

| Step | Page `title` | Page `lastModifiedDateTime` |
|---|---|---|
| after `POST /sections/{id}/pages` (201) | `probe-clock-2026-08-21` | `2026-08-21T12:17:52Z` |
| after `PATCH /pages/{id}/content` replacing the title (204) | `probe-clock-2026-08-21 RENAMED` | `2026-08-21T12:17:52Z` |
| re-read ~2 minutes later | `probe-clock-2026-08-21 RENAMED` | `2026-08-21T12:17:52Z` |

The title changed and the timestamp did not, so a rename is invisible to any comparison
that reads only `lastModifiedDateTime`. This is not lag: the third row is the same value
two minutes on.

The consequence is that a timestamp is not sufficient to decide whether a page needs
re-reading. Anything that skips work on an unmoved stamp — the sweep's drift check, and
the page-level skip — has to compare the title as well, or a page renamed outside this
server keeps its old title in the mirror for ever, which is the field `find_page_by_name`
and `search_pages` match on. The title costs nothing to compare: every listing this
repository makes already selects it.

Measured through `PATCH /pages/{id}/content` with a `title` target, which is what
`update_page_title` sends. Whether a rename performed **in the OneNote client** behaves
the same way is not yet measured; it is the case that matters, because a rename made
through this server already updates the mirror directly.

## Page timestamps have whole-second resolution, and this server's do not

**Measured 2026-08-21.** Graph returned `2026-08-21T12:17:52Z` for the page above — no
fractional seconds, in the same shape as the section timestamps recorded above.
`resyncPage` stamps a page locally with `new Date().toISOString()`, which always emits
three fractional digits.

So the two are not comparable as strings. `'2026-08-21T12:17:52.400Z' < '2026-08-21T12:17:52Z'`
is **true** lexicographically, because `.` is `0x2E` and `Z` is `0x5A` — a locally stamped
page landing in the same second as Graph's own stamp sorts *before* it and reads as
behind. `pageHasDrifted` compares `Date.parse` milliseconds for this reason. Do not
"simplify" it back to a string comparison.

## A section's page listing reports a title immediately; `GET /pages/{id}` does not

**Measured 2026-08-21.** `GET /sections/{id}/pages?$select=id,title,lastModifiedDateTime`
issued seconds after `POST /sections/{id}/pages` returned the new page with its title
already correct. That is the opposite of the page-metadata weakness recorded under
**Writing page content**, where `GET /pages/{id}?$select=title` answered `""` for pages
created seconds earlier.

So the listing endpoint is a usable source of titles and the per-page metadata endpoint is
not. A title comparison built on a listing does not need the empty-string guard that one
built on `GET /pages/{id}` would.

## Clock skew between Graph and this service: seconds, not minutes

**Measured 2026-08-21**, loosely. A page created between local `12:17:39.939Z` and
`12:17:55.387Z` was stamped by Graph at `12:17:52Z` — inside the bracket. The bracket is
wide because it contains a full client round trip, so this bounds the skew rather than
measuring it, and the clock it was taken against is a developer workstation rather than
the deployed service.

What it rules out is gross skew: not minutes, not hours. That is the property the
watermark overlap is sized against, so an overlap of minutes has room. A tighter number
needs the telemetry that reports the difference from inside the service on every run.


## Moving a section between notebooks: **unmeasured**

Nothing here has been run against the service, and the sync depends on it. Two questions,
each with a stated consequence, so one live run can settle both:

1. **Does a section moved between notebooks in the OneNote client keep its Graph id?**
   If it does, the mirror's existing section document survives the move and keeps its
   `pagesSyncedThrough` — `sectionIdentity` covers `notebookId` so the tree fields are
   rewritten, but the watermark is not one of them. If it does not, the old id vanishes
   from the tree, its document is deleted by absence, and a new one is created with
   `pagesSyncedThrough: null`, which is a candidate on the next run whatever the clock
   says.
2. **Does the move bump the section's `lastModifiedDateTime`?** The field rolls up from
   page writes (measured, above); whether a *move* counts as a write to it is not known.
   If it does, tier 1 of `pickCandidates` makes the section a candidate on the next run
   and question 1 stops mattering.

Only "id stable **and** timestamp unmoved" leaves a gap: a section carrying months of
arrears moved into a mirrored, active notebook is never widened and never re-listed. The
`pickCandidates` docstring in `src/mirror-sync.ts` records that class as knowingly
uncovered and points here.

How to settle it: read `/me/onenote/sections/{id}?$select=id,displayName,`
`lastModifiedDateTime,parentNotebook` for a scratch section, move it between two notebooks
in the OneNote client, then read the same section id again and re-read the account-wide
expanded tree. A 404 on the id answers question 1; a changed timestamp answers question 2.

## `$filter` on a datetime

**Measured 2026-08-19.** Unquoted ISO-8601 UTC is accepted on a section's pages:

```
GET /me/onenote/sections/{id}/pages?$select=id,title,lastModifiedDateTime&$top=100
    &$filter=lastModifiedDateTime ge 2026-05-21T19:32:31Z
```

200, and the filter is applied. The unfiltered control on the same section confirmed the
documented default sort — `lastModifiedTime desc` — holds in practice, so a client-side
cutoff on the first page of results costs the same one request and is a working fallback
if this filter ever stops being accepted.

## Case sensitivity in `$filter`

The doc says property names and string comparisons are case-sensitive, and recommends
`tolower()`.

**Measured, and this is where the service is inconsistent:**

| Request | Result |
|---|---|
| `/notebooks?$filter=displayName eq 'Bullet Journal - 2026'` | 200, exact case only |
| `/notebooks?$filter=tolower(displayName) eq 'bullet journal - 2026'` | 200 |
| `/notebooks?$filter=contains(displayName,'2026')`, `startswith(...)` | 200 |
| `/notebooks/{id}/sections?$filter=tolower(displayName) eq '…'` | 200 |
| `/sectionGroups/{id}/sections?$filter=tolower(displayName) eq '…'` | 200 |
| `/sections?$filter=displayName eq 'Monthly Log'` (account-wide) | 200 |
| `/sections?$filter=tolower(displayName) eq 'monthly log'` (account-wide) | **500, code 19999** |
| `/sections?$filter=contains(tolower(displayName),'monthly log')` (account-wide) | 200 |
| `/sections/{id}/pages?$filter=tolower(title) eq 'monthly log'` | 200 |

So on the account-wide sections collection, `tolower()` is usable inside `contains` and
not with `eq`. `findSectionsByName` in `src/graph-structure.ts` uses the `contains` form
and applies the full-name comparison itself.

## The two calls that make name lookup cheap

**A section anywhere, with the containers it sits in, in one request.** No walk reaches a
section nested past the `$expand` cap; this does:

```
GET /me/onenote/sections?$select=id,displayName
    &$expand=parentNotebook($select=id,displayName),parentSectionGroup($select=id,displayName)
    &$filter=contains(tolower(displayName), 'monthly log')
```

**Page titles matched by the service.** Removes any bound on how many pages a section can
hold before a title match could be missed:

```
GET /me/onenote/sections/{id}/pages?$select=id,title,lastModifiedDateTime
    &$filter=tolower(title) eq 'monthly log'
```

## Page content

`/pages/{id}/content` takes two query options this repository cares about:

- `includeIDs=true` — adds the generated `id` attributes that a PATCH targets. Without it
  a client-authored page carries no id and no `data-id`, so nothing positional can be
  addressed. Issue #18 needs this.
- `includeInkML=true` — returns `multipart/mixed`: the HTML in one part, the InkML strokes
  in another. Without it the handwriting is replaced by
  `<!-- InkNode is not supported -->` and is simply gone. `src/page-content.ts` is the
  only caller.

**Measured.** The two compose. `?includeIDs=true&includeInkML=true` answers `200` with
`multipart/mixed`, and the HTML part carries the generated ids.

**Measured.** `includeIDs=true` adds the generated `id` attributes and nothing else. A
`data-id` comes back either way — but only if the input HTML that created the element set
one. A page authored in the OneNote client carries no `data-id` anywhere, so on a real page
the generated ids are the only handle a PATCH has.

## Writing page content

A page's content is changed by `PATCH /me/onenote/pages/{id}/content` with
`Content-Type: application/json` and a JSON array of change objects. Success is `204`
with an empty body. Source:
<https://learn.microsoft.com/en-us/graph/onenote-update-page>.

| Attribute | Doc says |
|---|---|
| `target` | `title`, `body`, `#{data-id}`, or a generated `id`. `title` and `body` take no `#`; a generated `id` takes no `#`; a `data-id` requires one. |
| `action` | `append`, `insert`, `prepend`, `replace`. |
| `position` | `before` or `after`; `after` when omitted. With `append` it selects first or last child, with `insert` it selects preceding or subsequent sibling. |
| `content` | Well-formed HTML. Binary data requires a `multipart/form-data` request with a `Commands` part. |

What the doc says each element accepts:

| Element | Replace | Append child | Insert sibling |
|---|---|---|---|
| `body` (the first div on the page) | no | yes | no |
| `div`, absolutely positioned | no | yes | no |
| `div` within a div | yes, generated id only | yes | yes |
| `img`, `object` within a div | yes | no | yes |
| `ol`, `ul` | yes, generated id only | yes | yes |
| `table` | yes, generated id only | no | yes |
| `p`, `li`, `h1`–`h6` | yes, generated id only | no | yes |
| `title` | yes | no | no |

`tr`, `td`, `span`, `a`, `meta`, `head`, `style`, and absolutely positioned `img` and
`object` accept nothing. Writing needs `Notes.ReadWrite` or `Notes.ReadWrite.All`.

**Measured**, 2026-08-18, on throwaway pages this spike created and deleted in its own
scratch notebook. The three request bodies issue #18 needs:

```jsonc
// replace the title
[{ "target": "title", "action": "replace", "content": "New title" }]

// append to the page body
[{ "target": "body", "action": "append", "content": "<p>appended</p>" }]

// insert a sibling above or below an element addressed by its data-id
[{ "target": "#beta", "action": "insert", "position": "before", "content": "<p>above</p>" }]
```

- **The title is set to the content string verbatim, and nothing in it is parsed.**
  Content `<p>marked up title</p>` produced a page whose title is the literal string
  `<p>marked up title</p>`. `&`, `<`, `>` and `"` survive unescaped: `A & B <c> "d" 5 < 6`
  came back identical. A tool that builds a title needs no escaping and gets no markup.
- **`title` accepts `replace` and nothing else.** `action: "append"` is `400` code `20141`,
  "The PATCH target title for action APPEND is not supported."
- **`title` must not be written `#title`.** That is `400` code `20149`, "The target #title
  of your PATCH action can not be found".
- **`body` means the first top-level `div`, and which div that is depends on how the page
  was made.** A page created through Graph whose input `<body>` lacks
  `data-absolute-enabled="true"` comes back with everything inside one
  `<div data-id="_default">`, so `body` is effectively the whole page. A page authored in
  the OneNote client has sibling top-level divs — the page sampled here has three — and so
  does a created page whose body carries `data-absolute-enabled="true"`. On such a page,
  `body`+`append` lands as the last child of the **first** outline, not at the bottom of
  the page. Reaching another outline means targeting that div's generated `id`.
- **`prepend` and `append`+`position: before` do the same thing**: the content becomes the
  first child of the target. Both returned `204` and both landed inside the first outline.
- **Positional targeting relative to a `data-id` works.** `insert` with `position` `before`
  and `after` against `#beta` each returned `204` and placed the sibling as asked. That is
  what issue #27 depends on.
- **A `data-id` target needs the `#`.** `"target": "beta"` is `400` code `20134`, "The
  Patch request message is invalid: The selected target beta is not a valid updateable
  element."
- **Replacing a `p` needs the generated id.** By `#{data-id}` it is `400` code `20141`,
  "The PATCH target P for action replace is not supported"; by generated id it is `204`.
- **Generated ids change when the page is updated.** The id of a replaced paragraph was
  gone from the next read. Anything that targets a generated id has to `GET
  ?includeIDs=true` first, in the same operation. Their form is `p:{guid}{39}`; an outline
  div nested by the service gets a composite `div:{guid}{39}:{guid}{39}`.
- **A `data-id` given in submitted content survives.** Content
  `<p data-id="appended">…</p>` was still addressable as `#appended` on a later read.
- **The array is applied in order and as a unit.** A title replace and a body append in one
  request both took effect. An array holding one valid change and one naming a missing
  target failed `400` code `20120` and the valid change was *not* applied.
- **Malformed content does not fail.** `<p>unclosed` returned `204` and the service closed
  the tag.
- **An empty array is `400` code `20125`**, "The PATCH request contains no actions", and an
  unknown action is `400` code `20122`.
- **A new outline cannot be created, and no outline can be moved.** Measured 2026-08-19,
  across a page created with `data-absolute-enabled="true"` and two positioned outlines, a
  page created without it, and a client-authored page carrying ink. An absolutely
  positioned `<div>` sent as `content` to `target: "body"` returns `204`, but the div is
  flattened into the first outline and its `position`, `top` and `left` are dropped — the
  documented rule is that absolute positioning applies only to direct children of `<body>`,
  and `body` as a PATCH target is the first div, not the body element. `insert` beside an
  outline is `400` code `20135`, "The entity type is not supported for this operation", and
  so is `insert` against `body`. `replace` on an outline by generated id is `20134`, and by
  `#{data-id}` it is `20141`. `replace` on a *nested* div fails the same way. So absolute
  positioning is a create-time property: on an existing page, content can only be added
  inside an outline that is already there.
- **Vertical space can only be made with `<br>`.** `margin-top:278pt` on an appended
  paragraph came back as the default `5.5pt`, and an appended `<p>&nbsp;</p>` was dropped
  from the page entirely. Eight `<br>` elements in one paragraph came back as eight. That
  is what `src/page-layout.ts` uses to push appended text below existing ink.
- **A `data-id` survives only on an element that carries content.** In one PATCH,
  `<p data-id="m1-text">text</p>` and `<div data-id="m3-div"><p>text</p></div>` came back
  with their ids; `<p data-id="m2-breaks"><br /><br /></p>` and the same wrapped in a div
  came back with the attribute gone. A marker attribute has to ride on real content.
- **A PATCH is visible to the next content read** — measured 3.7 seconds after the request,
  including both round trips. Page *metadata* is weaker: `GET /pages/{id}?$select=title`
  returned `""` for two of the pages created during this spike within seconds of creating
  them, while the create response itself carried the right title. Do not confirm a write by
  re-reading metadata immediately.

### Creating a page

`POST /me/onenote/sections/{id}/pages` with the page as one HTML document. Source:
<https://learn.microsoft.com/en-us/graph/api/section-post-pages>.

| What | Value |
|---|---|
| `Content-Type` | `text/html` for an HTML-only page; `multipart/form-data` only when the request also carries binary parts. |
| Body | A whole document: `<html><head><title>…</title></head><body>…</body></html>`. |
| Success | `201` with the created page as JSON — `id`, `title`, `contentUrl`, and `links.oneNoteWebUrl` / `links.oneNoteClientUrl`. |

**Measured**, 2026-08-18, on a page created in a scratch notebook by the acceptance run
for issue #18. `Content-Type: text/html` with the document below answered `201` and the
body carried `id`, `title`, `contentUrl`, and both `links` hrefs.
`multipart/form-data` is not used, because this server submits no binary.

Two more things that run measured:

- **The whole round trip works.** `create_page`, then `append_to_page`, then
  `update_page_title` on the same page: the created heading, paragraph and list all came
  back on the next content read, the appended paragraph landed *after* the created
  content rather than replacing it, and the renamed title appeared in
  `/sections/{id}/pages` within 5 seconds. That is faster than the #17 spike saw, where a
  metadata read answered `""` — 5 seconds is one measurement, not a guarantee.
- **`includeIDs=true` produces targetable ids on a page created this way.** Seven of them
  on a page with six elements, in the documented form with the guid brace-wrapped:
  `div:{guid}{32}`, `h1:{guid}{39}`, `p:{guid}{42}`, `li:{guid}{45}`. All seven survive
  `trimPageHtml`, so what a caller reads is what a PATCH can target.

Two things about the submitted document come from the #17 spike:

- **The `<title>` element is the page title.** Everything that finds a page by name
  matches the title Graph stores, and that comes from `<title>`; a heading in the body
  sets nothing.
- **`<body data-absolute-enabled="true">` decides the page's shape.** With it the page has
  sibling top-level divs like a client-authored page. Without it Graph wraps the whole
  submission in one `<div data-id="_default">`, and `body` then addresses the whole page.
  `src/page-write.ts` omits it deliberately, so an append to a page this server created
  lands at the bottom of the page.

**Writes preserve ink.** Measured 2026-08-18 by
`test/ink-preservation.integration.test.ts` (issue #19) against a scratch page carrying 5
hand-written strokes. `append_to_page` and `update_page_title` each left the stroke count at
5 and the rendered PNG byte-identical — same 26,586 bytes, same SHA-256 — while the appended
paragraph appeared in the page HTML and the new title became visible, so both writes did
happen. A PATCH that names `body` or `title` does not disturb an `<ink>` node it says
nothing about. The test is skipped unless the environment names a live page; re-run it
against a page with handwriting before trusting a new change array shape.

## Errors

Failures carry an OData error body. The codes seen here:

| Code | Meaning |
|---|---|
| `10007` | Throttled. Arrives as 429. |
| `19999` | "Something failed, the API cannot share any more information." The account-wide sections and sectionGroups failures above. |
| `20112` | Invalid entity id. |
| `20120` | A PATCH target cannot be located. The whole change array is rejected. |
| `20122` | The PATCH action is not one Graph knows. |
| `20125` | The PATCH request contains no actions — an empty array. |
| `20129` | `$top` above 100. |
| `20134` | The PATCH target is not an updateable element. A `data-id` written without its `#` arrives here. |
| `20141` | The element does not support that action — `title` with `append`, or `p` replaced by `data-id`. |
| `20149` | The PATCH target cannot be found — `#title` rather than `title`. |
| `20266` | Maximum sections exceeded — the account-wide page list. |

Permissions for reads: `Notes.Read`, `Notes.ReadWrite`, or `Notes.ReadWrite.All`. This
service asks for the fully-qualified `Notes.Read` and `Notes.ReadWrite`; see
`GRAPH_SCOPES` in `src/graph-auth.ts`.
