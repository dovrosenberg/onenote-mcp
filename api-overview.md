# OneNote Graph API: what this repository relies on

A summary of <https://learn.microsoft.com/en-us/graph/onenote-get-content>, plus the
things that page does not say and that were found by running against the real account.
The source page is the authority on what the service offers; the **Measured** sections
below are the authority on what it actually did here, because several of them contradict
the documentation.

Read this before adding a Graph call. The throttling limits and the design principles
they force are in `docs/graph.md`.

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
re-reading. Anything that skips work on an unmoved stamp has to compare the title as well,
or a page renamed outside this server keeps its old title in the mirror for ever, which is
the field `find_page_by_name` and `search_pages` match on. Three places do:
`writePageFromRaw`'s short-circuit, `storedPageIsCurrent`, and `sweepSection` — the last
two through `pageListingDiffers` in `src/mirror-schema.ts`, which is one predicate so a
change to the rule cannot reach only one of them. The title costs nothing to compare: every
listing this repository makes already selects it.

The sweep is the place where omitting it has no upper bound. The incremental pass never
lists a page whose stamp is below the section watermark, so once a rename has been missed
the sweep is the only thing that could still see it.

Measured through `PATCH /pages/{id}/content` with a `title` target, which is what
`update_page_title` sends. Whether a rename performed **in the OneNote client** behaves
the same way is not yet measured; it is the case that matters, because a rename made
through this server already updates the mirror directly.

## Page timestamps have whole-second resolution, and this server's do not

**Measured 2026-08-21.** Graph returned `2026-08-21T12:17:52Z` for the page above — no
fractional seconds, in the same shape as the section timestamps recorded above.
`resyncPage` stamps a page locally with `new Date().toISOString()`, which always emits
three fractional digits.

So the two are not comparable as instants without parsing. `'2026-08-21T12:17:52.400Z' <
'2026-08-21T12:17:52Z'` is **true** lexicographically, because `.` is `0x2E` and `Z` is
`0x5A` — a locally stamped page landing in the same second as Graph's own stamp sorts
*before* it.

Nothing compares them as instants any more. `pageStampDiffers` in `src/mirror-schema.ts`
asks only whether the two strings are the same string, and the sweep answers a difference
by re-fetching the page. Two spellings of one instant therefore disagree, which is wanted:
the re-fetch's short-circuit writes Graph's own spelling back through
`MirrorStore.putPageMetadata`, so the locally stamped value is replaced rather than left in
a field every tool result prints.

## Graph reports the same page one second apart on two reads

**Measured 2026-08-21.** One scratch section's page listing
(`GET /sections/{id}/pages?$select=id,title,lastModifiedDateTime`) was read several times
in a row. Between two of those reads, every page in the section reported a
`lastModifiedDateTime` exactly one second later than it had before:

| Page | Earlier read | Later read |
|---|---|---|
| probe page | `2026-08-21T12:17:52Z` | `2026-08-21T12:17:53Z` |
| absolute spike | `2026-08-19T00:41:30Z` | `2026-08-19T00:41:31Z` |
| geometry spike | `2026-08-19T00:39:41Z` | `2026-08-19T00:39:42Z` |
| Renamed by update_page_title | `2026-08-19T00:24:22Z` | `2026-08-19T00:24:23Z` |

Three of the four had not been touched for two days. A genuine edit would have stamped
them with that day's date; they kept their 2026-08-19 dates and gained one second. So this
is not an edit — the same unchanged page reports two values one second apart on different
reads.

The uniform +1 is consistent with one read path flooring the sub-second component and
another ceiling it, but **the mechanism is not established**. What is established is the
observation: a page's reported stamp can move by one second with nothing having happened
to the page.

The consequence is that a stamp difference is not proof of an edit. It is **not** that a
margin is needed: a tolerance wide enough to absorb this second also discards every real
edit made inside it, and an edit the sweep never notices is served as current for ever.

What the sweep does instead is treat a difference as a hint. `pageStampDiffers` in
`src/mirror-schema.ts` compares the two strings and nothing more, and the sweep answers a
difference by re-fetching the page; `writePageFromRaw`'s content-hash comparison is the
authoritative check, and a page whose content turns out to be unchanged costs one Graph
request and a metadata write that stores Graph's stamp. So a page caught by this
one-second wobble is fetched once and then agrees. The direction of the difference is
irrelevant for the same reason — stored-ahead-of-live is the `resyncPage` local-stamp case
recorded under **Writing page content**, and a re-fetch corrects it where the old stale
mark destroyed it.

### The shift is not path-dependent

**Measured 2026-08-21**, after the shift above. The same page was read through two
different query shapes against the same section:

| Request | Reported `lastModifiedDateTime` |
|---|---|
| `GET /sections/{id}/pages?$select=id,title,lastModifiedDateTime&$orderby=lastModifiedDateTime desc&$top=N` | `2026-08-19T00:39:42Z` |
| `GET /sections/{id}/pages?$select=id,title,lastModifiedDateTime&$filter=tolower(title) eq '…'` | `2026-08-19T00:39:42Z` |

Same page, same value. Only two query shapes were tried; both are page listings on a
section, and no other endpoint was probed. The stamps have also been stable across every
read since the shift — the +1 second happened once and has not recurred.

Two things this does not establish. The mechanism of the original shift is still unknown;
the flooring-versus-ceiling explanation above remains a guess. And the shift was
correlated in time with the OneNote client first opening that notebook, which is a
correlation observed once and not a demonstrated cause.

Why it matters: the re-fetch design depends on the disagreement being transient. If two
query shapes reported the same page differently, then two sync passes reading through
different shapes would disagree permanently, `pageStampDiffers` would be true on every
run, and every page would be re-fetched every run — against a 400-request hourly budget,
with the mirror never converging. These two reads agree, so the design converges: a page
caught by the wobble is fetched once, its stored stamp is corrected to Graph's, and the
next run agrees.

## A section's page listing reports a title immediately; `GET /pages/{id}` does not

**Measured 2026-08-21.** `GET /sections/{id}/pages?$select=id,title,lastModifiedDateTime`
issued seconds after `POST /sections/{id}/pages` returned the new page with its title
already correct. That is the opposite of the page-metadata weakness recorded under
**Writing page content**, where `GET /pages/{id}?$select=title` answered `""` for pages
created seconds earlier.

So the listing endpoint is a usable source of titles and the per-page metadata endpoint is
not.

**That does not make `""` a safe sentinel for "no title read yet".** A later probe of the
same account found a page whose listing entry carries `"title": ""` — a genuinely untitled
page, not a read that failed. So an empty title is a legitimate value from either endpoint,
and any comparison has to handle it as one rather than treating it as missing data. What
the listing's reliability buys is that a title read from it can be *trusted*, not that it
can never be empty.

## Moving a page between sections changes its id and keeps its timestamp

**Measured 2026-08-21**, moving one page between two sections in the OneNote client.

**The page id changed completely** — both the GUID and the trailing section component:

```
before: 0-cc8f8e7e3eef4994be2452e57bb7ad1a!124-…!sc42d54e6be164de4b3b13f4a50c4160e
after:  0-e641c1d6852b089a057bcd40db72a242!1-…!s2b04cff65ae64bf6b5a31d46ce86cb57
```

A page id therefore embeds the section it sits in, and there is no id that survives a move.
To the mirror a moved page is a delete in the source section plus a create in the
destination, which is what the sweep already reconciles on both sides: the source's stored
page matches no live id and is tombstoned, and the destination's live id matches no stored
page and is fetched.

**The moved page kept its original `lastModifiedDateTime`** — `2026-08-21T12:17:53Z`, the
minute it was created, not the minute it was moved. This settles a premise the sweep's
rationale had been asserting without a measurement. A page moved into a section carries a
stamp older than the section's watermark, so `listPagesChangedSince` never returns it and
no incremental pass will ever see it. **Only the sweep finds a moved page.** That is the
measured justification for the sweep existing at all.

## Clock skew between Graph and this service: seconds, not minutes

**Measured 2026-08-21**, loosely. A page created between local `12:17:39.939Z` and
`12:17:55.387Z` was stamped by Graph at `12:17:52Z` — inside the bracket. The bracket is
wide because it contains a full client round trip, so this bounds the skew rather than
measuring it, and the clock it was taken against is a developer workstation rather than
the deployed service.

What it rules out is gross skew: not minutes, not hours. That is the property the
watermark overlap is sized against, so an overlap of minutes has room. A tighter number
needs the telemetry that reports the difference from inside the service on every run.


## Moving a section reissues its id across notebooks and keeps it within one

**Measured 2026-08-21**, moving one scratch section twice in the OneNote client. The
section held 3 pages and none of them was touched during either move.

**Move 1, into a section group in the same notebook.** The section id was unchanged:
`0-583EFEEEF6E35B4B!sc42d54e6be164de4b3b13f4a50c4160e` before and after. Its 3 pages kept
their ids and their `lastModifiedDateTime` values. The section went from being a direct
child of the notebook to a child of a section group named `test`.

**Move 2, into a different notebook.** The section id changed:

```
before: 0-583EFEEEF6E35B4B!sc42d54e6be164de4b3b13f4a50c4160e
after:  0-583EFEEEF6E35B4B!s7d0fdb0d111d440dbdba64cfcd92602e
```

Its 3 pages were reissued too, and the shape of the change is narrower than a page move's.
The GUID stayed the same and only the trailing section component followed the section:

```
before: 0-754d222de05c49a687a41a6d7d51c456!6-…!sc42d54e6…
after:  0-754d222de05c49a687a41a6d7d51c456!6-…!s7d0fdb0d…
```

Contrast a *page* moved between sections, above, where both the GUID and the section
component change. Page `lastModifiedDateTime` values were unchanged throughout both moves.

**A section id is stable across a move within a notebook and is reissued on a move between
notebooks.** That asymmetry is the non-obvious part: the id survives being reparented under
a section group, and does not survive changing notebook.

The consequence for the mirror is that the id change does the work on its own. The old
section id is absent from the next tree read, so its document is deleted by absence along
with its pages, and the new id creates a section document with `pagesSyncedThrough: null`,
which `pickCandidates` treats as a candidate whatever the tier-1 cutoff says and whatever
its notebook's activity says. The gap this section previously warned about — a section
carrying months of arrears moved into a mirrored active notebook and never widened —
**cannot occur**, because no watermark crosses a notebook boundary.

The cost that replaces it: a cross-notebook section move re-backfills that section in full.
Every page is fetched again under its new id and every page document under the old id is
deleted. Correct and self-healing, and it shows up in a run report as a section's worth of
requests against the hourly budget.

**Whether the move bumps the section's own `lastModifiedDateTime` was not measured.** The
section listing this account's tooling reads does not return that field, so the run that
settled the ids could not answer it. It no longer decides anything: the id change alone
makes the section a fresh backfill candidate, so tier 1 never gets the chance to skip it.

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
