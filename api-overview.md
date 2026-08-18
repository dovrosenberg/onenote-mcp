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

- `includeIDs=true` — adds the generated `id` / `data-id` attributes that a PATCH targets.
  Without it there are no ids in the HTML at all. Issue #18 needs this.
- `includeInkML=true` — returns `multipart/mixed`: the HTML in one part, the InkML strokes
  in another. Without it the handwriting is replaced by
  `<!-- InkNode is not supported -->` and is simply gone. `src/page-content.ts` is the
  only caller.

Whether the two compose on one request has not been tested.

## Errors

Failures carry an OData error body. The codes seen here:

| Code | Meaning |
|---|---|
| `10007` | Throttled. Arrives as 429. |
| `19999` | "Something failed, the API cannot share any more information." The account-wide sections and sectionGroups failures above. |
| `20112` | Invalid entity id. |
| `20129` | `$top` above 100. |
| `20266` | Maximum sections exceeded — the account-wide page list. |

Permissions for reads: `Notes.Read`, `Notes.ReadWrite`, or `Notes.ReadWrite.All`. This
service asks for the fully-qualified `Notes.Read` and `Notes.ReadWrite`; see
`GRAPH_SCOPES` in `src/graph-auth.ts`.
