# Active notebooks

2026-08-20

## Problem

The mirror's selection is one list: `notebookIds` in the root document names the notebooks
whose pages are mirrored. Every notebook in that list is treated identically by every sync
run — the incremental pass lists its changed sections, the scoped sweep enumerates its page
ids, the full sweep visits every one of its sections.

Most of the selected notebooks are not being edited. This account is a notebook per year,
so at any time one notebook is current and the rest are archives. Re-checking an archive
costs Graph requests against a 400-per-hour budget that the interactive tools share, and it
finds nothing.

## What this adds

A second, optional list in the same document: `activeNotebookIds`. A selected notebook that
is not active is still mirrored — its pages are backfilled once and served from the mirror
for ever after — but no scheduled or inline sync re-checks it. Writes through this server
still reach OneNote and still update the mirror, unchanged.

The change is bounded to three things: which sections a sync run visits, what a covered read
reports as its `source`, and one new route.

## Non-goals

- Activity does not change what is stored. An inactive notebook's pages, content, ink and
  structure are held exactly as an active notebook's are.
- Activity does not change the write path at all.
- Activity does not change structure freshness. `list_notebooks` and `list_sections` stay as
  current as they are today, for the reason given under **Structure is unaffected** below.
- No automatic promotion. A write to an inactive notebook does not make it active. The
  selection document is hand-edited and the service never writes it; that stays true.

---

## 1. The selection document

```json
{
  "notebookIds": ["1-abc", "1-def", "1-ghi"],
  "activeNotebookIds": ["1-abc"]
}
```

`readSelection` in `src/mirror-schema.ts` returns

```ts
export interface NotebookSelection {
  readonly notebookIds: readonly string[];
  /** null means "every selected notebook is active". */
  readonly activeNotebookIds: readonly string[] | null;
}
```

Reading rules, which follow the tolerance `readSelection` already has:

| In the document | Result | Why |
|---|---|---|
| field absent | `null` — all selected notebooks active | A deployment that has never heard of this feature behaves exactly as before. |
| not an array | `null` — all selected notebooks active | A malformed value must not silently freeze the mirror. Failing open costs Graph requests; failing closed stops the mirror updating with nothing to say so. |
| `[]` | none active | An empty array is a deliberate edit, not a half-finished one. It means "freeze everything", which is a state an operator may legitimately want. |
| array of strings | exactly those active | Entries trimmed, empty strings dropped, non-strings dropped individually, duplicates collapsed, order preserved — identical to `notebookIds`. |

The distinction between "absent" and "`[]`" is the one thing here a reader could get wrong,
and it is the reason `activeNotebookIds` is nullable in the type rather than defaulting to
`notebookIds`. A `readonly string[]` that happened to be empty could not tell those two
cases apart.

An id in `activeNotebookIds` that is not in `notebookIds` matches no mirrored notebook. It
is not an error: it is counted as `unknownActiveNotebookIds` in the sync report and logged
once per run as `mirror-selection-unknown-active` with a count and no ids, exactly as
`unknownNotebookIds` and `mirror-selection-unknown` already work.

New pure function, unit-tested directly:

```ts
export function isActive(selection: NotebookSelection, notebookId: string): boolean;
```

`true` when `activeNotebookIds` is null, or when it contains `notebookId`. It does not check
`notebookIds`; a notebook that is not mirrored never reaches a code path that asks.

## 2. What activity gates

### Structure is unaffected

`getExpandedTree()` is one Graph request for the whole account, and `buildStructure` already
stores every notebook, section group and section it returns regardless of selection — that
is what stops `list_notebooks` answering partially. Skipping inactive notebooks would save
nothing, because the request has already been made, and would break the property that makes
`list_notebooks` safe.

So `reconcileStructure`, `buildStructure`, `structureHashOf`, `putStructure` and
`learnNestedGroups` are untouched by activity. `list_notebooks` and `list_sections` keep the
freshness they have today.

### Page work is gated in three places

All in `src/mirror-sync.ts`.

| Place | Rule |
|---|---|
| `pickCandidates` | A section in an inactive notebook is a candidate only when `pagesSyncedThrough === null`. |
| `sweepPass` for `sweep` and `sweep-full` | Sections in inactive notebooks are filtered out before the cursor is applied. |
| `sweepPass` for the new `sweep-all` | No activity filter. |

The `pagesSyncedThrough === null` clause is the backfill. An inactive notebook fills up
exactly once — one `listPagesChangedSince` per section from the epoch, plus one content
fetch per page — and from then on no incremental run lists it again. This is what the
requirement "the initial backfill should fill them up" means mechanically.

`pickCandidates` is pure and exported so the rule is asserted directly rather than through a
fake store. It gains a `selection: NotebookSelection` parameter. Its existing early return
for `!state.sectionRollUpTrusted || !timestampsAreFresh` must apply the activity filter too:
distrusting the section roll-up is a reason to visit every *active* section, not a reason to
visit inactive ones.

`learnNestedGroups` keeps its `group.mirrored && !group.childGroupsKnown` filter with no
activity term. It is one request per group, once ever, and what it learns is structure.

### The new mode

`SyncMode` gains `'sweep-all'`. Follow-on edits: `syncModeOrNull` in `src/mirror-schema.ts`,
`SyncTarget` in `src/sync-route.ts`, a `runSweepAll` export in `src/mirror-sync.ts`, and the
binding in `src/tools.ts`.

```
POST /sync             incremental, active notebooks only
POST /sync/sweep       scoped sweep, active notebooks only
POST /sync/sweep/full  every section of every active notebook
POST /sync/sweep/all   every section of every mirrored notebook, active or not
```

The mode is the path rather than a body field or a query parameter, for the reason
`src/sync-route.ts` already gives: `src/logging.ts` records the path and deliberately
records neither of the other two, and "which job ran, and did it answer?" is the first
question when the mirror looks wrong.

`sweepCursorSectionId` stays a single field shared by all three sweep modes. A `sweep-all`
resuming onto a cursor left by a scoped `sweep` finds no match, `findIndex` returns -1, and
`Math.max(0, -1)` restarts from the top. That is the existing behaviour and it costs a
repeat rather than a wrong answer.

## 3. Re-activation

### The failure this prevents

Tier 1 of the incremental pass skips a section unless

```
section.graphLastModifiedDateTime >= overlapFrom(state.sectionsScannedThrough)
```

`sectionsScannedThrough` advances on every run that completes. A section edited three months
ago, while its notebook was inactive, has a `graphLastModifiedDateTime` far older than that
cutoff. Moving the notebook back to active would therefore change nothing: the section is
still older than the cutoff, so it is still skipped, permanently.

### The fix

`reconcileStructure` computes `activeSelectionHash` — a sha256 over the sorted active
notebook ids, or a fixed sentinel when `activeNotebookIds` is null — and compares it with
the value in sync state. When it differs, the run patches

```
sectionsScannedThrough: null
activeSelectionHash: <new value>
```

before candidates are picked. `overlapFrom(null)` is the epoch, so every mirrored active
section becomes a candidate on that one run. Each costs one `listPagesChangedSince` against
its own per-section watermark, which is intact, so only pages that genuinely changed are
fetched. The run is budget-bounded and resumable like any other.

The hash is computed over the active ids alone rather than over the whole selection, because
a change to `notebookIds` already moves the structure hash and rewrites the structure.

### What is deliberately not done

`active` is **not** stored on the notebook or section documents, and does not enter
`structureHashOf`.

`MirrorStore.putStructure` replaces documents wholesale — `#replaceCollection` calls
`batch.set` with no merge option — and `buildStructure` emits `pagesSyncedThrough: null` and
`pageCount: 0`. So anything that enters the structure hash resets every section's watermark
when it changes. Putting `active` there would make an activation edit trigger a full
re-backfill of the entire selection, which is hours of the request budget.

## 4. The read path

### The three source values

`MirrorSource` becomes `'onenote' | 'best-available' | 'mirror'`.

| value | claim |
|---|---|
| `onenote` | Everything returned is confirmed current with OneNote: it came from Graph, or from the mirror after a refresh that finished, or from a page document no write has marked stale. |
| `best-available` | Everything returned is either confirmed current, or from the mirror for a notebook the operator marked inactive. Treat it as current. |
| `mirror` | Some of it is from an **active** notebook and could not be confirmed on this call. It may be behind; `mirroredAt` says how far. |

`best-available` exists because without it an unscoped `search_pages` would report `mirror`
from the moment one notebook went inactive, for ever. `mirror` is the label for a degraded
answer — the refresh failed, or ran out of budget, or the scheduler held the lease — and a
result that skipped an inactive notebook on purpose is not degraded. Reporting the two the
same way would train a model to ignore the one that matters.

Note on the middle value: `mirror` today arises when the inline refresh did not finish, which
covers a Graph failure but also a budget that ran out and a lease held by the scheduler. It
does not mean "a Graph call threw" — a mirror miss falls through to Graph, and a Graph
failure there surfaces as a tool error, not as a `mirror` result.

### How it is decided

`SourcedRead` gains one optional member:

```ts
/** How much of this answer comes from notebooks the operator marked inactive. */
inactiveCoverage?(reader: MirrorReader): Promise<'none' | 'some' | 'all'>;
```

`readSourced` evaluates it only on a mirror hit, after `fromMirror` has answered. On
`origin: 'graph'` the source is `onenote` unconditionally and the member is never called.

```
confirmed = freshness === 'current' || staleTracked === true
coverage  = await inactiveCoverage?.(mirror) ?? 'none'

coverage === 'none' → confirmed ? 'onenote'        : 'mirror'
coverage === 'all'  → 'best-available'
coverage === 'some' → confirmed ? 'best-available' : 'mirror'
```

The `'all'` row is the one that looks wrong and is not. A `list_pages` on an inactive section
reports `best-available` even when the inline refresh failed outright, because that refresh
was never going to check that notebook. The claim rests on two things, neither of which
involves the refresh: the operator's assertion that the notebook is not edited, and the fact
that a write through this server marks its page stale or holds its section listing — and a
held listing is a mirror miss, which goes to Graph.

An error thrown by `inactiveCoverage` is caught and treated as `'some'`, the pessimistic
value: it cannot produce `onenote`, and it cannot produce `mirror` on an answer whose refresh
was fine. A failure to read the selection must not turn into a stronger claim than the data
supports, and must not fail the read.

`mirroredAt` is reported on `best-available` exactly as on `mirror`.

### What the reader needs

`MirrorReadStore` gains `getSelection(): Promise<NotebookSelection>`. `MirrorStore` already
has the method, so only the interface and the test fakes change.

`MirrorReader` gains:

```ts
notebookIsActive(notebookId: string): Promise<boolean>;
inactiveCoverageOfAccount(): Promise<'none' | 'some' | 'all'>;
```

Both read the selection through one memoised accessor. The memo holds for
`INLINE_SYNC_MIN_INTERVAL_MS` (30 s), matching the inline refresh's own interval; the worst
consequence of the memo is a source label 30 seconds behind an operator's edit, and the memo
is what stops every covered tool call adding a Firestore document read.

`inactiveCoverageOfAccount` compares the mirrored notebooks against the active set:
`'none'` when every mirrored notebook is active, `'all'` when none is, `'some'` otherwise.
It reads `listNotebooks()`, which the unscoped search path has already read.

### Per tool

| tool | `inactiveCoverage` |
|---|---|
| `list_notebooks` | not set — structure, unaffected |
| `list_sections` | not set — structure, unaffected |
| `get_page_content` | `'all'` when the page document's `notebookId` is inactive, else `'none'` |
| `list_pages` | `'all'` when the section's `notebookId` is inactive, else `'none'` |
| `list_pages_by_name` | as `list_pages`, on the resolved section |
| `find_page_by_name` | as `list_pages`, on the resolved section |
| `search_pages` with `sectionId` | as `list_pages`, on that section |
| `search_pages` unscoped | `inactiveCoverageOfAccount()` |

`MirrorPage` and `MirrorSection` both already carry `notebookId`, so no tool pays an extra
document read to answer this.

### Two additions to what a result says

- `list_notebooks` gains `pagesActive: boolean` on each entry from the mirror, beside the
  existing `pagesMirrored`. A model choosing where to look benefits from knowing which
  notebooks are kept current. Graph-origin entries carry neither field, as today.
- Unscoped `search_pages` gains `inactiveNotebooks: N` beside `notebooksSearched` and
  `notebooksInAccount`. `best-available` says a skip happened; the count says how large it
  was.

### The tool descriptions have to say what the values mean

Nothing in `src/structure-tools.ts` documents `source` to the calling model today, so
`best-available` would arrive as an unexplained string. A single exported constant — a
two-sentence note naming the three values — is appended to the description of each covered
tool. One constant rather than eight copies, so the descriptions cannot drift apart.

## 5. The write path

Unchanged. `resyncPage` consults `section.mirrored` and never activity. A write to an
inactive notebook:

1. calls `beginWrite`, marking the page stale and/or holding the section listing,
2. writes to OneNote,
3. resyncs the page into the mirror,
4. calls `endWrite`.

This gets an explicit test rather than being left implied, because it is the property most
likely to be broken by someone adding an activity check to `resyncPage` for symmetry.

A consequence worth stating: the mirror's copy of an inactive notebook is correct for every
edit made through this server, and stale for every edit made in the OneNote client. That is
the trade the operator makes by marking a notebook inactive.

## 6. Report and observability

`SyncReport` gains:

- `sectionsSkippedInactive: number` — sections a run declined to visit because their notebook
  is inactive.
- `unknownActiveNotebookIds: number` — active ids matching no notebook in the tree.

`MirrorSyncState` gains `activeSelectionHash: string | null` and
`unknownActiveNotebookIds: number`, with the defaults an unwritten document reads as.

New log event `mirror-selection-unknown-active`, carrying a count and nothing else.

## 7. Tests

`test/mirror-schema.test.ts`
- `readSelection` with `activeNotebookIds` absent, a non-array, `[]`, duplicated entries,
  entries needing trimming, and non-string entries.
- `isActive` for the null case and the listed case.

`test/mirror-sync.test.ts`
- A section in an inactive notebook with `pagesSyncedThrough === null` is a candidate, so the
  backfill runs.
- The same section once backfilled is not a candidate.
- The activity filter applies even when `sectionRollUpTrusted` is false or the tree read
  failed, which is the branch `pickCandidates` returns early from.
- A scoped sweep and a full sweep skip inactive sections; `sweep-all` visits them.
- Changing the active set nulls `sectionsScannedThrough` on the next run.
- `structureHashOf` does **not** move when only the active set changed, so no watermark is
  reset.
- `resyncPage` writes a page in an inactive notebook.
- `sectionsSkippedInactive` counts what was skipped.

`test/mirror-reader.test.ts`
- The `readSourced` truth table: all nine combinations of `'none' | 'some' | 'all'` against
  `confirmed` true and false, since that is the whole contract.
- `inactiveCoverage` throwing yields `'some'`, not a failed read.
- A `get_page_content` hit in an inactive notebook reports `best-available` despite
  `staleTracked`.
- A read scoped to an active notebook is unchanged from today.

`test/mirror-tools.test.ts`
- Each covered tool passes the coverage its table row says, including that `list_notebooks`
  and `list_sections` pass none.
- `pagesActive` and `inactiveNotebooks` appear where the design says.

`test/sync-route.test.ts`
- `POST /sync/sweep/all` reaches `runSweepAll` and no other mode.
- `GET /sync/sweep/all` is 405 with an `Allow: POST` header.
- The route is absent when `MIRROR_SYNC_SECRET` is unset.

`test/structure-tools.test.ts` and `test/page-tools.test.ts` are unchanged. They build tools
with no mirror argument, so every assertion in them still holds — which is the property that
keeps `MIRROR_READ_ENABLED=false` a rollback rather than a code path with its own bugs.

## 8. Documentation

- `CLAUDE.md` — a conventions block for the active set, one for the three source values, and
  one for why `active` is not in the structure hash.
- `README.md` — the fourth route in the schedule and runbook, and the selection document's
  new field.
- `project-spec.md` — the mirror section.

## Acceptance

1. A selection document with no `activeNotebookIds` produces byte-identical sync behaviour to
   today.
2. A notebook listed in `notebookIds` but not in a present `activeNotebookIds` is backfilled
   once and then visited by no incremental run, no scoped sweep and no full sweep.
3. `POST /sync/sweep/all` visits it.
4. Moving it back into `activeNotebookIds` causes the next incremental run to re-examine
   every section of it, without resetting any watermark elsewhere.
5. All five writing tools work against it, and the mirror reflects the write immediately.
6. A read confined to it reports `source: "best-available"`.
7. An unscoped `search_pages` reports `onenote` when no notebook is inactive,
   `best-available` when some are and the refresh finished, and `mirror` when the refresh did
   not finish.
8. `npm run typecheck` and `npm test` pass.
