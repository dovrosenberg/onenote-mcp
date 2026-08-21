# Mirror conventions

Read before changing anything under `src/mirror-*.ts`, `src/read-sync.ts`, or the sync
route. The mirror is the Firestore + GCS copy of the account's pages; `MIRROR_READ_ENABLED`
turns the read half off, which is the rollback switch.

## Reading through the mirror

**Every covered read is refresh, read, report — in that order.** `readSourced` in
`src/mirror-reader.ts` is the one place it is written. The refresh is the inline
incremental sync in `src/read-sync.ts`; it runs *before* the mirror is read, because a
refresh landing afterwards would have answered the question the caller already asked with
data it never saw. A miss still falls through to Graph, unchanged, which is what keeps a
notebook outside the mirrored selection readable.

**`source` is a claim about the answer, `origin` is a fact about its shape, and they are
separate fields for that reason.** `origin` is `mirror` or `graph` and decides which
branch of `Sourced<M, G>` a handler reads — `list_pages` from the mirror carries `total`
and from Graph carries an array, and a cast is exactly how a field that means nothing on
one path ends up in the other path's result. `source` is `onenote` or `mirror` and is the
only one of the two that reaches the caller: `onenote` when the answer equals what OneNote
holds, `mirror` when it may be behind, with `mirroredAt` saying how far. A mirror-origin
read carries either label, which is why one field cannot do both jobs.

**`source` is `onenote` in three cases and `mirror` in one.** Graph answered; or the
mirror answered and the refresh finished; or the mirror answered a single page nothing has
marked stale. Otherwise `mirror`. "Finished" is `freshnessOf` in `src/read-sync.ts`: the
outcome is `complete`, `done` is true, `treeRead` is true, and `pagesFailed` is zero. The
last two are the ones a simplification would drop and they are not decoration — a failed
expanded-tree read is deliberately non-fatal to a sync, so a run that carried on against
the structure already in Firestore returns `complete` while missing every notebook created
since the last good tree read; and a failed page fetch leaves that page's stored copy
exactly where it was.

**There is a third `source`, `best-available`, and it exists for one failure.** Without it
an unscoped `search_pages` would report `mirror` from the moment a single notebook was
marked inactive, for ever after. `mirror` is the label for a *degraded* answer — the
refresh failed, ran out of budget, or the scheduler held the lease — and an answer that
skipped a frozen notebook on purpose is not degraded. Reporting the two the same way would
train a model to ignore the one that matters. The whole table is `sourceFor` in
`src/mirror-reader.ts`: coverage `all` is `best-available` **whether or not the refresh
finished**, because that refresh was never going to check that notebook, so reporting its
failure would describe something that could not have changed the answer.

**A tool says how much of its answer is frozen through `inactiveCoverage`, and absent
means `none`.** It is evaluated only on a mirror hit, only after `fromMirror` answered,
and it is handed that answer — the two `_by_name` reading tools resolve a section id inside
`fromMirror`, and that id is the only thing saying which notebook the answer came from.
An error thrown by it is caught and treated as `some`, the pessimistic value: it cannot
produce `onenote`, and it cannot produce `mirror` on an answer whose refresh was fine.
`list_notebooks` and `list_sections` set it on neither, because structure is stored for the
whole account on every run whatever the active set says — weakening their label would
report a staleness they do not have.

**A selection the mirror could not read weakens a decoration and fails a label, and those
are opposite directions.** `ActivitySnapshot.selection` is nullable: the `getSelection`
read is caught inside the snapshot, so `pagesActive` and `inactiveNotebooks` degrade to
"frozen" and the answer beside them survives — a Firestore blip on the selection document
must not send an unscoped `search_pages` to Graph at 61 requests over a decoration.
`accountActivity` and `#coverageOfNotebook` read that null and answer `'some'` themselves
rather than treating it as "nothing is active", because `sourceFor` maps `'all'` to
`best-available` whatever the refresh did — which is a **stronger** claim than the `'some'`
`readSourced`'s own catch produces. A snapshot whose selection is null is dropped from the
memo, for the reason a rejected one is.

**The activity snapshot is memoised for `ACTIVITY_MEMO_MS`, and a failure is not cached.**
Every covered tool call asks about activity, so without the memo each would add a Firestore
document read to a request a person is waiting on. The promise is memoised rather than its
result, so two calls arriving together share one pair of reads; a rejection drops the entry,
because a cached rejection would pin every later answer to the pessimistic label for the
whole window over one Firestore blip.

**`staleTracked` is set on `get_page_content` and must not be set on anything else.** It
is what lets a page hit report `onenote` without a finished refresh, and the licence comes
from the page document rather than from the sync: every write calls `markPageStale` before
it touches OneNote, and `MirrorReader.getPageContent` answers a miss for anything but
`present`. The other six tools answer from *listings*, and a page created in a section this
process never wrote to leaves every stored document in that section untouched and still
correct-looking — so a listing has no per-document staleness to lean on and its label has
to come from the refresh.

**The refresh is rate-limited, and the interval is the cost control.** Without it a
conversation of thirty tool calls runs thirty syncs. A finished refresh holds for
`INLINE_SYNC_MIN_INTERVAL_MS` (30 seconds) and one that did not finish holds for
`INLINE_SYNC_RETRY_INTERVAL_MS` (5 minutes) — longer, not shorter, because a budget that
ran out, a lease that was held and a Graph that refused are none of them better in thirty
seconds, and retrying at the fast cadence would spend the hourly budget learning the same
thing repeatedly. Two callers arriving together share one run; the second would be refused
by the sync lease anyway, and sharing gives it the first one's answer rather than a
`behind` it did not earn.

**A refresh never fails a read, and never throws.** Graph refusing it, Firestore being
unreachable and the scheduler holding the sync lease all end as `behind` with a
`read-sync-skipped` log line. The tool still has its Graph fallback and the mirror still
holds whatever it held; the only consequence is the weaker label.

**There is no `useLiveData` argument, and adding one back would be a regression.** It
existed to force a read past a mirror that might be behind. A read that refreshes first
and then reports whether the refresh finished answers the same question without a schema
field the calling model has to reason about, and `mirror` in the result is a stronger
signal than an argument the model has to know to pass. The escape hatch for a notebook
outside the mirrored selection is `sectionId` on `search_pages`, which misses in the
mirror and walks that section in OneNote.

**The read branch lives in the tool modules, not behind an adapter.** An adapter
implementing `StructureClient` would return `PageSummary[]` with nowhere to say who
answered; reporting the source would need either state on an object `createTools` shares
across every request — wrong the moment two calls overlap — or a widening of every narrow
interface's return type, which would change every existing fake in three test files. And
"which source answered" is part of the tool's contract with the model, exactly like
`moreAvailable`, `stoppedEarly` and `deepSearchUsed`. `readSourced` in
`src/mirror-reader.ts` is the one place the branch is written; every covered tool calls it,
narrows on `origin` where the two shapes differ, and spreads `source` and `mirroredAt` into
its JSON.

**Every mirror read answers `null` on a miss, and every miss means "ask Graph".** A page
the mirror does not hold, a write-marked-stale page, and a Firestore outage all end in the
same place. Refusing a tool call because a cache is
down would be strictly worse than the behaviour before the mirror existed, which is the
bar. Four misses look like hits if you write them carelessly: a page whose stored ink
object is gone (answering `ink: null` claims the page has no handwriting), a section group
whose `childGroupsKnown` is false (answering a short list that looks complete), an
empty structure collection (answering "no notebooks" rather than "never synced"), and a
section whose `pendingWrites` is held (answering a page listing that a create or a rename
has already made wrong).

**`list_notebooks` and unscoped `search_pages` cannot miss, which is why structure is
mirrored for the whole account.** Neither takes an argument the mirror could fail to find,
so both would answer confidently and partially from a mirror holding three notebooks out
of fifty-five. The tree read returns every notebook and section for the same one request,
so storing all of it is free; `mirrored` on each document records which have their *pages*
held. Unscoped `search_pages` still reports `accountCoverage.searched` against
`accountCoverage.inAccount`, because its pages really are a subset — and it drops
`stoppedEarly` and the section counts, which describe a walk that did not happen.

**A scoped `search_pages` reports on the section it searched, and on nothing else.**
`accountCoverage` is `null` when the caller named a section, so the tool omits
`notebooksSearched` and `notebooksInAccount` and `mirrorSearchNote` drops the clause that
tells the caller to pass a `sectionId` it has already passed. `inactiveNotebooks` is
counted over the notebooks *in scope* for the same reason, and that one is not cosmetic: a
search scoped into an active notebook resolves `inactiveCoverage` through
`coverageOfSection`, reports `source: "onenote"`, and used to carry an account-wide frozen
count beside it — a staleness warning about a notebook the query never touched, next to a
claim that the answer is confirmed current.

**A `NameLookupError` from the mirror is a miss, not an answer.** The mirror's structure
equals Graph's expanded tree by construction but is only as fresh as the last sync, so a
section created ten minutes ago is absent — the inline refresh narrows that window but a
run that did not finish leaves it open — and `resolveSection` would turn that into
"sectionName matched nothing" listing the wrong siblings — which reads to a model as "no
such section". Both `_by_name` reading tools retry the whole resolve-and-list against
Graph before reporting. The retry costs one request and only on a failure, and a name that
exists nowhere still raises.

**The sweep enumerates page summaries, not page ids, and the extra two fields do two
jobs.** `listPageSummaries` asks for `$select=id,title,lastModifiedDateTime` on a request
the sweep was making anyway, so the fields are free. `title` is what a discovered page is
stored with — before this the sweep synthesized `''` and `1970-01-01T00:00:00.000Z`, and
both reach the calling model: `lastModifiedDateTime` is printed in every tool result and
`titleLower` is what by-name matching compares. Neither self-heals, because a page moved
into a section may not have its own timestamp bumped by the move, so no later incremental
lists it — measured 2026-08-21 and recorded in `api-overview.md`: a page moved between
sections keeps the stamp it had, so it is below any later watermark and **only the sweep
can find it**. `title` and `lastModifiedDateTime` together are what make `pageListingDiffers` in
`src/mirror-schema.ts` possible: a page in both places whose stamp or whose title differs
is re-fetched, which is the only thing in this repository that notices an edit or a rename
made in the OneNote client that the incremental pass missed.

**The sweep compares the title as well as the stamp, and the title half is not
redundant.** Measured 2026-08-21 and recorded in `api-overview.md`: a rename moves no
`lastModifiedDateTime`. The incremental never lists a page whose stamp is below the section
watermark, so the sweep is the only pass that ever looks at one — and a stamp-only
comparison there left a page renamed outside this server carrying its old title in the
mirror permanently, which is the field every listing and every by-name lookup matches on.
The reachable route needs nothing unmeasured: `update_page_title` calls `markPageStale`,
which leaves the stamp alone; the PATCH renames the page without moving it; and a
`resyncPage` that hit a transient failure is documented below as non-fatal. Both the sweep
and `storedPageIsCurrent` call `pageListingDiffers`, one predicate rather than two written
out, because a margin or a parse added to one copy would not reach the other and the
sweep's copy is where the omission has no upper bound on how long it lasts.

**A listing difference triggers a re-fetch, never a stale mark, and the difference between
those two is data loss.** `markPageStale` deletes the page-content document, and nothing
re-fetches a stale page — no read path writes to the mirror, the incremental will not list
a page whose Graph stamp is behind the section watermark, so a mark is permanent. Since
`resyncPage` stamps `new Date().toISOString()` from this process's clock after a write
returns, every page written through this server is stored *ahead* of Graph, and a
mark-on-difference would delete the mirrored content of exactly the account's most-used
pages. A re-fetch does the opposite: the page comes back with Graph's own string in place
of the local one, which is the repair. **Do not put a mark back in that branch.**

**The comparison has no margin, no direction test and no `Date.parse`, and the re-fetch is
what makes all three affordable.** A tolerance buys jitter-safety by discarding every real
edit inside its window, and an edit the sweep never notices is served as current for ever;
a false positive costs one Graph request. Direction stopped mattering once the response
became a re-fetch. Parsing stopped mattering for the same reason: two spellings of one
instant (`…:00Z` from Graph, `…:00.000Z` from `resyncPage`) disagree, they are re-fetched
once, and the short-circuit stores Graph's spelling — where a parsed compare would leave
the local one in a field every tool result prints. Graph's own one-second wobble on an
unchanged page, measured 2026-08-21, is absorbed the same way: one request, then agreement.
The pre-2026-08-21 documents carrying `title: ''` and `1970-01-01T00:00:00.000Z` need no
predicate of their own — the epoch disagrees with anything Graph sends.

`contentState` is **not** consulted, and the listing comparison is what covers for it. A
stale or missing copy whose stamp and title both still match Graph's listing is skipped,
not repaired. Nothing routinely produces one: every write that marks a page stale moves
either the stamp (`append_to_page`, and `create_page` has no document to mark) or the title
(`update_page_title`). The residue is a rename to a byte-identical title whose resync then
failed, which leaves a permanent `get_page_content` miss and one Graph request per call.
Closing it means projecting `contentState` into `listPageDigestsInSection`, which is a
`src/mirror-store.ts` change with no test available to it.

**The stamp is a hint; `writePageFromRaw`'s content hash is the check.** A re-fetch whose
content is identical writes no page document, so a false positive costs one Graph request
and one small Firestore write rather than a re-render and a blob upload.

**A short-circuited page write corrects the stored stamp, and without that the sweep never
converges.** `writePageFromRaw` compares the content hash, the title and the section and
deliberately does not compare `lastModifiedDateTime` — including it would rewrite every
page the watermark overlap re-read. But leaving the stored stamp alone means the next sweep
disagrees with Graph again and fetches the same page for ever. So the short-circuit calls
`MirrorStore.putPageMetadata`, which writes the metadata fields and refreshes
`contentSyncedAt` — the caller has just confirmed the content against Graph, and
`contentSyncedAt` is what `mirroredAt` reports to the model.

It fires on two conditions, not one, because it is a Firestore write per page and the
overlap re-reads every page edited in the last hour on every run. The stamps differing is
the first. The second is the stored copy failing the settle guard — `contentSyncedAt`
absent, which every page document written before that field existed is, or within
`TIMESTAMP_SETTLE_MS` of Graph's stamp. Without that second condition the guard had no
exit: `storedPageIsCurrent` refused the copy, the content was fetched and found identical,
the stamps agreed so nothing was written, `contentSyncedAt` never moved, and the guard
refused it identically on the next run — one Graph content request per run per page for
the whole hour-wide listing window, spent on the freshest pages in the account, with
`pagesUpdated` and `pagesSkipped` both reading zero and only `graphRequests` moving. The
exit is bounded rather than immediate: a page fetched five seconds after its stamp settles
to a `contentSyncedAt` five seconds after it, and the *next* run's write puts it at that
run's clock, past the margin. `contentCopyIsSettled` in `src/mirror-schema.ts` is the one
place the guard is written, so the skip and this write cannot disagree about it.

A live stamp of `''` never overwrites a stored one: that is `toPageSummary`'s fallback for
a field Graph did not send, not a timestamp. Such a page fails the settle guard for ever
and is re-read once per run — the documented price of never hiding an edit, which no page
Graph stamps ever pays.

**The sweep's reconciliation loop checks the request budget, and that check now guards
Graph requests rather than only Firestore writes.** A disagreement costs a content fetch,
so a loop that ignored the budget could spend one per mirrored page in a section, against
400 an hour shared with every interactive tool call.

**An inactive notebook is backfilled once and then never re-listed.** `activeNotebookIds`
in the selection document names which mirrored notebooks a sync still re-checks; absent or
malformed means all of them, and `[]` means none — which is why `NotebookSelection` holds
`readonly string[] | null` rather than an array that happened to be empty. `splitByActivity`
in `src/mirror-sync.ts` is the one place the filter is written, and `includeBackfill` is
what makes the backfill run exactly once: a section with `pagesSyncedThrough === null` is
eligible whatever its notebook's activity. A sweep passes false, because a sweep reconciles
a section that is already filled rather than filling one. `sweep-all` skips the filter entirely and
is the only thing that reaches a frozen notebook, which is what it is for.

**`active` stays out of `structureHashOf` and off the section documents, and that is not
tidiness.** Activity is not a field on any structure document, so a structure pass has
nothing to say about it: the hash it moved would name no document to write, and
`planStructureWrite` would skip every one of them on an unchanged identity. The two changes
also need different responses — a structure change is answered by the narrower rules listed
under `mayFilterByTimestamp` below and widens nothing, and an activation change widens the
scan for the notebooks it named. So `active` is recorded separately, in
`activeNotebookIdsSeen`, and `reconcileSelection` acts on it.

**A selection change widens the scan for the notebooks it named, and for no others.**
Tier 1 of `pickCandidates` skips a section whose `graphLastModifiedDateTime` is older than
`overlapFrom(state.sectionsScannedThrough, SECTION_SCAN_OVERLAP_MS)`, and that cutoff
advances on every completed run — so a notebook just added to `notebookIds`, or just added
to `activeNotebookIds`, has sections months older than the cutoff that nothing else would
ever make candidates again.
`notebooksNeedingWideScan` in `src/mirror-schema.ts` diffs the two lists the state document
recorded against the two the selection document now holds, and returns only the ids that
became mirrored or became active; `pickCandidates` bypasses the cutoff for those notebooks'
sections and nothing else. The mechanism this replaced nulled `sectionsScannedThrough`,
which is one global value: activating one notebook made every mirrored active section a
candidate, about 70 listing requests on this account against an hourly budget of 400 for a
change that concerned one notebook. Nothing here touches `sectionsScannedThrough` any more.

**`mayFilterByTimestamp` is `treeRead` alone, and the `!rewritten` it used to carry was the
same bug through a second door.** `notebookIdentity` carries `mirrored`, so editing
`notebookIds` moves the structure hash — and `pickCandidates` returned every observed
section on a moved hash, before the wide-scan clause was evaluated. Each reason a structure
change used to need a wide pass now has a narrower rule: a section the tree just gained is
created with `pagesSyncedThrough: null`, a renamed one keeps a watermark no structure write
touches and a rename changes no page, and a section whose notebook just became mirrored or
active is the wide-scan set's job.

**A section moved between notebooks brings no watermark with it, and that is measured.**
Measured 2026-08-21 and recorded in `api-overview.md`: the OneNote client reissues a
section's Graph id when the section changes notebook, and keeps the id when the section is
only reparented under a section group in the same notebook. So a cross-notebook move leaves
the old id absent from the next tree read, its document is deleted by absence, and the new
id creates a document with `pagesSyncedThrough: null` — which is backfill-eligible whatever
the tier-1 cutoff says and whatever its notebook's activity says. The class the
`pickCandidates` docstring used to record as knowingly uncovered, a section carrying arrears
into a mirrored active notebook, is unreachable. The cost that replaces it is a full
re-backfill of that section: every page fetched again under its new id, every page document
under the old id deleted. Correct and self-healing, and a section's worth of requests
against the hourly budget, so it is visible in a run report rather than free.

Only *candidacy* is widened. Each named notebook's sections still list against their own
`pagesSyncedThrough`, which no part of this touches, so a widened section costs one
`listPagesChangedSince` and re-fetches only the pages that actually changed. The set is
stored in `wideScanNotebookIds` rather than held for one run, because a run is
budget-bounded and may stop with sections outstanding; it is cleared in the same block that
advances `sectionsScannedThrough` and on exactly that condition, because a run stopped by
its budget has not visited what it was widened for. `sweepPass` reads the set and never
clears it — a sweep reconciles a section's pages rather than resuming a watermark, so it is
not what the widening is waiting for. A run that read no tree records nothing at all:
`mirroredNotebookIds` is empty there, so an `activeNotebookIds` of `null` would resolve to
"nothing is active", widen nothing, and still record `activeNotebookIdsSeen: null` — and
the next healthy run would diff `null` against `null` and never widen the notebook the
operator just unfroze.

Removal widens nothing and is still recorded, and those are two separate rules.
`notebooksNeedingWideScan` returns `[]` for a notebook dropped from either list, because
there is no backlog to catch up on. `selectionMatchesSeen` is what decides whether to
write, and it is a different question: a run that wrote only when it had something to widen
would leave the deactivated notebook in `activeNotebookIdsSeen`, and re-activating it later
would diff against a list it is already in and widen nothing. A state document written
before these fields existed carries `mirroredNotebookIdsSeen: null`, and both functions key
on that one field: the diff answers `[]` and the match answers false, so such a run records
the lists and widens nothing. There is no separate flag for it — treating it as a change
would make the first run after the deploy a full-width scan, and a boolean saying which
null it was would be a second answer to a question `mirroredNotebookIdsSeen` already
settles.

**The sweep's resume cursor names the section the budget stopped inside, not the one after
it.** `sweepSection` returns false when it ran out part-way through a section's pages, and
`sweepPass` writes that section's own id. Recording the *next* one — which is what a cursor
written only at the top of the loop does — leaves the interrupted section half reconciled,
and nothing returns to it until the cursor reaches the end of the list and resets, several
runs later. It is the section that most needed the visit, because the budget ran out inside
it. A failed enumeration returns true instead: it reconciled nothing, it is logged, and the
sweep carries on past it exactly as before.

**A section whose page listing failed holds `sectionsScannedThrough` where it is, and that
is `sectionListingFailed` rather than `done`.** The two are separate because `done: false`
is what makes `runMode` report `budget-exhausted`, and a 429 on one listing is not that —
reporting it as that sends the scheduler to retry immediately and spend the next hour's
Graph budget inside this one. What they share is the consequence: the run did not scan
every candidate, and the cutoff only moves forward, so advancing past the missed section
leaves it one `SECTION_SCAN_OVERLAP_MS` window to be retried in and then drops it from the
candidate set for good — watermark stuck behind its real edits until the nightly full
sweep. Holding the cutoff costs the next run a re-list of the sections inside the window.

**There are two overlap windows and they are different widths on purpose.**
`WATERMARK_OVERLAP_MS` (an hour) is how far back a *section's page listing* reaches beyond
that section's `pagesSyncedThrough`; a page it surfaces that has not changed costs no Graph
request, because `storedPageIsCurrent` skips it from the listing's own stamp and title, so
the width buys margin and spends bytes. `SECTION_SCAN_OVERLAP_MS` (fifteen minutes) is how
far back tier 1 of `pickCandidates` reaches beyond `sectionsScannedThrough`; every section
it surfaces costs one `listPagesChangedSince` on every run for as long as the window lasts.
Do not collapse them back into one constant.

**The `sync-overlap-save` line measures the hour, not the fifteen minutes.**
`overlapSaveAgeMs` is called with `section.pagesSyncedThrough`, so its `ageMs` is bounded
by `WATERMARK_OVERLAP_MS`; an `ageMs` approaching 3600000 says the hour is too narrow, and
one in the tens of minutes is ordinary. **Nothing observes whether the section-scan window
is wide enough.** A section that window declines is never a candidate, so `syncSection`
never runs, no page listing is made, and no line is written at all — the failure is silence
rather than a large number. The nightly `/sync/sweep/full` is the backstop: it compares
stamps and titles against every stored page whatever the timestamps say, so a section the
15-minute window declined is reconciled by morning. `README.md` carries the query under
**Proving it works**, beside the `graph-clock-skew` one.

**Activity never touches the write path.** `resyncPage` consults `section.mirrored` and
nothing else, so a write to a frozen notebook still marks the page stale, still reaches
OneNote, and still updates the mirror. Adding a symmetrical activity check there for tidiness
would leave every write to a frozen notebook serving pre-write content until a `sweep-all`
ran, which may be never. The consequence worth stating: a frozen notebook's mirror copy is
correct for every edit made through this server and stale for every edit made in the OneNote
client, and that is the trade the operator makes by freezing it.

**Every write resyncs its page immediately, and that costs one Graph request.** All five
writing tools call `resyncPage` after a successful write — including `create_page`, whose
page the mirror has never seen and which would otherwise be a miss until the next
scheduled run. The resync is also what ends the section's listing hold with the listing
correct rather than merely un-held. The alternative, marking the page stale and letting the next sync repair
it, leaves every read falling through to Graph for up to a whole poll interval, which in
the middle of a conversation is the window that matters most.

**A resync re-reads content and nothing else, and this is the measured reason.**
`api-overview.md` records that a PATCH *is* visible to the next content read — 3.7 seconds
including both round trips — while page *metadata* is weaker: `GET /pages/{id}?$select=title`
returned `""` for pages created seconds earlier. So the title travels in a hint from the
caller, which either just set it (`update_page_title`, `create_page`) or knows an append
cannot change it, and `lastModifiedDateTime` is stamped locally. Reading either back from
Graph here would trade a correct value for an unreliable one. Do not "improve" this by
fetching page metadata.

**`writePageFromRaw`'s short-circuit compares the title and the section, not just the
content hash, and both are load-bearing.** `update_page_title` changes a title and nothing
else, so a content-hash-only comparison short-circuited every rename: the mirror kept
serving the old title, which `find_page_by_name` and `search_pages` then matched against.
A page moved between sections is the same shape of miss — page ids are stable across a
move, so only the placement changed. `lastModifiedDateTime` is deliberately **not**
compared: it moves on every write, so including it would rewrite every page the watermark
overlap re-read and defeat the short-circuit entirely.

**An append or create that resyncs to `unchanged` is treated as a lost race, and the page
is marked stale.** Those two always change a page's content, so a resync that found
nothing to write did not read what was just written. Measured 2026-08-19, a PATCH is
visible to the next content read at 3.7 seconds — but that is one observation, and if the
read ever loses the race the stored copy is pre-write content marked `present`, which the
read path serves as current with nothing saying so. A stale marker sends the next read to
Graph, which cannot be wrong. A rename does not fall through this branch, because it
changes no content by design.

**There is one page writer, `writePageFromRaw`, shared by the sync and the resync.** A
second copy that skipped the ink render, or spilled to GCS at a different threshold, would
make a page's stored form depend on which path last touched it — and the difference would
only surface as a wrong answer to a model days later. `test/mirror-sync.test.ts` asserts
both paths build the same document from the same response.

**Every write is three steps in this order: invalidate the mirror, write to OneNote,
resync the mirror.** All five writing tools call `beginWrite` before they touch Graph.
That ordering closes a window nothing else can. Between the PATCH succeeding and the
resync completing, OneNote and the mirror disagree. If the process merely errors, the
resync's `catch` marks the page stale. If the process **stops** — Cloud Run cutting the
request at 300 seconds, or the instance being reclaimed after an idle period — nothing
catches anything: the resync is sitting in the request gate's queue, the queue goes with
the instance, and no `catch` runs because nothing threw. The mirror would keep serving
superseded data as `present`, reporting `source: "mirror"` and a recent `mirroredAt`,
until the next scheduled sync noticed the page's `lastModifiedDateTime` had moved.
Invalidating first makes the whole window pessimistic: a death anywhere in it leaves a
miss, and a miss goes to Graph.

**There are two things to invalidate, and which ones a tool marks follows from what it can
make wrong.** `beginWrite` takes a `pageId`, a `sectionId`, or both.

| | page content (`markPageStale`) | section page listing (`holdSectionListing`) |
|---|---|---|
| `append_to_page`, `append_to_page_by_name` | yes | no |
| `update_page_title` | yes | yes |
| `create_page`, `create_page_by_name` | n/a — no page document to mark | yes |

The page marker covers `get_page_content`. The listing hold covers `list_pages`,
`list_pages_by_name`, `find_page_by_name` and `search_pages`, every one of which answers
from stored page documents rather than from page *content*. A create has no page to mark
and the listing is exactly what it makes wrong: without the hold, a create whose resync
never ran leaves `list_pages` answering from the mirror with a section that does not
contain the page just created, which reads to a model as "the page was not created". A
rename holds the listing because the title is what every listing and by-name lookup
matches on — this is the same failure `writePageFromRaw`'s title comparison exists to
prevent, reached through a different door. An append holds nothing extra: it changes
content, which the page marker covers, and `lastModifiedDateTime`, which only moves a page
within an ordering. Keeping appends out of the hold is what keeps the common write from
sending an unscoped `search_pages` to Graph at 61 requests.

**The listing hold is a count, and it expires on age.** A count, `pendingWrites` on the
section document, because two writes against one section can overlap and the first to
finish must not clear the second's hold. `endWrite` lowers it from a `finally`, so a Graph
write that failed releases too — the listing it was going to invalidate is still correct.
It expires because a `finally` does not run when the process stops, and a hold nothing can
lower would send every listing for that section to Graph forever. `LISTING_HOLD_EXPIRY_MS`
is ten minutes, longer than Cloud Run's 300-second request ceiling, so it can never fire
on a write still in progress. `listingIsHeld` reads a count with no timestamp as *expired*
rather than as held, for the reason `leaseIsHeld` reads an unparseable `runningSince` that
way: the alternative is a state nothing can ever clear.

**`markPageStale` and both listing-hold writes use `update`, not a merging `set`.** A
merging set creates the document when it is absent, and the write tools reach the whole
account while only the selection is mirrored — so every write to an unmirrored page would
leave a stub in the queried collection carrying nothing but `contentState: 'stale'`. On
the section it is worse: a stub section document carrying nothing but a counter is what
`listAllSections` then feeds to `expandedTree` as a section with no name. `NOT_FOUND` is
swallowed in both, because "there was no copy to invalidate" is the ordinary answer for a
page or section outside the selection.

**A failed invalidation never fails the write.** `beginWrite` and `endWrite` log
`mirror-invalidate-failed`, `mirror-listing-hold-failed` or `mirror-listing-release-failed`
and carry on. Both narrow an existing window; neither is a precondition for writing, and a
tool that refused to write because Firestore was down would be strictly worse than the
behaviour before the mirror existed.

**Two failure levels below a write, and neither fails the write.** A resync that throws
falls back to `markPageStale`, which makes the next read a miss — correct, just slower. If
that fails too, the event is logged and nothing else happens. The write has already
happened by then, so turning either into a reported error would send the caller to retry a
change that is already made. An append or a create is self-healing regardless: the write
moved the page's `lastModifiedDateTime`, so the next incremental run repairs whatever this
could not. A **rename** is not — measured 2026-08-21, `PATCH /pages/{id}/content` replacing
a title moves no stamp, so no later incremental lists the page and the sweep is what
repairs it, on the scheduler's own cadence rather than within the poll interval.

**The write-sync is bound whenever a mirror exists, not only when reads are enabled.** A
mirror being filled by the sync while `MIRROR_READ_ENABLED` is false still holds copies a
write supersedes, and keeping them current then is what makes turning reads on later safe
rather than a race with whatever was written in between. It shares the Graph content
client with the read tools, so a resync passes through the same process-wide request gate
as everything else — a burst of writes cannot outrun the per-user rate limit through this
path.

**An overrunning request either freezes or vanishes, and the lease release is written for
the first.** The service runs with CPU throttling — the default, unset in the deploy — and
`minScale` is unset, so a request cut at the 300-second timeout has two possible fates and
Cloud Run guarantees neither. If the instance is still warm the process is *suspended* and
resumes whenever the next request arrives, possibly many minutes later; if the instance has
scaled to zero it is gone. The scheduled sync is what makes the first case likely rather
than theoretical, because it wakes the instance on its own interval.

In the warm case the sync lease has expired on age by the time the run resumes, and another
run has taken it. An unconditional `releaseLease` would then clear the *live* run's lease
and let a third start alongside it, both spending the same hourly Graph budget — the exact
failure the lease exists to prevent, reached by way of the lease itself. So
`releaseLease(heldSince)` runs in a transaction and clears nothing unless `runningSince`
still matches the value that run wrote. A superseded run logs `mirror-lease-superseded` and
leaves the document alone.

## The sync route

**`POST /sync` is the page mirror's way in, on the same terms as `/keepalive` and for the
same reason.** Its own secret rather than the keepalive one, because the two reach
different things and a credential should reach one of them. Unmounted when
`MIRROR_SYNC_SECRET` is unset, so the path 404s and an operator learns the service is
unconfigured rather than that they mistyped a secret.

**The sync's mode is the path — `/sync`, `/sync/sweep`, `/sync/sweep/full`,
`/sync/sweep/all` — not a body field or a query parameter.** `src/logging.ts` records the method, the path and the
status, and deliberately records no query string and no body. A mode carried in either
would appear in no log line, and "which job ran, and did it answer?" is the first
question when the mirror looks wrong. A body would also need a JSON parser on a route
outside the bearer gate. A time-based rule would be worse than both: it makes behaviour
depend on the container's clock and removes the ability to force a sweep on demand, which
is the move the keepalive runbook documents as the way to prove a job works.

**A budget-exhausted sync run answers 200, not 503.** It is a normal outcome with
committed work behind it and a report attached. A 503 makes the scheduler retry
immediately and spend the next hour's Graph budget inside this one, which is the failure
the budget exists to prevent. 503 is reserved for a run that could not start; a held
lease is 409.

## Blobs

**The mirror's blobs are in GCS, and the InkML is stored beside the PNG.** Firestore
caps a document at 1 MiB and `MAX_INK_PNG_BYTES` is already 750 KB, so a rendered PNG
would fill most of a document alone. The InkML going to the bucket too is not redundant:
`MAX_INK_PNG_BYTES` is documented above as a chosen number rather than a measured one,
and `fitInkToByteBudget` shrinks a render by re-rasterising and measuring — so with only
the PNG kept, changing that budget would mean re-fetching every inked page from Graph,
hours of the request budget to correct a guess. It is the same protection the raw HTML
gets by being stored untrimmed. `src/mirror-blobs.ts` treats a 404 on read as a mirror
miss and a 404 on delete as success, and lets everything else propagate: a permission
failure that read as "not mirrored" would send every request to Graph and exhaust the
hourly budget with nothing saying why.
