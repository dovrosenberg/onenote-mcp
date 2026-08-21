# Mirror Sync Graph-Call Reduction Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers-extended-cc:subagent-driven-development (recommended) or superpowers-extended-cc:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make the mirror's incremental sync notice OneNote-client edits again, stop a change to one notebook from starting work on every other notebook, and cut the Graph requests one run spends from "one listing plus one content fetch per page edited in the last hour" to "one listing plus one fetch per page that actually changed".

**Architecture:** Eight changes to the sync, in dependency order. First a correctness fix — the section timestamp `pickCandidates` filters on is read from Firestore and never refreshed, so the incremental sync currently goes blind about an hour after the last structure write. Then two changes that confine the blast radius of an edit to the notebook it names: structure writes stop resetting sync-owned state, and the widening that re-activation needs becomes per notebook instead of global. Then a backstop so a wrong margin is recoverable within a night, telemetry to turn chosen constants into measured ones, the saving itself — skip the content fetch when Graph's stamp has not moved — and finally the section-scan overlap drops from one hour to fifteen minutes.

**Tech Stack:** TypeScript run directly under `node --test` (native type stripping, so the source must be erasable), Firestore, GCS, Microsoft Graph OneNote endpoints.

---

## Background an implementer needs

Read these before starting. They are short and every task depends on them.

**The clock rule.** Three clocks are in play: Graph's (stamps `lastModifiedDateTime`), this process's `Date.now()` (stamps watermarks, `sectionsScannedThrough`, leases, and `resyncPage`'s locally generated value), and Firestore's `serverTimestamp()` (stamps `contentSyncedAt`). A comparison whose two sides come from the **same** clock may use equality. A comparison that crosses clocks needs a margin. Every new predicate in this plan is labelled with which case it is. Do not "simplify" an equality into a tolerance: on a same-clock comparison a tolerance silently drops real edits, and the watermark then advances past them permanently.

**Graph's stamps are whole seconds.** `api-overview.md` records `2026-08-19T19:32:39Z`, `19:32:48Z`, `19:32:57Z` from the section roll-up probe. That resolution is why Task 7's skip needs a settle guard as well as an equality test: an edit landing in the same wall-clock second as our content fetch would not move the string.

**The watermark is the time the pass started**, not the newest thing it saw (`src/mirror-sync.ts:733`). That already absorbs the run's own duration, so the overlap window is covering only Graph's propagation lag plus clock skew.

**Blast radius is a requirement, not a preference.** After Tasks 2 and 3, adding or removing one notebook from `notebookIds` or from `activeNotebookIds` must cause **zero** Graph requests and **zero** Firestore writes against any other notebook, and must reset no other notebook's watermark. Every acceptance criterion in those tasks is written to be checkable against that sentence.

**Where things live.** `src/mirror-schema.ts` holds pure functions and types and touches no backend, and it has **no imports at all** — keep it that way, and define structural interfaces rather than importing types from `src/graph-structure.ts`. `src/mirror-store.ts` and `src/mirror-blobs.ts` hold only Firestore/GCS calls and have no automated test; anything a person could get wrong belongs in `mirror-schema.ts`. `src/mirror-sync.ts` holds the algorithm and is tested through fakes that count their calls.

**Conventions that will bite you.** Import specifiers carry the `.ts` extension. Type-only imports must be written `import type`. No `enum`, no `namespace`, no constructor parameter properties. `exactOptionalPropertyTypes` is on, so assigning an explicit `undefined` to an optional property is a type error — use a conditional spread. `npm run typecheck` is the static-analysis gate; there is no ESLint.

---

## File Structure

| File | Responsibility | Change |
|---|---|---|
| `src/mirror-schema.ts` | pure types, constants and predicates | Identity helpers, tree/sync field split, `notebooksNeedingWideScan`, `PageStamp`, `MirrorPageDigest`, `pageHasDrifted`, `storedPageIsCurrent`, `TIMESTAMP_SETTLE_MS`, `SECTION_SCAN_OVERLAP_MS`, new sync-state fields |
| `src/mirror-sync.ts` | the sync algorithm | `reconcileStructure` returns live timestamps; `withLiveMtimes`; `reconcileSelection`; `pickCandidates` takes a wide-scan set; page-level skip; sweep drift check; overlap-save telemetry |
| `src/mirror-store.ts` | Firestore calls only | `#replaceCollection` merges and skips unchanged documents; `listPageDigestsInSection`; `putPageMetadata` also stamps `contentSyncedAt` |
| `src/graph-structure.ts` | Graph URLs and decoding | `listPageIds` becomes `listPageSummaries` with a wider `$select`; skew probe in `graphGet` |
| `src/graph-throttle.ts` | the request gate, retry policy | New `recordClockSkew` |
| `src/page-content.ts`, `src/page-write.ts` | the other two fetch sites | Call `recordClockSkew` |
| `test/mirror-sync.test.ts` | the algorithm | Faithful `putStructure` fake; regression, blast-radius and skip tests |
| `test/mirror-schema.test.ts` | the pure predicates | Identity, wide-scan and both page predicates |
| `test/graph-structure.test.ts` | exact-URL assertions | Updated URL for `listPageSummaries` |
| `test/graph-throttle.test.ts` | the gate and retry policy | New tests for `recordClockSkew` |

---

### Task 1: Live section timestamps reach `pickCandidates`

**Goal:** The incremental sync filters sections on the timestamp this run's tree read observed, not on a copy frozen in Firestore since the last structure rewrite.

**Why this is first:** `pickCandidates` (`src/mirror-sync.ts:686`) compares `section.graphLastModifiedDateTime` — read from Firestore via `listSectionsToSync` — against a cutoff that advances on every completed run. That field is written only by `buildStructure`, whose output reaches Firestore only through `putStructure`, which is guarded by a hash that deliberately excludes timestamps (`src/mirror-sync.ts:573`, `structureHashOf` at `:313`). So the stored value freezes and the cutoff marches past it. Once `sectionsScannedThrough` exceeds `frozen_mtime + overlap`, the section is never a candidate again and no OneNote-client edit is ever mirrored. Nothing repairs it: there is no scheduled incremental in steady state, and the nightly `/sync/sweep/full` reconciles page **ids** only.

**Files:**
- Modify: `src/mirror-sync.ts` — `reconcileStructure` (`:534-577`), `incrementalPass` (`:490-524`), `sweepPass` (`:999-1037`); add `withLiveMtimes`
- Modify: `test/mirror-sync.test.ts` — `fakeStore`'s `putStructure` (`:229-232`), `harness` (`:303`)

**Acceptance Criteria:**
- [ ] `reconcileStructure` returns the live section timestamps whether or not it rewrote the structure
- [ ] `incrementalPass` and `sweepPass` both overlay them before calling `pickCandidates`
- [ ] The test fake's `putStructure` replaces `data.sections` and `data.groups` the way Firestore's does
- [ ] A regression test drives three runs with a stable structure and asserts an edit on the third run is listed and fetched
- [ ] The eight tests listed in Step 3 pass again

**Verify:** `npm run typecheck && npm test` → 0 failures

**Steps:**

- [ ] **Step 1: Write the failing regression test**

Add to `test/mirror-sync.test.ts`, after the existing `pickCandidates` tests:

```ts
test('an edit is still noticed after the structure has stopped changing', async () => {
  // The failure this guards: `graphLastModifiedDateTime` is written only by
  // `putStructure`, which runs only when the tree hash moves — and the hash excludes
  // timestamps. A stored copy therefore freezes while `sectionsScannedThrough` advances
  // on every completed run, and the section stops being a candidate for ever.
  let sectionMtime = '2026-08-19T09:00:00Z';
  let pageMtime = '2026-08-19T09:00:00Z';

  const tree = (): ExpandedNotebook[] => [
    {
      id: NB,
      displayName: '2026',
      lastModifiedDateTime: sectionMtime,
      sections: [{ id: 'sec-1', displayName: 'Daily', lastModifiedDateTime: sectionMtime }],
      sectionGroups: [],
    },
  ];

  const h = harness({ tree: tree() }, { sections: [], state: initialSyncState() });

  // The scripted tree is answered once; re-point both calls at the live values.
  h.deps.graph.getExpandedTree = () => Promise.resolve(tree());
  h.deps.graph.listPagesChangedSince = (sectionId, since) => {
    h.graphCalls.changedSince.push(sectionId);
    return Promise.resolve(
      pageMtime >= since ? [{ id: 'p1', title: 'Page', lastModifiedDateTime: pageMtime }] : [],
    );
  };

  let clock = Date.parse('2026-08-19T10:00:00Z');
  const deps = { ...h.deps, now: () => clock };

  await runIncremental(deps, { requestBudget: 100 });

  // A quiet run two hours later. Nothing changed, and the cutoff advances past the
  // section's stored timestamp.
  clock += 2 * 3600_000;
  await runIncremental(deps, { requestBudget: 100 });

  // Two hours after that, someone edits the page in the OneNote client.
  clock += 2 * 3600_000;
  sectionMtime = '2026-08-19T14:00:00Z';
  pageMtime = '2026-08-19T14:00:00Z';

  const listedBefore = h.graphCalls.changedSince.length;
  const report = await runIncremental(deps, { requestBudget: 100 });

  assert.ok(
    h.graphCalls.changedSince.length > listedBefore,
    'the section whose timestamp moved must be listed again',
  );
  assert.equal(report.pagesUpdated, 1, 'the edit must reach the mirror');
});
```

- [ ] **Step 2: Make the store fake faithful**

Replace `fakeStore`'s `putStructure` in `test/mirror-sync.test.ts:229-232`:

```ts
    // Firestore's `putStructure` replaces the collections wholesale — `#replaceCollection`
    // calls `batch.set` with no merge — so a fake that only counted calls could not see a
    // stored section's timestamp go stale, which is the failure Task 1 fixes.
    putStructure: (structure) => {
      calls.structures += 1;
      data.sections = [...structure.sections];
      data.groups = [...structure.sectionGroups];
      return Promise.resolve();
    },
```

- [ ] **Step 3: Run the suite and see what the faithful fake breaks**

Run: `npm test`
Expected: the new test fails, plus these eight, because they supply sections the scripted tree does not describe and the default empty tree now wipes them:

```
a changed page is fetched and stored, and the watermark moves to the pass start
an unchanged content hash writes nothing and renders no ink
an incremental run backfills an inactive notebook and then stops listing it
the activity filter still applies when the timestamps cannot be trusted
a scoped and a full sweep skip inactive sections; sweep-all visits them
the sweep does not advance a section watermark
a budget-exhausted sweep records where to resume
a scoped sweep visits only moved sections; a full sweep visits all of them
```

- [ ] **Step 4: Give `harness` a tree derived from the store**

Add `Section` to the existing `import type { ... } from '../src/graph-structure.ts'` clause at the top of `test/mirror-sync.test.ts`, then add above `harness`:

```ts
/**
 * The expanded tree that would produce `data`'s sections and groups.
 *
 * A test that scripts no tree used to get `[]`, which was harmless while the store fake
 * ignored `putStructure`. With a faithful fake an empty tree deletes every section, so
 * the default has to describe what the store was seeded with.
 */
function treeFrom(data: StoreState): ExpandedNotebook[] {
  // A conditional spread rather than `?? undefined`: `exactOptionalPropertyTypes` is on,
  // so writing `lastModifiedDateTime: undefined` on an optional property is a type error.
  const asGraphSection = (s: MirrorSection): Section => ({
    id: s.id,
    displayName: s.displayName,
    ...(s.graphLastModifiedDateTime === null
      ? {}
      : { lastModifiedDateTime: s.graphLastModifiedDateTime }),
  });

  return [
    {
      id: NB,
      displayName: '2026',
      sections: data.sections.filter((s) => s.parentKind === 'notebook').map(asGraphSection),
      sectionGroups: data.groups.map((g) => ({
        id: g.id,
        displayName: g.displayName,
        sections: data.sections
          .filter((s) => s.parentKind === 'sectionGroup' && s.parentId === g.id)
          .map(asGraphSection),
      })),
    },
  ];
}
```

Then replace `harness` entirely, so the graph fake is built after the store and can default its tree from it:

```ts
function harness(
  script: GraphScript = {},
  storeInit: Partial<StoreState> = {},
  content?: SyncDeps['content'],
): Harness {
  const { store, calls: storeCalls, data } = fakeStore(storeInit);
  const { graph, calls: graphCalls } = fakeGraph({
    ...script,
    tree: script.tree ?? treeFrom(data),
  });
  const { blobs, calls: blobCalls } = fakeBlobs();

  return {
    graphCalls,
    storeCalls,
    blobCalls,
    data,
    deps: {
      graph,
      store,
      blobs,
      now: () => NOW,
      content: content ?? { fetchRaw: () => Promise.resolve(rawHtml('<p>typed</p>')) },
    },
  };
}
```

Run: `npm test`
Expected: the eight listed tests pass again; the new regression test still fails, because the production bug is real.

- [ ] **Step 5: Return the live timestamps from `reconcileStructure`**

In `src/mirror-sync.ts`, add above `reconcileStructure`:

```ts
/** Section `lastModifiedDateTime` as one tree read saw it, by section id. */
export type SectionMtimes = ReadonlyMap<string, string | null>;

/** What one structure pass learned. Empty timestamps mean the tree read failed. */
interface StructureResult {
  /** True when the structure documents were written. */
  readonly rewritten: boolean;
  readonly liveMtimes: SectionMtimes;
  /** Ids of the notebooks the selection covers, for `reconcileSelection` in Task 3. */
  readonly mirroredNotebookIds: readonly string[];
}
```

Change `reconcileStructure`'s signature and its three returns:

```ts
async function reconcileStructure(ctx: PassContext): Promise<StructureResult> {
  const none: StructureResult = {
    rewritten: false,
    liveMtimes: new Map(),
    mirroredNotebookIds: [],
  };

  if (ctx.budget.exhausted) {
    ctx.tally.done = false;
    return none;
  }

  ctx.budget.take();
  let tree: ExpandedNotebook[];
  try {
    tree = await ctx.deps.graph.getExpandedTree();
  } catch (err) {
    if (!(err instanceof GraphRequestError)) throw err;
    logEvent('sync-tree-failed', { status: err.status, reason: reasonOf(err) });
    await ctx.deps.store.patchSyncState({ lastTreeFailureAt: ctx.startedAtIso });
    return none;
  }

  ctx.tally.treeRead = true;

  const built = buildStructure(tree, ctx.selection);

  // Taken whether or not the hash moved. This is the whole fix: `structureHashOf`
  // excludes timestamps on purpose, so when the hash matches nothing writes them and the
  // stored copies stay at whatever the last structure rewrite recorded.
  const liveMtimes: SectionMtimes = new Map(
    built.sections.map((section) => [section.id, section.graphLastModifiedDateTime]),
  );
  const mirroredNotebookIds = built.notebooks.filter((n) => n.mirrored).map((n) => n.id);

  ctx.tally.unknownNotebookIds = built.unknownNotebookIds.length;
  if (built.unknownNotebookIds.length > 0) {
    logEvent('mirror-selection-unknown', { count: built.unknownNotebookIds.length });
  }

  ctx.tally.unknownActiveNotebookIds = built.unknownActiveNotebookIds.length;
  if (built.unknownActiveNotebookIds.length > 0) {
    logEvent('mirror-selection-unknown-active', { count: built.unknownActiveNotebookIds.length });
  }

  const hash = structureHashOf(built);
  if (hash === ctx.state.structureHash) {
    return { rewritten: false, liveMtimes, mirroredNotebookIds };
  }

  await ctx.deps.store.putStructure(built);
  await ctx.deps.store.patchSyncState({ structureHash: hash });
  return { rewritten: true, liveMtimes, mirroredNotebookIds };
}
```

- [ ] **Step 6: Add the overlay and use it in both passes**

Add to `src/mirror-sync.ts`, directly above `pickCandidates`:

```ts
/**
 * Stored sections carrying this run's observed timestamps.
 *
 * A section absent from `live` keeps its stored value rather than losing it. That case is
 * only reachable when the tree read failed, and `timestampsAreFresh` is false then, so
 * `pickCandidates` returns everything regardless.
 */
export function withLiveMtimes(
  sections: readonly MirrorSection[],
  live: SectionMtimes,
): MirrorSection[] {
  return sections.map((section) =>
    live.has(section.id)
      ? { ...section, graphLastModifiedDateTime: live.get(section.id) ?? null }
      : section,
  );
}
```

In `incrementalPass` (`src/mirror-sync.ts:499-510`):

```ts
  const structure = await reconcileStructure(ctx);
  const reactivated = await reconcileActivity(ctx);

  const sections = await ctx.deps.store.listSectionsToSync();
  const { eligible, skippedInactive } = splitByActivity(sections, ctx.selection, true);
  ctx.tally.sectionsSkippedInactive = skippedInactive;

  const state = reactivated ? { ...ctx.state, sectionsScannedThrough: null } : ctx.state;
  const candidates = pickCandidates(
    withLiveMtimes(eligible, structure.liveMtimes),
    state,
    ctx.tally.treeRead && !structure.rewritten,
  );
```

In `sweepPass` (`src/mirror-sync.ts:1006`, `:1024`):

```ts
  const structure = await reconcileStructure(ctx);
  await learnNestedGroups(ctx);
```

```ts
  const sections = unscoped
    ? all
    : pickCandidates(withLiveMtimes(all, structure.liveMtimes), ctx.state, ctx.tally.treeRead);
```

- [ ] **Step 7: Run the tests**

Run: `npm run typecheck && npm test`
Expected: PASS, including the new regression test.

- [ ] **Step 8: Correct the comment that made the bug invisible**

`structureHashOf`'s docstring in `src/mirror-sync.ts:308-312` says the timestamps "are read from the live tree, not from the stored copy". That was the intent and not the behaviour. Replace that sentence:

```ts
 * month for a tree that changes when someone adds a notebook. The timestamps are
 * deliberately **excluded**, because they move constantly and would rewrite every
 * document on every run. `reconcileStructure` therefore returns them separately, and
 * `withLiveMtimes` overlays them onto the stored sections before `pickCandidates` reads
 * them — without that overlay the stored copies freeze here and the sync goes blind.
```

- [ ] **Step 9: Commit**

```bash
git add src/mirror-sync.ts test/mirror-sync.test.ts
git commit -m "fix: filter sections on live timestamps, not frozen stored ones

pickCandidates compared section.graphLastModifiedDateTime read from
Firestore against a cutoff that advances every completed run. That field
is written only by putStructure, which runs only when structureHashOf
moves — and the hash excludes timestamps by design. So the stored value
froze at the last structure rewrite while the cutoff marched past it, and
about an hour later no section was ever a candidate again. Nothing
repaired it: there is no scheduled incremental in steady state, and the
nightly full sweep reconciles page ids without re-reading content.

reconcileStructure now returns the timestamps the tree read observed
whether or not it rewrote anything, and both passes overlay them before
filtering. The store fake's putStructure now replaces the collections the
way Firestore's does, which is what makes the regression test possible at
all — the counting fake could not express the failure."
```

---

### Task 2: A structure write touches only what the tree changed

**Goal:** Rewriting the structure stops resetting sync-owned state, and stops writing documents whose tree fields did not move. A notebook nobody edited must come out of a structure write with its watermark, its page count and its `childGroupsKnown` intact, and with no Firestore write against it at all.

**Why this matters most:** `structureHashOf` hashes every notebook, group and section **in the whole account**, mirrored or not. Any hash change calls `putStructure`, and `#replaceCollection` (`src/mirror-store.ts:563`) does `batch.set(...)` with **no merge**, using documents `buildStructure` produced — and `buildStructure` emits `pagesSyncedThrough: null`, `pageCount: 0` and `childGroupsKnown: false` on every one (`:265`, `:279`, `:293`). So the watermark reset is not a decision about the selection; it is what happens to every document whenever the hash moves, for any reason:

| Edit | Hash moves? | Result today |
|---|---|---|
| Add a notebook to `notebookIds` | yes | every section's watermark nulled |
| Rename one section in the OneNote client | yes — `displayName` is hashed | every section's watermark nulled |
| Create a notebook in an unmirrored part of the account | yes — the hash covers all 55 | every section's watermark nulled |
| Move a section between groups | yes — `parentId` is hashed | every section's watermark nulled |

The cost of a reset is a full re-backfill of the mirrored selection: 202 listings plus one content fetch per page across ~2000 pages, which is what needed a dedicated `POST /sync` scheduler job in the first place. `childGroupsKnown: false` separately sends every `list_sections` on a section group to Graph until the next sweep re-learns them (`src/mirror-reader.ts:340`), and re-learning them is one Graph request per group.

**The split:** every structure document has tree-owned fields, which `buildStructure` knows, and sync-owned fields, which only the sync and the write tools know. A structure write may set the first and must never touch the second. An `identity` string on each document, built from the tree-owned fields alone, is what lets the store skip a document that did not change — and hashing those same strings is what `structureHashOf` becomes, so the two notions of "changed" cannot drift.

**Files:**
- Modify: `src/mirror-schema.ts` — identity helpers, tree-field types, creation defaults
- Modify: `src/mirror-sync.ts` — `buildStructure` emits tree fields plus `identity`; `structureHashOf` hashes the identity lines
- Modify: `src/mirror-store.ts:563-604` — `#replaceCollection` merges, and skips unchanged documents
- Test: `test/mirror-schema.test.ts`, `test/mirror-sync.test.ts`

**Acceptance Criteria:**
- [ ] Renaming one section writes exactly one section document and no other
- [ ] No structure write clears `pagesSyncedThrough`, `pageCount`, `lastSweptAt`, `childGroupsKnown`, `pendingWrites` or `pendingWritesSince` on a document that already exists
- [ ] A newly appearing section is created with `pagesSyncedThrough: null` and a newly appearing group with `childGroupsKnown: false`
- [ ] Adding a notebook to `notebookIds` writes only that notebook's own documents
- [ ] A section that disappears from the tree is still deleted
- [ ] `structureHashOf` and the per-document skip agree, because both read the same identity string

**Verify:** `npm run typecheck && npm test` → 0 failures

**Steps:**

- [ ] **Step 1: Write the failing identity tests**

Add to `test/mirror-schema.test.ts`:

```ts
test('a section identity covers the tree fields and excludes the timestamp', () => {
  const base = {
    id: 'sec-1',
    displayName: 'Daily',
    notebookId: 'nb-1',
    parentId: 'nb-1',
    parentKind: 'notebook' as const,
    path: '2026 / Daily',
    mirrored: true,
    graphLastModifiedDateTime: '2026-08-19T11:00:00Z',
  };

  // The timestamp moves whenever anyone edits a page. Including it would rewrite every
  // section document on every structure change, which is the cost this whole task removes.
  assert.equal(
    sectionIdentity(base),
    sectionIdentity({ ...base, graphLastModifiedDateTime: '2026-08-19T23:59:00Z' }),
  );

  // Each of these is a real structural change and must produce a different identity.
  assert.notEqual(sectionIdentity(base), sectionIdentity({ ...base, displayName: 'Weekly' }));
  assert.notEqual(sectionIdentity(base), sectionIdentity({ ...base, mirrored: false }));
  assert.notEqual(sectionIdentity(base), sectionIdentity({ ...base, parentId: 'grp-1' }));
  assert.notEqual(sectionIdentity(base), sectionIdentity({ ...base, path: '2026 / Other' }));
});

test('a field separator cannot be forged out of a display name', () => {
  // The separator is a NUL, which no OneNote display name can contain, so two different
  // trees cannot collide on one identity string.
  const a = { id: 'a', displayName: 'x y', notebookId: 'nb', parentId: 'nb', parentKind: 'notebook' as const, path: 'p', mirrored: true, graphLastModifiedDateTime: null };
  assert.ok(!sectionIdentity(a).includes('  '));
});
```

- [ ] **Step 2: Run to see them fail**

Run: `npm test -- --test-name-pattern="section identity"`
Expected: FAIL — `sectionIdentity is not defined`

- [ ] **Step 3: Add the types, the defaults and the identity helpers**

Add to `src/mirror-schema.ts`, after the `MirrorSection` interface:

```ts
/**
 * The fields on a structure document that come out of the tree.
 *
 * Everything else — a section's watermark and page count, a group's `childGroupsKnown`, a
 * section's listing hold — is owned by the sync and by the write tools. A structure write
 * sets these and merges; it must never send the rest, because `buildStructure` does not
 * know them and would send a default that silently resets hours of work.
 */
export type NotebookTreeFields = Pick<
  MirrorNotebook,
  'id' | 'displayName' | 'mirrored' | 'sectionCount' | 'sectionGroupCount' | 'graphLastModifiedDateTime'
>;

export type SectionGroupTreeFields = Pick<
  MirrorSectionGroup,
  'id' | 'displayName' | 'notebookId' | 'parentId' | 'parentKind' | 'path' | 'mirrored'
>;

export type SectionTreeFields = Pick<
  MirrorSection,
  | 'id'
  | 'displayName'
  | 'notebookId'
  | 'parentId'
  | 'parentKind'
  | 'path'
  | 'mirrored'
  | 'graphLastModifiedDateTime'
>;

/**
 * What a structure document is created with the first time it appears, and never again.
 *
 * `pagesSyncedThrough: null` is what makes a brand new section a backfill candidate.
 * Re-sending it to a section that already exists is the reset this task removes.
 */
export const NEW_SECTION_DEFAULTS = { pagesSyncedThrough: null, pageCount: 0 } as const;
export const NEW_GROUP_DEFAULTS = { childGroupsKnown: false } as const;

/**
 * A document's tree fields as one string, so "did the tree change here?" is one compare.
 *
 * The separator is a NUL, which no OneNote display name can contain, so two different
 * trees cannot produce the same identity. `graphLastModifiedDateTime` is deliberately
 * absent from all three: it moves whenever anyone edits a page, and a document rewritten
 * because a timestamp moved is exactly what the skip in `#replaceCollection` exists to
 * avoid. The sync reads the live value off the tree anyway — see `withLiveMtimes`.
 */
export function notebookIdentity(notebook: NotebookTreeFields): string {
  return ['n', notebook.id, notebook.displayName, String(notebook.mirrored),
    String(notebook.sectionCount), String(notebook.sectionGroupCount)].join(' ');
}

export function groupIdentity(group: SectionGroupTreeFields): string {
  return ['g', group.id, group.displayName, group.parentId, group.parentKind, group.path,
    String(group.mirrored)].join(' ');
}

export function sectionIdentity(section: SectionTreeFields): string {
  return ['s', section.id, section.displayName, section.parentId, section.parentKind,
    section.path, String(section.mirrored)].join(' ');
}
```

- [ ] **Step 4: Run the identity tests**

Run: `npm test -- --test-name-pattern="identity"`
Expected: PASS

- [ ] **Step 5: Write the failing blast-radius tests**

Add to `test/mirror-sync.test.ts`:

```ts
test('renaming one section writes that section and nothing else', async () => {
  const untouched = section({ id: 'sec-2', displayName: 'Weekly', path: '2026 / Weekly' });
  const renamed = section({ id: 'sec-1', displayName: 'Daily', path: '2026 / Daily' });

  const h = harness({}, { sections: [renamed, untouched] });

  // Rename sec-1 in the tree only; the store still holds the old name and both watermarks.
  h.deps.graph.getExpandedTree = () =>
    Promise.resolve([
      {
        id: NB,
        displayName: '2026',
        sections: [
          { id: 'sec-1', displayName: 'Journal' },
          { id: 'sec-2', displayName: 'Weekly' },
        ],
        sectionGroups: [],
      },
    ]);

  await runIncremental(h.deps, BUDGET);

  assert.deepEqual(
    h.storeCalls.structureWrites.map((w) => w.id).sort(),
    ['sec-1'],
    'only the renamed section is written',
  );
  assert.deepEqual(h.storeCalls.structureResets, [], 'and no watermark is reset');
});

test('adding a notebook to the selection leaves every other notebook alone', async () => {
  const mine = section({ id: 'sec-1', notebookId: NB, parentId: NB });
  const theirs = section({ id: 'sec-9', notebookId: 'nb-2', parentId: 'nb-2', mirrored: false });

  const h = harness(
    {},
    { sections: [mine, theirs], selection: sel([NB, 'nb-2']) },
  );

  h.deps.graph.getExpandedTree = () =>
    Promise.resolve([
      { id: NB, displayName: '2026', sections: [{ id: 'sec-1', displayName: 'Daily' }], sectionGroups: [] },
      { id: 'nb-2', displayName: '2025', sections: [{ id: 'sec-9', displayName: 'Old' }], sectionGroups: [] },
    ]);

  await runIncremental(h.deps, BUDGET);

  assert.deepEqual(
    h.storeCalls.structureWrites.map((w) => w.id).sort(),
    ['nb-2', 'sec-9'],
    'only the notebook that changed and its own sections are written',
  );
});

test('a section that leaves the tree is still deleted', async () => {
  const h = harness({}, { sections: [section({ id: 'sec-1' }), section({ id: 'gone' })] });

  h.deps.graph.getExpandedTree = () =>
    Promise.resolve([
      { id: NB, displayName: '2026', sections: [{ id: 'sec-1', displayName: 'Daily' }], sectionGroups: [] },
    ]);

  await runIncremental(h.deps, BUDGET);

  assert.deepEqual(h.storeCalls.structureDeletes, ['gone']);
});
```

Change the store fake's `putStructure` (edited in Task 1 Step 2) to record what a merging, skipping store would do, and add `structureWrites: { id: string }[]`, `structureResets: string[]` and `structureDeletes: string[]` to `StoreCalls`:

```ts
    // Mirrors `#replaceCollection`: create what is new, merge tree fields onto what
    // changed, skip what did not, delete what the tree no longer holds. Recording the
    // three separately is what makes a blast-radius assertion possible.
    putStructure: (structure) => {
      calls.structures += 1;

      const applyOne = <T extends { id: string }>(
        existing: T[],
        incoming: readonly T[],
        identityOf: (item: T) => string,
        defaults: Partial<T>,
      ): T[] => {
        const byId = new Map(existing.map((item) => [item.id, item]));
        const kept: T[] = [];

        for (const item of incoming) {
          const stored = byId.get(item.id);
          if (stored === undefined) {
            calls.structureWrites.push({ id: item.id });
            calls.structureResets.push(item.id);
            kept.push({ ...defaults, ...item });
            continue;
          }
          byId.delete(item.id);
          if (identityOf(stored) === identityOf(item)) {
            kept.push(stored);
            continue;
          }
          calls.structureWrites.push({ id: item.id });
          kept.push({ ...stored, ...item });
        }

        for (const orphan of byId.keys()) calls.structureDeletes.push(orphan);
        return kept;
      };

      data.sections = applyOne(data.sections, structure.sections, sectionIdentity, {
        pagesSyncedThrough: null,
        pageCount: 0,
      } as Partial<MirrorSection>);
      data.groups = applyOne(data.groups, structure.sectionGroups, groupIdentity, {
        childGroupsKnown: false,
      } as Partial<MirrorSectionGroup>);

      return Promise.resolve();
    },
```

- [ ] **Step 6: Run to see them fail**

Run: `npm test -- --test-name-pattern="writes that section and nothing else"`
Expected: FAIL — `buildStructure` still emits `pagesSyncedThrough` on every section, so `structureResets` names both

- [ ] **Step 7: Stop `buildStructure` emitting sync-owned fields**

In `src/mirror-sync.ts`, change `BuiltStructure` to carry tree fields plus the identity, and drop the defaults from the three `push` calls:

```ts
export interface BuiltStructure {
  readonly notebooks: (NotebookTreeFields & { readonly identity: string })[];
  readonly sectionGroups: (SectionGroupTreeFields & { readonly identity: string })[];
  readonly sections: (SectionTreeFields & { readonly identity: string })[];
  readonly unknownNotebookIds: string[];
  readonly unknownActiveNotebookIds: string[];
}
```

In `buildStructure`, replace each `push`. The notebook push (`:246-253`):

```ts
    const notebook_ = {
      id: notebook.id,
      displayName: notebook.displayName,
      mirrored,
      sectionCount: notebook.sections.length,
      sectionGroupCount: notebook.sectionGroups.length,
      graphLastModifiedDateTime: notebook.lastModifiedDateTime ?? null,
    };
    notebooks.push({ ...notebook_, identity: notebookIdentity(notebook_) });
```

The direct-section push (`:255-267`) — note `pagesSyncedThrough` and `pageCount` are **gone**:

```ts
    for (const section of notebook.sections) {
      const fields = {
        id: section.id,
        displayName: section.displayName,
        notebookId: notebook.id,
        parentId: notebook.id,
        parentKind: 'notebook' as const,
        path: `${notebook.displayName} / ${section.displayName}`,
        mirrored,
        graphLastModifiedDateTime: section.lastModifiedDateTime ?? null,
      };
      sections.push({ ...fields, identity: sectionIdentity(fields) });
    }
```

The group push (`:269-280`) — `childGroupsKnown` is **gone**:

```ts
      const fields = {
        id: group.id,
        displayName: group.displayName,
        notebookId: notebook.id,
        parentId: notebook.id,
        parentKind: 'notebook' as const,
        mirrored,
        path: groupPath,
      };
      sectionGroups.push({ ...fields, identity: groupIdentity(fields) });
```

And the group's sections (`:282-295`), the same shape as the direct sections but with `parentId: group.id`, `parentKind: 'sectionGroup' as const` and `path: \`${groupPath} / ${section.displayName}\``.

Replace `structureHashOf` so it hashes the identity strings the documents carry:

```ts
/**
 * A hash of everything about the tree the mirror stores.
 *
 * An unchanged hash skips the structure pass entirely. It is built from the same identity
 * strings the documents carry, so "the hash moved" and "this document changed" can never
 * disagree — `#replaceCollection` uses the per-document form to write only what moved,
 * and this is the cheap check that decides whether to call it at all.
 *
 * The timestamps are absent for the reason `sectionIdentity` gives.
 */
export function structureHashOf(built: BuiltStructure): string {
  const hash = createHash('sha256');
  for (const notebook of built.notebooks) hash.update(`${notebook.identity}\n`);
  for (const group of built.sectionGroups) hash.update(`${group.identity}\n`);
  for (const section of built.sections) hash.update(`${section.identity}\n`);
  return hash.digest('hex');
}
```

Widen `SyncStore.putStructure`'s parameter type to `BuiltStructure`'s three arrays.

- [ ] **Step 8: Make the real store merge and skip**

Replace `#replaceCollection` in `src/mirror-store.ts:563-604`:

```ts
  /**
   * Bring a structure collection in line with `documents`, writing only what moved.
   *
   * Three behaviours, and each one is a bug that existed before it:
   *
   * - **Merge, not replace.** `batch.set` with no merge sent whatever `buildStructure`
   *   emitted, and that used to include `pagesSyncedThrough: null` — so any tree change
   *   anywhere in the account reset every section's watermark and triggered a full
   *   re-backfill. Sync-owned fields are no longer in the incoming documents at all, and
   *   merging is what keeps the ones already stored.
   * - **Defaults only on creation.** A document that does not exist yet needs
   *   `pagesSyncedThrough: null` or `childGroupsKnown: false` to start life correctly. One
   *   that does exist must not be given them again.
   * - **Skip an unchanged document.** The identity string decides. Renaming one section
   *   used to rewrite all 570 documents in the account; now it writes one.
   *
   * Deletion is still by absence from the incoming set, which the same read supplies.
   */
  async #replaceCollection(
    collection: CollectionReference,
    documents: readonly { id: string; identity: string }[],
    createDefaults: Record<string, unknown>,
  ): Promise<void> {
    await this.#run(`replacing ${collection.id}`, async () => {
      const snapshot = await collection.select('identity').get();
      const existing = new Map(
        snapshot.docs.map((doc) => [doc.id, doc.get('identity') as unknown]),
      );

      let batch = this.#firestoreOf(this.#root).batch();
      let queued = 0;
      const flush = async (): Promise<void> => {
        if (queued === 0) return;
        await batch.commit();
        batch = this.#firestoreOf(this.#root).batch();
        queued = 0;
      };

      for (const document of documents) {
        const encoded = encodeMirrorId(document.id);
        const had = existing.has(encoded);
        const storedIdentity = existing.get(encoded);
        existing.delete(encoded);

        // An unchanged document is left exactly as it is — no write, and `lastSeenAt`
        // stays where it was. Absence from the tree is what deletes, not staleness here.
        if (had && storedIdentity === document.identity) continue;

        batch.set(
          collection.doc(encoded),
          {
            ...(had ? {} : createDefaults),
            ...document,
            lastSeenAt: FieldValue.serverTimestamp(),
          },
          { merge: true },
        );
        queued += 1;
        if (queued >= BATCH_LIMIT) await flush();
      }

      for (const staleId of existing.keys()) {
        batch.delete(collection.doc(staleId));
        queued += 1;
        if (queued >= BATCH_LIMIT) await flush();
      }

      await flush();
    });
  }
```

And its three call sites in `putStructure`:

```ts
    await this.#replaceCollection(this.#notebooks(), structure.notebooks, {});
    await this.#replaceCollection(this.#sectionGroups(), structure.sectionGroups, NEW_GROUP_DEFAULTS);
    await this.#replaceCollection(this.#sections(), structure.sections, NEW_SECTION_DEFAULTS);
```

Import `NEW_GROUP_DEFAULTS` and `NEW_SECTION_DEFAULTS` from `./mirror-schema.ts`.

- [ ] **Step 9: Run the tests**

Run: `npm run typecheck && npm test`
Expected: PASS. Existing tests that assert on `calls.structures` still hold — the counter is per `putStructure` call, not per document.

- [ ] **Step 10: Record the one-off cost of the first deploy in the README**

Existing documents carry no `identity` field, so the first `putStructure` after this deploy writes every document once. Add to `README.md` under **What keeps it current**:

```markdown
**The first structure write after upgrading rewrites every document once.** Structure
documents carry an `identity` string that says what the tree said about them, and a
document written before that field existed has none — so it does not match and is written.
It is one pass of Firestore writes and no Graph requests, and it happens once.
```

- [ ] **Step 11: Commit**

```bash
git add src/mirror-schema.ts src/mirror-sync.ts src/mirror-store.ts test/ README.md
git commit -m "fix: a structure write touches only what the tree changed

structureHashOf covers every notebook and section in the account, and any
change to it called putStructure, which set every document with no merge
using what buildStructure emitted — including pagesSyncedThrough: null,
pageCount: 0 and childGroupsKnown: false. So renaming one section in the
OneNote client, or creating a notebook in an unmirrored corner of the
account, reset every section's watermark and triggered a full re-backfill
of the whole mirror: 202 listings and a content fetch per page across
~2000 pages. childGroupsKnown separately sent every list_sections on a
group to Graph until the next sweep re-learned them.

buildStructure now emits only the fields the tree supplies, plus an
identity string over them. #replaceCollection merges rather than replaces,
applies creation defaults only to a document that does not exist yet, and
skips one whose identity is unchanged. structureHashOf hashes the same
identity strings, so the collection-level check and the per-document check
cannot disagree. Renaming a section now writes one document."
```

---

### Task 3: Widening is per notebook, not global

**Goal:** Adding a notebook to `notebookIds` or to `activeNotebookIds` makes only that notebook's sections eligible for a wide scan. No other notebook is listed, fetched or re-checked because of it.

**Why:** `reconcileActivity` (`src/mirror-sync.ts:595`) patches `sectionsScannedThrough: null` when the active set changes. Tier 1 of `pickCandidates` then compares against `overlapFrom(null)`, which is the epoch, so **every** mirrored active section becomes a candidate and costs one `listPagesChangedSince` — about 70 requests on this account for a change that concerns one notebook.

The widening is necessary and the aim is wrong. A notebook frozen for three months has sections whose `graphLastModifiedDateTime` is three months old, and the tier-1 cutoff advances on every completed run, so without widening they are older than the cutoff for ever. But `sectionsScannedThrough` is a single global value and a global lever cannot be aimed. A per-notebook set can.

The same gap applies to `notebookIds`, and Task 2 exposes it: with watermarks now surviving, a notebook removed from the selection and later re-added keeps its old watermarks — correct, and it lists from where it left off — but its sections' timestamps may be older than the cutoff, so they would never become candidates. Both list changes therefore feed one mechanism.

**What each section still needs is only to be a *candidate*.** Its own `pagesSyncedThrough` is untouched by any of this, so `listPagesChangedSince(section, its own watermark − overlap)` returns everything edited while the notebook was frozen. No watermark is reset and no page is re-fetched that has not changed.

**Files:**
- Modify: `src/mirror-schema.ts` — sync-state fields, `notebooksNeedingWideScan`; remove `activeSelectionHash`
- Modify: `src/mirror-sync.ts` — `reconcileActivity` becomes `reconcileSelection`; `pickCandidates` takes a wide-scan set; `incrementalPass` clears it; remove `activeSelectionHashOf`
- Test: `test/mirror-schema.test.ts`, `test/mirror-sync.test.ts`

**Acceptance Criteria:**
- [ ] Adding one notebook to `activeNotebookIds` lists only that notebook's sections
- [ ] Adding one notebook to `notebookIds` lists only that notebook's sections
- [ ] Removing a notebook from either list widens nothing
- [ ] `sectionsScannedThrough` is never nulled by a selection change
- [ ] A run that did not complete keeps the wide-scan set, so the widening survives a budget stop
- [ ] The first run after this deploy records the lists and widens nothing
- [ ] `activeNotebookIds` going from a list to `null` widens every mirrored notebook that was not already active, and nothing else

**Verify:** `npm run typecheck && npm test` → 0 failures

**Steps:**

- [ ] **Step 1: Write the failing diff tests**

Add to `test/mirror-schema.test.ts`:

```ts
test('notebooksNeedingWideScan names only what became mirrored or active', () => {
  const mirroredIds = ['a', 'b', 'c'];

  // Adding one notebook to the selection widens that one.
  assert.deepEqual(
    notebooksNeedingWideScan(
      { mirrored: ['a', 'b'], active: ['a', 'b'] },
      { notebookIds: ['a', 'b', 'c'], activeNotebookIds: ['a', 'b'] },
      mirroredIds,
    ),
    ['c'],
  );

  // Activating one widens that one.
  assert.deepEqual(
    notebooksNeedingWideScan(
      { mirrored: ['a', 'b'], active: ['a'] },
      { notebookIds: ['a', 'b'], activeNotebookIds: ['a', 'b'] },
      ['a', 'b'],
    ),
    ['b'],
  );

  // Removing widens nothing: nothing has to be caught up on.
  assert.deepEqual(
    notebooksNeedingWideScan(
      { mirrored: ['a', 'b'], active: ['a', 'b'] },
      { notebookIds: ['a'], activeNotebookIds: ['a'] },
      ['a'],
    ),
    [],
  );

  // `null` means every mirrored notebook is active, so a list becoming null activates
  // everything that was not already in the list.
  assert.deepEqual(
    notebooksNeedingWideScan(
      { mirrored: ['a', 'b', 'c'], active: ['a'] },
      { notebookIds: ['a', 'b', 'c'], activeNotebookIds: null },
      mirroredIds,
    ),
    ['b', 'c'],
  );

  // ...and null becoming a subset activates nothing new.
  assert.deepEqual(
    notebooksNeedingWideScan(
      { mirrored: ['a', 'b'], active: null },
      { notebookIds: ['a', 'b'], activeNotebookIds: ['a'] },
      ['a', 'b'],
    ),
    [],
  );

  // Never recorded — a state document written before these fields existed. Nothing was
  // skipped under it, so there is nothing to widen; recording is the whole job.
  assert.deepEqual(
    notebooksNeedingWideScan(
      { mirrored: null, active: null },
      { notebookIds: ['a', 'b'], activeNotebookIds: null },
      ['a', 'b'],
    ),
    [],
  );
});
```

- [ ] **Step 2: Run to see it fail**

Run: `npm test -- --test-name-pattern="notebooksNeedingWideScan"`
Expected: FAIL — `notebooksNeedingWideScan is not defined`

- [ ] **Step 3: Add the sync-state fields and the diff**

In `src/mirror-schema.ts`, replace `activeSelectionHash` in `MirrorSyncState` with:

```ts
  /**
   * The two selection lists this service last reconciled, so a change can be diffed.
   *
   * A hash could say *that* they changed and not *which* notebook, and which notebook is
   * the whole point: widening the scan for every notebook when one was activated costs a
   * listing request per section of the account. `activeNotebookIdsSeen` is `null` for
   * "every mirrored notebook is active", exactly as in the selection document — which is
   * why `selectionSeen` exists to tell that apart from "never recorded".
   */
  readonly selectionSeen: boolean;
  readonly mirroredNotebookIdsSeen: readonly string[] | null;
  readonly activeNotebookIdsSeen: readonly string[] | null;
  /**
   * Notebooks whose sections bypass the tier-1 timestamp cutoff until a run completes.
   *
   * Stored rather than held for one run, because a run is budget-bounded and may stop
   * with sections outstanding. It is cleared in the same place `sectionsScannedThrough`
   * advances — only when every candidate was visited.
   */
  readonly wideScanNotebookIds: readonly string[];
```

Add to `initialSyncState`:

```ts
    selectionSeen: false,
    mirroredNotebookIdsSeen: null,
    activeNotebookIdsSeen: null,
    wideScanNotebookIds: [],
```

Add to `readSyncState`, replacing the `activeSelectionHash` line:

```ts
    selectionSeen: booleanOr(data['selectionSeen'], false),
    mirroredNotebookIdsSeen: readIdList(data['mirroredNotebookIdsSeen']),
    activeNotebookIdsSeen: readIdList(data['activeNotebookIdsSeen']),
    wideScanNotebookIds: readIdList(data['wideScanNotebookIds']) ?? [],
```

Add the diff, after `isActive`:

```ts
/** The two lists as they were last recorded. `null` on either means "never recorded". */
export interface SeenSelection {
  readonly mirrored: readonly string[] | null;
  readonly active: readonly string[] | null;
}

/**
 * Which notebooks this run must scan without the timestamp cutoff, and no others.
 *
 * A notebook that has just been mirrored or just been activated has sections whose
 * `graphLastModifiedDateTime` may be months old — older than a cutoff that advances on
 * every completed run — so without this they would never be candidates again. A notebook
 * that was *removed* from either list needs nothing: there is no catching up to do.
 *
 * Only candidacy is widened. Each section still lists against its own untouched
 * `pagesSyncedThrough`, so a notebook coming back after three months costs one listing
 * per section and one fetch per page that really changed.
 *
 * `mirrored: null` means the lists were never recorded — a state document written before
 * these fields existed. Nothing was skipped for inactivity under it, so nothing is
 * widened; recording is all the first run does.
 */
export function notebooksNeedingWideScan(
  previous: SeenSelection,
  current: NotebookSelection,
  mirroredNotebookIds: readonly string[],
): string[] {
  if (previous.mirrored === null) return [];

  const activeSet = (list: readonly string[] | null): Set<string> =>
    new Set(list === null ? mirroredNotebookIds : list);

  const wasMirrored = new Set(previous.mirrored);
  const wasActive = activeSet(previous.active);
  const isActiveNow = activeSet(current.activeNotebookIds);

  const widened = new Set<string>();
  for (const id of current.notebookIds) if (!wasMirrored.has(id)) widened.add(id);
  for (const id of isActiveNow) if (!wasActive.has(id)) widened.add(id);

  return [...widened].sort();
}
```

- [ ] **Step 4: Run the diff test**

Run: `npm test -- --test-name-pattern="notebooksNeedingWideScan"`
Expected: PASS

- [ ] **Step 5: Write the failing sync tests**

Add to `test/mirror-sync.test.ts`:

```ts
test('activating one notebook lists only that notebook’s sections', async () => {
  const mine = section({ id: 'sec-1', notebookId: NB, parentId: NB });
  const theirs = section({ id: 'sec-9', notebookId: 'nb-2', parentId: 'nb-2' });

  const h = harness(
    {},
    {
      sections: [mine, theirs],
      selection: sel([NB, 'nb-2'], [NB, 'nb-2']),
      state: {
        ...initialSyncState(),
        selectionSeen: true,
        mirroredNotebookIdsSeen: [NB, 'nb-2'],
        activeNotebookIdsSeen: [NB],
        // Both sections' timestamps are far older than this, so only the wide scan can
        // make either one a candidate.
        sectionsScannedThrough: '2026-08-19T11:59:00Z',
      },
    },
  );

  h.deps.graph.getExpandedTree = () =>
    Promise.resolve([
      { id: NB, displayName: '2026', sections: [{ id: 'sec-1', displayName: 'Daily', lastModifiedDateTime: '2020-01-01T00:00:00Z' }], sectionGroups: [] },
      { id: 'nb-2', displayName: '2025', sections: [{ id: 'sec-9', displayName: 'Old', lastModifiedDateTime: '2020-01-01T00:00:00Z' }], sectionGroups: [] },
    ]);

  await runIncremental(h.deps, BUDGET);

  assert.deepEqual(h.graphCalls.changedSince, ['sec-9'], 'only the activated notebook');
  assert.deepEqual(
    h.storeCalls.patches.filter((p) => 'sectionsScannedThrough' in p && p.sectionsScannedThrough === null),
    [],
    'the global cutoff is never nulled',
  );
});

test('removing a notebook from the active list widens nothing', async () => {
  const h = harness(
    {},
    {
      sections: [section({ id: 'sec-1', graphLastModifiedDateTime: '2020-01-01T00:00:00Z' })],
      selection: sel([NB], []),
      state: {
        ...initialSyncState(),
        selectionSeen: true,
        mirroredNotebookIdsSeen: [NB],
        activeNotebookIdsSeen: [NB],
        sectionsScannedThrough: '2026-08-19T11:59:00Z',
      },
    },
  );

  await runIncremental(h.deps, BUDGET);

  assert.deepEqual(h.graphCalls.changedSince, []);
});

test('a budget-stopped run keeps the notebooks it still has to scan wide', async () => {
  const h = harness(
    {},
    {
      sections: [
        section({ id: 'sec-1', graphLastModifiedDateTime: '2020-01-01T00:00:00Z' }),
        section({ id: 'sec-2', graphLastModifiedDateTime: '2020-01-01T00:00:00Z' }),
      ],
      selection: sel([NB], [NB]),
      state: {
        ...initialSyncState(),
        selectionSeen: true,
        mirroredNotebookIdsSeen: [NB],
        activeNotebookIdsSeen: [],
        sectionsScannedThrough: '2026-08-19T11:59:00Z',
      },
    },
  );

  // Two sections, budget enough for the tree and one listing.
  const report = await runIncremental(h.deps, { requestBudget: 2 });

  assert.equal(report.done, false);
  const cleared = h.storeCalls.patches.filter((p) => p.wideScanNotebookIds?.length === 0);
  assert.deepEqual(cleared, [], 'an unfinished run must not clear the wide-scan set');
});
```

Give `fakeStore` a `selection` it can be seeded with — it already takes `StoreState.selection` — and make sure `sel(ids, active)` is used with both arguments in these tests.

- [ ] **Step 6: Run to see them fail**

Run: `npm test -- --test-name-pattern="lists only that notebook"`
Expected: FAIL — `changedSince` holds both sections, because the global cutoff was nulled

- [ ] **Step 7: Replace `reconcileActivity` with `reconcileSelection`**

Delete `activeSelectionHashOf` and `reconcileActivity` from `src/mirror-sync.ts` and add:

```ts
/**
 * Notice a change to either selection list, and widen this run for the notebooks it named.
 *
 * Returns the set of notebook ids whose sections bypass the tier-1 cutoff. It is the
 * stored set unioned with whatever this run's diff added, because a previous run may have
 * been stopped by its budget before it worked through them.
 *
 * `sectionsScannedThrough` is deliberately **not** touched. Nulling it was the old
 * mechanism and it widened the scan for every notebook in the account — around 70 listing
 * requests here for a change that concerns one.
 */
async function reconcileSelection(
  ctx: PassContext,
  mirroredNotebookIds: readonly string[],
): Promise<ReadonlySet<string>> {
  const carried = ctx.state.wideScanNotebookIds;
  const seenPatch = {
    selectionSeen: true,
    mirroredNotebookIdsSeen: [...ctx.selection.notebookIds],
    activeNotebookIdsSeen:
      ctx.selection.activeNotebookIds === null ? null : [...ctx.selection.activeNotebookIds],
  };

  if (!ctx.state.selectionSeen) {
    // First run against a state document written before these fields existed. Record the
    // lists and widen nothing — the same rule the old null-hash branch followed, and for
    // the same reason: nothing was skipped for inactivity under it.
    await ctx.deps.store.patchSyncState(seenPatch);
    return new Set(carried);
  }

  const added = notebooksNeedingWideScan(
    { mirrored: ctx.state.mirroredNotebookIdsSeen, active: ctx.state.activeNotebookIdsSeen },
    ctx.selection,
    mirroredNotebookIds,
  );

  if (added.length === 0) return new Set(carried);

  const widened = [...new Set([...carried, ...added])].sort();
  await ctx.deps.store.patchSyncState({ ...seenPatch, wideScanNotebookIds: widened });
  logEvent('mirror-selection-widened', { count: added.length });
  return new Set(widened);
}
```

- [ ] **Step 8: Let `pickCandidates` take the set, and clear it on a completed run**

Task 1 landed `pickCandidates` as `(sections, live, state, mayFilterByTimestamp)` — it applies the `withLiveMtimes` overlay itself, so a call site cannot forget it. Adding a fifth positional parameter puts a `ReadonlySet<string>` next to a `readonly MirrorSection[]`, which is where positional order stops being self-evident; the Task 1 implementer flagged this commit as the place to reconsider. **Decide between the two shapes below and say which you chose and why.** Either is acceptable; do not leave five bare positionals without justifying it.

```ts
export function pickCandidates(
  sections: readonly MirrorSection[],
  live: SectionMtimes,
  state: MirrorSyncState,
  mayFilterByTimestamp: boolean,
  wideScan: ReadonlySet<string> = new Set(),
): MirrorSection[] {
  const observed = withLiveMtimes(sections, live);
  if (!state.sectionRollUpTrusted || !mayFilterByTimestamp) return observed;

  const since = overlapFrom(state.sectionsScannedThrough);
  return observed.filter(
    (section) =>
      // A notebook just mirrored or just activated: its sections' timestamps may predate
      // the cutoff by months, so nothing else would ever make them candidates.
      wideScan.has(section.notebookId) ||
      section.pagesSyncedThrough === null ||
      section.graphLastModifiedDateTime === null ||
      section.graphLastModifiedDateTime >= since,
  );
}
```

In `incrementalPass`:

```ts
  const structure = await reconcileStructure(ctx);
  const wideScan = await reconcileSelection(ctx, structure.mirroredNotebookIds);

  const sections = await ctx.deps.store.listSectionsToSync();
  const { eligible, skippedInactive } = splitByActivity(sections, ctx.selection, true);
  ctx.tally.sectionsSkippedInactive = skippedInactive;

  const candidates = pickCandidates(
    eligible,
    structure.liveMtimes,
    ctx.state,
    ctx.tally.treeRead && !structure.rewritten,
    wideScan,
  );
```

and in the completion block at the end of `incrementalPass`:

```ts
  // Only when every candidate completed, for the reason `sectionsScannedThrough` gives —
  // and the wide-scan set clears on exactly the same condition, because a run stopped by
  // its budget has not finished visiting the notebooks it was widened for.
  if (ctx.tally.done && ctx.tally.treeRead) {
    await ctx.deps.store.patchSyncState({
      sectionsScannedThrough: ctx.startedAtIso,
      backfillComplete: sections.every((section) => section.pagesSyncedThrough !== null),
      wideScanNotebookIds: [],
    });
  }
```

In `sweepPass`, pass the stored set without clearing it:

```ts
  const sections = unscoped
    ? all
    : pickCandidates(
        all,
        structure.liveMtimes,
        ctx.state,
        ctx.tally.treeRead,
        new Set(ctx.state.wideScanNotebookIds),
      );
```

- [ ] **Step 9: Run the tests**

Run: `npm run typecheck && npm test`
Expected: PASS. The old activity tests that asserted `sectionsScannedThrough: null` need rewriting to assert on the wide-scan set instead; that is the behaviour change this task is for.

- [ ] **Step 10: Update the CLAUDE.md paragraph this replaces**

CLAUDE.md documents "Changing the active set clears `sectionsScannedThrough`, and a null stored hash does not." Replace that paragraph:

```markdown
**Changing either selection list widens the scan for the notebooks it named, and for no
others.** `notebooksNeedingWideScan` diffs the stored `mirroredNotebookIdsSeen` and
`activeNotebookIdsSeen` against the selection document and returns only the ids that
became mirrored or became active; `pickCandidates` bypasses the tier-1 cutoff for those
notebooks' sections and nothing else. The old mechanism nulled `sectionsScannedThrough`,
which is global — one activation cost a listing request for every section in the account.
Only *candidacy* is widened: each section still lists against its own untouched
`pagesSyncedThrough`, so a notebook coming back after three months costs one listing per
section and one fetch per page that really changed. The set is stored rather than held for
one run, because a budget-bounded run may stop with sections outstanding, and it clears on
exactly the condition that advances `sectionsScannedThrough`. `selectionSeen` false is a
state document written before these fields existed: record the lists, widen nothing.
```

- [ ] **Step 11: Commit**

```bash
git add src/mirror-schema.ts src/mirror-sync.ts test/ CLAUDE.md
git commit -m "fix: widen the scan per notebook, not for the whole account

reconcileActivity nulled sectionsScannedThrough on any change to the
active set, which made every mirrored active section a candidate — around
70 listing requests here for a change concerning one notebook. The
widening is necessary: a notebook frozen for months has section timestamps
older than a cutoff that advances every run, so nothing else would ever
make its sections candidates again. But the cutoff is global and a global
lever cannot be aimed.

The state now stores the two lists rather than a hash of one, so a change
can be diffed to the notebooks it named, and pickCandidates bypasses the
cutoff for those alone. The same mechanism covers notebookIds, which
Task 2 exposed: with watermarks surviving a structure write, a re-added
notebook keeps its own and needs only to become a candidate again.

Only candidacy is widened. Each section still lists against its own
pagesSyncedThrough, so nothing is re-fetched that has not changed. The set
survives a budget-stopped run and clears on the same condition that
advances sectionsScannedThrough."
```

---

### Task 4: The sweep carries titles and notices content drift

**Goal:** The nightly sweep stores real page titles instead of `''`, and marks a page stale when Graph's timestamp has moved past the mirror's — at no extra Graph request.

**Why before the skip:** Task 7 makes a missed edit permanent rather than merely delayed. This gives the sweep the ability to notice one, so a wrong margin costs a night of staleness instead of silent unbounded loss. It also fixes a live defect: `sweepSection` synthesizes a summary with `title: ''` and `lastModifiedDateTime: '1970-01-01T00:00:00.000Z'` (`src/mirror-sync.ts:1107-1109`) because `listPageIds` asks for `$select=id` alone. Both fields reach the model — `MirrorPage.lastModifiedDateTime` is documented as "exactly as Graph spelled it; every tool result already prints this string", and `titleLower` feeds by-name matching. It does not self-heal for the case the sweep exists for: a page moved into a section may not have its own timestamp bumped by the move, so no later incremental lists it.

**Files:**
- Modify: `src/graph-structure.ts:559-565` — `listPageIds` → `listPageSummaries`
- Modify: `src/mirror-schema.ts` — add `PageStamp`, `MirrorPageDigest`, `pageHasDrifted`
- Modify: `src/mirror-store.ts:367-373` — `listPageIdsInSection` → `listPageDigestsInSection`
- Modify: `src/mirror-sync.ts` — `SyncGraph`, `SyncStore`, `sweepSection` (`:1073-1113`)
- Test: `test/mirror-schema.test.ts`, `test/mirror-sync.test.ts`, `test/graph-structure.test.ts`

**Acceptance Criteria:**
- [ ] `listPageSummaries` requests `$select=id,title,lastModifiedDateTime`
- [ ] A page the sweep discovers is stored with its real title and Graph's timestamp
- [ ] A page present in both Graph and the mirror whose timestamps differ is marked stale and costs no Graph request
- [ ] A page whose stored `contentState` is not `present` is not marked again
- [ ] A failed enumeration still deletes nothing and marks nothing

**Verify:** `npm run typecheck && npm test` → 0 failures

**Steps:**

- [ ] **Step 1: Write the failing predicate test**

Add to `test/mirror-schema.test.ts`:

```ts
test('pageHasDrifted compares Graph against Graph, and only for a present copy', () => {
  const stored = {
    id: 'p1',
    title: 'Page',
    lastModifiedDateTime: '2026-08-19T12:00:00Z',
    contentState: 'present' as const,
  };
  const live = { id: 'p1', title: 'Page', lastModifiedDateTime: '2026-08-19T12:00:00Z' };

  assert.equal(pageHasDrifted(stored, live), false, 'identical stamps are not drift');
  assert.equal(
    pageHasDrifted(stored, { ...live, lastModifiedDateTime: '2026-08-19T12:00:01Z' }),
    true,
    'one second later is drift — both sides are Graph’s own string',
  );
  assert.equal(
    pageHasDrifted({ ...stored, contentState: 'stale' }, { ...live, lastModifiedDateTime: 'x' }),
    false,
    'a copy already stale has nothing to invalidate',
  );
  assert.equal(
    pageHasDrifted(stored, { ...live, lastModifiedDateTime: '' }),
    false,
    'an absent timestamp is not evidence of a change',
  );
});
```

- [ ] **Step 2: Run to see it fail**

Run: `npm test -- --test-name-pattern="pageHasDrifted"`
Expected: FAIL — `pageHasDrifted is not defined`

- [ ] **Step 3: Add the types and the predicate**

Add to `src/mirror-schema.ts`, after `timestampToIso`:

```ts
/**
 * The three fields every page listing carries.
 *
 * Structural rather than imported, so this module keeps its property of having no imports
 * at all. `PageSummary` from ./graph-structure.ts satisfies it.
 */
export interface PageStamp {
  readonly id: string;
  readonly title: string;
  readonly lastModifiedDateTime: string;
}

/** What a sweep needs off a stored page document, and nothing more. */
export interface MirrorPageDigest {
  readonly id: string;
  readonly title: string;
  readonly lastModifiedDateTime: string;
  readonly contentState: ContentState;
}

/**
 * Has Graph's copy of this page moved past the mirror's?
 *
 * Both sides are Graph's own string, stored verbatim and re-read from the same
 * collection, so this is a same-clock comparison and equality is the right test. A
 * tolerance here would swallow a real edit, and the sweep is the last thing that would
 * ever notice it.
 *
 * A copy that is not `present` is already a miss for every read, so there is nothing to
 * invalidate. An empty live stamp is the decoder's fallback for an absent field, which is
 * not evidence of anything.
 */
export function pageHasDrifted(stored: MirrorPageDigest, live: PageStamp): boolean {
  return (
    stored.contentState === 'present' &&
    live.lastModifiedDateTime !== '' &&
    stored.lastModifiedDateTime !== live.lastModifiedDateTime
  );
}
```

- [ ] **Step 4: Run the predicate test**

Run: `npm test -- --test-name-pattern="pageHasDrifted"`
Expected: PASS

- [ ] **Step 5: Widen the Graph call**

Replace `listPageIds` in `src/graph-structure.ts:559-565`:

```ts
  /**
   * Every page in one section, with the two fields a reconciliation needs.
   *
   * This is the mirror's deletion sweep, and it is as cheap as deletion detection gets
   * on this account. Graph has no /delta on any OneNote resource and no tombstone for a
   * deleted page, and the account-wide page list — the one call that would enumerate
   * everything in one request per 100 pages — is the banned one, error 20266. So the
   * floor is one request per section plus one per additional 100 pages, and what makes
   * that affordable is that only the mirrored notebooks are swept.
   *
   * `title` and `lastModifiedDateTime` cost nothing beyond bytes on a request that is
   * already being made, and the sweep needs both: a page it discovers has no other source
   * for its title, and a page it already holds can only be checked for content drift
   * against Graph's own timestamp.
   *
   * No `top` argument: a sweep that stopped early would report pages as deleted that are
   * merely past the cutoff, which is the one mistake here that destroys data.
   */
  async listPageSummaries(sectionId: string): Promise<PageSummary[]> {
    const url =
      `${GRAPH_ROOT}/me/onenote/sections/${encodeURIComponent(sectionId)}/pages` +
      `?$select=id,title,lastModifiedDateTime&$top=${MAX_GRAPH_TOP}`;

    return (await this.#collect(url)).map((item) => toPageSummary(item, url));
  }
```

Update the exact-URL route in `test/graph-structure.test.ts` — search for `?$select=id&$top=` and change that route's key and its assertions to the new URL and the summary shape.

- [ ] **Step 6: Widen the Firestore projection**

Replace `listPageIdsInSection` in `src/mirror-store.ts:367-373`:

```ts
  /** Every stored page in one section, projected to what a sweep reconciles on. */
  async listPageDigestsInSection(sectionId: string): Promise<MirrorPageDigest[]> {
    return this.#query<MirrorPageDigest>(
      'listing mirrored page digests',
      this.#pages()
        .where('sectionId', '==', sectionId)
        .select('id', 'title', 'lastModifiedDateTime', 'contentState'),
    );
  }
```

Add `MirrorPageDigest` to the `import type` list from `./mirror-schema.ts`.

- [ ] **Step 7: Write the failing sweep tests**

Add to `test/mirror-sync.test.ts`:

```ts
test('the sweep stores a discovered page with its real title and timestamp', async () => {
  const h = harness(
    { summaries: { 'sec-1': [summary('new-page', '2026-08-19T11:45:00Z')] } },
    { sections: [section()] },
  );

  await runFullSweep(h.deps, BUDGET);

  const written = h.storeCalls.puts.find((p) => p.page.id === 'new-page');
  assert.ok(written, 'the page Graph has and the mirror lacks must be fetched');
  assert.equal(written.page.title, 'Page new-page');
  assert.equal(written.page.lastModifiedDateTime, '2026-08-19T11:45:00Z');
});

test('the sweep marks a drifted page stale and spends no Graph request on it', async () => {
  // The page is in both places, so the sweep neither deletes nor fetches it. All it can
  // do — and all it needs to do — is send the next read to Graph.
  const h = harness(
    { summaries: { 'sec-1': [summary('p1', '2026-08-19T13:00:00Z')] } },
    {
      sections: [section()],
      digestsBySection: new Map([
        [
          'sec-1',
          [{ id: 'p1', title: 'Page p1', lastModifiedDateTime: '2026-08-19T11:00:00Z', contentState: 'present' as const }],
        ],
      ]),
    },
  );

  await runFullSweep(h.deps, BUDGET);

  assert.deepEqual(h.storeCalls.staled, ['p1']);
  assert.deepEqual(h.storeCalls.puts, [], 'a drift check costs no fetch and no write');
  assert.deepEqual(h.storeCalls.deletes, []);
});

test('a sweep whose enumeration failed marks nothing stale', async () => {
  const h = harness(
    {
      summaries: {
        'sec-1': () => {
          throw graphError(500);
        },
      },
    },
    {
      sections: [section()],
      digestsBySection: new Map([
        [
          'sec-1',
          [{ id: 'p1', title: 'Page p1', lastModifiedDateTime: '2026-08-19T11:00:00Z', contentState: 'present' as const }],
        ],
      ]),
    },
  );

  await runFullSweep(h.deps, BUDGET);

  assert.deepEqual(h.storeCalls.deletes, [], 'a failed enumeration deletes nothing');
  assert.deepEqual(h.storeCalls.staled, [], 'and marks nothing stale either');
});
```

Update the fakes to match: rename `GraphScript.ids` to `summaries` holding `PageSummary[]`, rename `GraphCalls.pageIds` accordingly, rename `StoreState.pageIdsBySection` to `digestsBySection` holding `MirrorPageDigest[]`, add `staled: string[]` to `StoreCalls`, and give the store fake a `markPageStale` that pushes to it.

- [ ] **Step 8: Run to see them fail**

Run: `npm test`
Expected: FAIL — `listPageSummaries is not a function` on the sync's `SyncGraph`

- [ ] **Step 9: Rewrite `sweepSection`**

In `src/mirror-sync.ts`, change the two interface members:

```ts
export interface SyncGraph {
  getExpandedTree(): Promise<ExpandedNotebook[]>;
  listContainerChildren(kind: ContainerKind, containerId: string): Promise<ContainerChildren>;
  listPagesChangedSince(sectionId: string, sinceIso: string): Promise<PageSummary[]>;
  listPagesInSection(sectionId: string, top?: number): Promise<PageSummary[]>;
  listPageSummaries(sectionId: string): Promise<PageSummary[]>;
}
```

```ts
  listPageDigestsInSection(sectionId: string): Promise<MirrorPageDigest[]>;
  markPageStale(pageId: string): Promise<void>;
```

Replace the body of `sweepSection` (`src/mirror-sync.ts:1073-1113`):

```ts
async function sweepSection(ctx: PassContext, section: MirrorSection): Promise<void> {
  ctx.budget.take();

  let live: PageSummary[];
  try {
    live = await ctx.deps.graph.listPageSummaries(section.id);
  } catch (err) {
    logEvent('sync-section-failed', { sectionId: section.id, reason: reasonOf(err) });
    ctx.tally.pagesFailed += 1;
    return;
  }

  ctx.tally.sectionsVisited += 1;

  const liveById = new Map(live.map((page) => [page.id, page]));
  const mirrored = await ctx.deps.store.listPageDigestsInSection(section.id);

  for (const stored of mirrored) {
    const match = liveById.get(stored.id);

    if (match === undefined) {
      await deletePage(ctx, stored.id, section, 'sweep');
      continue;
    }

    // The sweep's second job, and it costs no Graph request: a page whose stored stamp is
    // behind Graph's is content the mirror would otherwise serve as current. Marking it
    // stale sends the next read to Graph, which cannot be wrong. Fetching it here instead
    // would put a content request per drifted page inside a run already sized against the
    // hourly budget.
    if (pageHasDrifted(stored, match)) {
      try {
        await ctx.deps.store.markPageStale(stored.id);
        ctx.tally.pagesStaled += 1;
      } catch (err) {
        logEvent('mirror-page-failed', { pageId: stored.id, reason: reasonOf(err) });
        ctx.tally.pagesFailed += 1;
      }
    }
  }

  // Pages Graph has that the mirror lacks — new, or moved in. Fetched now, budget
  // permitting; anything left is picked up by the next incremental, because this
  // section's watermark is deliberately not advanced here.
  const mirroredIds = new Set(mirrored.map((page) => page.id));
  for (const page of live) {
    if (mirroredIds.has(page.id)) continue;
    if (ctx.budget.exhausted) {
      ctx.tally.done = false;
      return;
    }
    await syncPage(ctx, section, page);
  }

  await ctx.deps.store.setSectionSweepResult(section.id, {
    pageCount: live.length,
    lastSweptAt: ctx.startedAtIso,
  });
}
```

Add `pagesStaled: number` to `Tally` (initialised 0 in `runMode`), to `SyncReport`, and to the `sync-completed` log line. Add `pageHasDrifted` and `MirrorPageDigest` to the `mirror-schema.ts` import list.

- [ ] **Step 10: Run the tests**

Run: `npm run typecheck && npm test`
Expected: PASS

- [ ] **Step 11: Commit**

```bash
git add src/graph-structure.ts src/mirror-schema.ts src/mirror-store.ts src/mirror-sync.ts test/
git commit -m "feat: sweep carries page titles and notices content drift

listPageIds asked for \$select=id alone, so a page the sweep discovered was
stored with title '' and lastModifiedDateTime 1970-01-01 — both of which
reach the model, and neither of which self-heals: a page moved into a
section may not have its own timestamp bumped, so no later incremental
lists it. Selecting title and lastModifiedDateTime costs no extra request.

Having the timestamp also lets the sweep compare Graph's stamp against
the stored one and markPageStale on a mismatch, which is the first thing
in this repository that can notice a page edited in the OneNote client
that the incremental pass missed. It costs no Graph request: a stale
marker sends the next read to Graph, which cannot be wrong."
```

---

### Task 5: A short-circuited write corrects a locally stamped timestamp

**Goal:** A page whose `lastModifiedDateTime` holds `resyncPage`'s local clock value gets Graph's own string back the next time the sync reads it, so Task 7's equality test can ever match.

**Why:** `resyncPage` stamps `lastModifiedDateTime` from `Date.now()` (`src/mirror-sync.ts:1198`) into a field documented as "exactly as Graph spelled it" and printed in every tool result. `writePageFromRaw`'s short-circuit deliberately does not compare that field, so when the content is unchanged it returns `false` and the local value stays for ever. `MirrorStore.putPageMetadata` (`src/mirror-store.ts:426`) already exists for exactly this write and is called from nowhere.

**Files:**
- Modify: `src/mirror-store.ts:426-436` — stamp `contentSyncedAt` too
- Modify: `src/mirror-sync.ts` — `PageWriteDeps`, `SyncStore`, `writePageFromRaw` short-circuit (`:884-895`)
- Test: `test/mirror-sync.test.ts`

**Acceptance Criteria:**
- [ ] A short-circuit whose live timestamp differs from the stored one writes metadata and returns `false`
- [ ] A short-circuit whose timestamps match writes nothing
- [ ] A live timestamp of `''` never overwrites a stored one
- [ ] `putPageMetadata` refreshes `contentSyncedAt`, because the caller has just re-read and compared the content

**Verify:** `npm run typecheck && npm test` → 0 failures

**Steps:**

- [ ] **Step 1: Write the failing test**

Add to `test/mirror-sync.test.ts`:

```ts
test('an unchanged page whose stamp is the local clock gets Graph’s string back', async () => {
  // resyncPage stamps lastModifiedDateTime from this process's clock, and the
  // short-circuit does not compare that field — so without this the local value is
  // permanent, and the Task 7 skip could never match.
  const stored = storedPage({
    id: 'p1',
    contentHash: hashOf('<p>same</p>'),
    title: 'Page p1',
    lastModifiedDateTime: '2026-08-19T11:59:58.123Z',
  });

  const h = harness(
    { changed: { 'sec-1': [summary('p1', '2026-08-19T12:00:00Z')] } },
    { sections: [section()], pages: new Map([['p1', stored]]) },
    { fetchRaw: () => Promise.resolve(rawHtml('<p>same</p>')) },
  );

  const report = await runIncremental(h.deps, BUDGET);

  assert.equal(report.pagesUpdated, 0, 'the content did not change, so nothing is rewritten');
  assert.deepEqual(h.storeCalls.puts, [], 'and no full page write happens');
  assert.deepEqual(h.storeCalls.metadata, [
    { id: 'p1', lastModifiedDateTime: '2026-08-19T12:00:00Z' },
  ]);
});

test('a short-circuit whose stamps already match writes nothing at all', async () => {
  const stored = storedPage({
    id: 'p1',
    contentHash: hashOf('<p>same</p>'),
    title: 'Page p1',
    lastModifiedDateTime: '2026-08-19T12:00:00Z',
  });

  const h = harness(
    { changed: { 'sec-1': [summary('p1', '2026-08-19T12:00:00Z')] } },
    { sections: [section()], pages: new Map([['p1', stored]]) },
    { fetchRaw: () => Promise.resolve(rawHtml('<p>same</p>')) },
  );

  await runIncremental(h.deps, BUDGET);

  assert.deepEqual(h.storeCalls.puts, []);
  assert.deepEqual(h.storeCalls.metadata, []);
});
```

Add to the fakes: `metadata: { id: string; lastModifiedDateTime: string }[]` on `StoreCalls`, a `putPageMetadata` on the store fake that records `{ id, lastModifiedDateTime }`, and helpers `storedPage(overrides)` building a `MirrorPage` and `hashOf(html)` returning `createHash('sha256').update(html).digest('hex')`.

- [ ] **Step 2: Run to see it fail**

Run: `npm test -- --test-name-pattern="local clock"`
Expected: FAIL — `putPageMetadata is not a function`

- [ ] **Step 3: Make `putPageMetadata` stamp the sync time**

Replace `src/mirror-store.ts:424-436`:

```ts
  /**
   * Update the metadata Graph's page list carries, leaving the content documents alone.
   *
   * `contentSyncedAt` is refreshed too, and that is not an oversight. The only caller is
   * `writePageFromRaw`'s short-circuit, which has just re-read the page from Graph and
   * found the stored content identical — so the copy really was confirmed at this moment,
   * and the settle guard in `storedPageIsCurrent` reads this field to decide whether the
   * copy has been in hand long enough to trust. Without the refresh a page corrected here
   * would fail that guard on every later run and be re-fetched each time.
   */
  async putPageMetadata(
    page: Pick<
      MirrorPage,
      'id' | 'title' | 'titleLower' | 'sectionId' | 'notebookId' | 'sectionPath' | 'lastModifiedDateTime'
    >,
  ): Promise<void> {
    await this.#run('writing mirrored page metadata', () =>
      this.#pages()
        .doc(encodeMirrorId(page.id))
        .set(
          {
            ...page,
            lastModified: new Date(page.lastModifiedDateTime),
            contentSyncedAt: FieldValue.serverTimestamp(),
          },
          { merge: true },
        ),
    );
  }
```

- [ ] **Step 4: Wire it into the short-circuit**

In `src/mirror-sync.ts`, widen the two interfaces:

```ts
export interface PageWriteDeps {
  readonly store: Pick<SyncStore, 'getPage' | 'putPage' | 'putPageMetadata'>;
  readonly blobs: MirrorBlobWriter;
}
```

```ts
  putPageMetadata(
    page: Pick<
      MirrorPage,
      'id' | 'title' | 'titleLower' | 'sectionId' | 'notebookId' | 'sectionPath' | 'lastModifiedDateTime'
    >,
  ): Promise<void>;
```

Replace the short-circuit body in `writePageFromRaw` (`src/mirror-sync.ts:884-895`):

```ts
  const stored = await deps.store.getPage(summary.id);
  if (
    stored !== null &&
    stored.contentState === 'present' &&
    stored.contentHash === hash &&
    stored.title === summary.title &&
    stored.sectionId === placement.sectionId
  ) {
    // The content is right. The stamp may not be: `resyncPage` writes this process's
    // clock, and nothing above compares the field, so the local value would otherwise
    // stay for ever — and `storedPageIsCurrent` can then never match it against Graph's.
    // An empty live stamp is the decoder's fallback for an absent field and must not
    // overwrite a good stored one.
    if (
      summary.lastModifiedDateTime !== '' &&
      summary.lastModifiedDateTime !== stored.lastModifiedDateTime
    ) {
      await deps.store.putPageMetadata({
        id: summary.id,
        title: summary.title,
        titleLower: summary.title.toLowerCase(),
        sectionId: placement.sectionId,
        notebookId: placement.notebookId,
        sectionPath: placement.sectionPath,
        lastModifiedDateTime: summary.lastModifiedDateTime,
      });
    }
    return false;
  }
```

Also add `putPageMetadata` to `ResyncDeps.store`'s `Pick` list so `resyncPage` still typechecks, and to `test/write-tools.test.ts`'s fake resync store if it constructs one.

- [ ] **Step 5: Run the tests**

Run: `npm run typecheck && npm test`
Expected: PASS

- [ ] **Step 6: Commit**

```bash
git add src/mirror-store.ts src/mirror-sync.ts test/
git commit -m "fix: restore Graph's timestamp on a short-circuited page write

resyncPage stamps lastModifiedDateTime from this process's clock, because
page metadata reads back unreliably right after a write. The
short-circuit in writePageFromRaw deliberately does not compare that
field, so an unchanged page kept the local value permanently — visible to
the model, which prints the string, and fatal to the timestamp skip two
commits from now, whose equality test could never match it.

putPageMetadata already existed for this write and was called from
nowhere. It now also stamps contentSyncedAt, because its only caller has
just confirmed the content against Graph."
```

---

### Task 6: Measure the clock skew and what the overlap actually catches

**Goal:** Two log events that turn `TIMESTAMP_SETTLE_MS` and the overlap width from chosen numbers into measured ones.

**Why before the shrink:** an overlap-save event is only informative while the window is still wide. Measuring at one hour tells you what a fifteen-minute window would have missed; measuring after the shrink cannot.

**Files:**
- Modify: `src/graph-throttle.ts` — add `recordClockSkew`
- Modify: `src/graph-structure.ts` — call it in `graphGet`
- Modify: `src/page-content.ts:126` and `src/page-write.ts:261` — call it
- Modify: `src/mirror-sync.ts` — `syncPage` returns whether it wrote; `syncSection` logs `sync-overlap-save`
- Test: `test/graph-throttle.test.ts`, `test/mirror-sync.test.ts`

**Acceptance Criteria:**
- [ ] `recordClockSkew` returns the signed skew in ms and logs only past a threshold
- [ ] A missing or unparseable `Date` header returns null and logs nothing
- [ ] `sync-overlap-save` fires only for a page that was actually written and whose stamp predates the section's watermark
- [ ] Neither event carries a page title, a section name, or any URL

**Verify:** `npm run typecheck && npm test` → 0 failures

**Steps:**

- [ ] **Step 1: Write the failing skew test**

Add to `test/graph-throttle.test.ts`:

```ts
test('recordClockSkew reads the Date header and logs only a large gap', () => {
  const lines: string[] = [];
  setEventSink((line) => lines.push(line));

  const at = (iso: string): Headers => new Headers({ date: new Date(iso).toUTCString() });
  const now = (iso: string) => () => Date.parse(iso);

  // HTTP dates carry whole seconds, so anything under a second is not measurable here.
  assert.equal(recordClockSkew(at('2026-08-19T12:00:00Z'), now('2026-08-19T12:00:02Z')), 2000);
  assert.deepEqual(lines, [], 'a two-second gap is inside the threshold');

  recordClockSkew(at('2026-08-19T12:00:00Z'), now('2026-08-19T12:00:30Z'));
  assert.equal(lines.length, 1);
  assert.deepEqual(JSON.parse(lines[0] as string), { event: 'graph-clock-skew', skewMs: 30000 });

  assert.equal(recordClockSkew(new Headers(), now('2026-08-19T12:00:00Z')), null);
  assert.equal(
    recordClockSkew(new Headers({ date: 'not a date' }), now('2026-08-19T12:00:00Z')),
    null,
  );
  assert.equal(lines.length, 1, 'an unreadable header is not an event');

  setEventSink(() => {});
});
```

- [ ] **Step 2: Run to see it fail**

Run: `npm test -- --test-name-pattern="recordClockSkew"`
Expected: FAIL — `recordClockSkew is not defined`

- [ ] **Step 3: Implement it**

Add to `src/graph-throttle.ts`:

```ts
/**
 * How far this process's clock may sit from Graph's before it is worth an event.
 *
 * An HTTP `Date` header carries whole seconds and the value is read after a round trip,
 * so nothing below a second or so is measurable. Five seconds is well clear of that and
 * well below every margin in the sync that depends on the two clocks being close.
 */
export const CLOCK_SKEW_LOG_THRESHOLD_MS = 5_000;

/**
 * Record how far this process's clock is from the service's, from a header already sent.
 *
 * Every cross-clock comparison in the mirror sync — the watermark overlap, and the settle
 * guard in `storedPageIsCurrent` — is sized on an assumption about this number that
 * nothing measured. It costs no request: `Date` is on every response.
 *
 * Returns the signed skew in milliseconds, positive when this process is ahead, or null
 * when the header is absent or unreadable. The value carries no URL and no body, so it is
 * safe in a log line.
 */
export function recordClockSkew(headers: Headers, now: () => number = Date.now): number | null {
  const header = headers.get('date');
  if (header === null) return null;

  const serverMs = Date.parse(header);
  if (Number.isNaN(serverMs)) return null;

  const skewMs = now() - serverMs;
  if (Math.abs(skewMs) >= CLOCK_SKEW_LOG_THRESHOLD_MS) {
    logEvent('graph-clock-skew', { skewMs: Math.round(skewMs) });
  }
  return skewMs;
}
```

Add `import { logEvent } from './logging.ts';` to `src/graph-throttle.ts` if it is not already there.

- [ ] **Step 4: Call it at the three fetch sites**

In `src/graph-structure.ts`, immediately after the `graphGet` fetch resolves:

```ts
  const response = await fetchImpl(url, { headers: { Authorization: `Bearer ${accessToken}` } });
  recordClockSkew(response.headers);
```

In `src/page-content.ts`, after `const response = await this.#fetch(...)` in `fetchRaw`:

```ts
      recordClockSkew(response.headers);
```

In `src/page-write.ts`, in `patchPage` after its fetch resolves (the same statement that feeds `requestError`):

```ts
    recordClockSkew(response.headers);
```

`graph-structure.ts` and `page-write.ts` already import `parseRetryAfter` from `./graph-throttle.ts`, so extend the existing clause; `page-content.ts` needs `recordClockSkew` added to its own.

- [ ] **Step 5: Write the failing overlap-save test**

Add to `test/mirror-sync.test.ts`:

```ts
test('a page only the overlap surfaced is logged with its age', async () => {
  const lines: string[] = [];
  setEventSink((line) => lines.push(line));

  // The watermark is 10:00 and the page's stamp is 09:30, so the only reason this page is
  // in the listing at all is the overlap. It is also genuinely changed, so it was worth
  // the window — which is the number that says how narrow the window may safely get.
  const h = harness(
    { changed: { 'sec-1': [summary('p1', '2026-08-19T09:30:00Z')] } },
    { sections: [section({ pagesSyncedThrough: '2026-08-19T10:00:00Z' })] },
  );

  await runIncremental(h.deps, BUDGET);

  const saves = lines.map((l) => JSON.parse(l)).filter((e) => e.event === 'sync-overlap-save');
  assert.deepEqual(saves, [{ event: 'sync-overlap-save', ageMs: 1_800_000 }]);

  setEventSink(() => {});
});

test('a page inside the watermark is not logged as an overlap save', async () => {
  const lines: string[] = [];
  setEventSink((line) => lines.push(line));

  const h = harness(
    { changed: { 'sec-1': [summary('p1', '2026-08-19T10:30:00Z')] } },
    { sections: [section({ pagesSyncedThrough: '2026-08-19T10:00:00Z' })] },
  );

  await runIncremental(h.deps, BUDGET);

  assert.deepEqual(
    lines.map((l) => JSON.parse(l)).filter((e) => e.event === 'sync-overlap-save'),
    [],
  );

  setEventSink(() => {});
});
```

- [ ] **Step 6: Run to see them fail**

Run: `npm test -- --test-name-pattern="overlap"`
Expected: FAIL — the assertions find no `sync-overlap-save` line

- [ ] **Step 7: Make `syncPage` report, and log in `syncSection`**

Change `syncPage`'s signature and its three exits in `src/mirror-sync.ts`:

```ts
/** True when the page was written. False on a 404, a failure, or an unchanged copy. */
async function syncPage(
  ctx: PassContext,
  section: MirrorSection,
  summary: PageSummary,
): Promise<boolean> {
  ctx.budget.take();

  let raw: RawPageContent;
  try {
    raw = await ctx.deps.content.fetchRaw(summary.id);
  } catch (err) {
    if (err instanceof GraphRequestError && err.status === 404) {
      await deletePage(ctx, summary.id, section, 'not-found');
      return false;
    }
    logEvent('mirror-page-failed', { pageId: summary.id, reason: reasonOf(err) });
    ctx.tally.pagesFailed += 1;
    return false;
  }

  try {
    if (await storePage(ctx, section, summary, raw)) {
      ctx.tally.pagesUpdated += 1;
      return true;
    }
    return false;
  } catch (err) {
    logEvent('mirror-page-failed', { pageId: summary.id, reason: reasonOf(err) });
    ctx.tally.pagesFailed += 1;
    return false;
  }
}
```

In `syncSection`'s loop:

```ts
    const written = await syncPage(ctx, section, summary);
    if (written) logOverlapSave(section, summary);
```

And add below `syncSection`:

```ts
/**
 * Record a page the overlap window is the only reason this run saw.
 *
 * The overlap subtracted in `overlapFrom` is a margin against Graph's propagation lag and
 * the gap between its clock and this process's, and it was never measured. A page whose
 * stamp predates the watermark and that turned out to have really changed is one the
 * window caught; the largest `ageMs` ever logged is the smallest width that would still
 * have caught everything.
 *
 * The age alone, with no page id and no title: this is a distribution to read off a
 * log-based metric, not a record of who edited what.
 */
function logOverlapSave(section: MirrorSection, summary: PageSummary): void {
  if (section.pagesSyncedThrough === null) return;

  const watermark = Date.parse(section.pagesSyncedThrough);
  const modified = Date.parse(summary.lastModifiedDateTime);
  if (Number.isNaN(watermark) || Number.isNaN(modified)) return;
  if (modified >= watermark) return;

  logEvent('sync-overlap-save', { ageMs: watermark - modified });
}
```

- [ ] **Step 8: Run the tests**

Run: `npm run typecheck && npm test`
Expected: PASS

- [ ] **Step 9: Commit**

```bash
git add src/graph-throttle.ts src/graph-structure.ts src/page-content.ts src/page-write.ts src/mirror-sync.ts test/
git commit -m "feat: measure the clock skew and what the overlap window catches

Every cross-clock comparison in the sync is sized on an assumption about
how far this process's clock sits from Graph's, and nothing measured it.
The Date header is on every response, so recordClockSkew costs no request.
It is whole-second resolution, hence a five-second threshold.

sync-overlap-save is the other half: a page whose stamp predates the
section watermark is one only the overlap surfaced, and logging its age
when it turned out to have really changed gives the distribution that says
how narrow the window may safely get. Both events carry a number and
nothing else."
```

---

### Task 7: Skip the content fetch for a page whose stamp has not moved

**Goal:** A page listed only because it fell inside the overlap window costs no Graph request.

**Why:** `listPagesChangedSince` already returns `id,title,lastModifiedDateTime` (`src/graph-structure.ts:540`), and `syncPage` then spends a request on `fetchRaw` for every listed page (`:786`). The content-hash short-circuit runs afterwards, so it saves the Firestore write, the GCS write and the resvg render — but not the Graph request, which is the resource under the 400/hour cap. With a one-hour window every page edited in the last hour is re-fetched by every run for the following hour.

**Safety:** both sides of the equality are Graph's own string re-read from the same collection, so it is a same-clock comparison. The settle guard is a separate predicate for a separate question: Graph stamps whole seconds, and a content fetch landing in the same second as a concurrent edit would store older content under a stamp that never moves. The guard compares Firestore's `contentSyncedAt` against Graph's stamp, which does cross clocks, hence 30 seconds rather than 2. It only ever forces an extra fetch — it never defers work into a later window — so it cannot interact with the overlap.

**Files:**
- Modify: `src/mirror-schema.ts` — `TIMESTAMP_SETTLE_MS`, `storedPageIsCurrent`
- Modify: `src/mirror-sync.ts` — `syncSection` pre-check, `syncPage` and `writePageFromRaw` accept the already-read document, `pagesSkipped` in the report
- Test: `test/mirror-schema.test.ts`, `test/mirror-sync.test.ts`

**Acceptance Criteria:**
- [ ] An unchanged page inside the overlap window costs no `fetchRaw` and no Firestore write
- [ ] A changed stamp, a changed title, a moved section, a non-`present` state and an unsettled copy each force the fetch
- [ ] The skip reads the page document once, not twice — the pre-check hands it to `writePageFromRaw`
- [ ] `pagesSkipped` appears in the report and in the `sync-completed` line

**Verify:** `npm run typecheck && npm test` → 0 failures

**Steps:**

- [ ] **Step 1: Write the failing predicate test**

Add to `test/mirror-schema.test.ts`:

```ts
test('storedPageIsCurrent needs an identical stamp and a settled copy', () => {
  const stored = {
    id: 'p1',
    title: 'Page',
    titleLower: 'page',
    sectionId: 'sec-1',
    notebookId: 'nb-1',
    sectionPath: '2026 / Daily',
    lastModifiedDateTime: '2026-08-19T12:00:00Z',
    contentState: 'present' as const,
    contentHash: 'h',
    htmlLocation: 'firestore' as const,
    htmlObject: null,
    htmlBytes: 10,
    ink: null,
    contentSyncedAt: '2026-08-19T12:01:00Z',
  };
  const live = { id: 'p1', title: 'Page', lastModifiedDateTime: '2026-08-19T12:00:00Z' };

  assert.equal(storedPageIsCurrent(stored, live, 'sec-1'), true);

  // Each of these is a way a careless simplification would answer "current" wrongly.
  assert.equal(
    storedPageIsCurrent(stored, { ...live, lastModifiedDateTime: '2026-08-19T12:00:01Z' }, 'sec-1'),
    false,
    'one second later is a real edit',
  );
  assert.equal(
    storedPageIsCurrent(stored, { ...live, title: 'Renamed' }, 'sec-1'),
    false,
    'a rename changes no content and every listing',
  );
  assert.equal(storedPageIsCurrent(stored, live, 'sec-2'), false, 'the page moved section');
  assert.equal(
    storedPageIsCurrent({ ...stored, contentState: 'stale' }, live, 'sec-1'),
    false,
    'a write marked it stale',
  );
  assert.equal(
    storedPageIsCurrent({ ...stored, contentSyncedAt: '2026-08-19T12:00:05Z' }, live, 'sec-1'),
    false,
    'fetched five seconds after the stamp, so an edit inside that second is possible',
  );
  assert.equal(
    storedPageIsCurrent({ ...stored, contentSyncedAt: undefined }, live, 'sec-1'),
    false,
    'a document written before the field existed cannot prove it settled',
  );
  assert.equal(
    storedPageIsCurrent(stored, { ...live, lastModifiedDateTime: '' }, 'sec-1'),
    false,
    'an absent live stamp proves nothing',
  );
});
```

- [ ] **Step 2: Run to see it fail**

Run: `npm test -- --test-name-pattern="storedPageIsCurrent"`
Expected: FAIL — `storedPageIsCurrent is not defined`

- [ ] **Step 3: Add the constant and the predicate**

Add to `src/mirror-schema.ts`, after `pageHasDrifted`:

```ts
/**
 * How long a stored copy must have been in hand before its stamp is trusted.
 *
 * Graph stamps `lastModifiedDateTime` to the whole second — measured 2026-08-19, see the
 * section roll-up probe in api-overview.md. So a content fetch landing in the same second
 * as a concurrent edit stores the older content under a stamp that never moves, and an
 * equality test alone would skip that page for ever. Requiring the stored copy to have
 * been fetched well after the stamp closes that window.
 *
 * The comparison crosses clocks — Firestore's `serverTimestamp()` against Graph's stamp —
 * which is why this is 30 seconds rather than 2. It costs nothing in practice: the inline
 * refresh cannot run more often than that anyway, so at most one extra fetch per edit.
 * Failing it costs a request; failing to apply it costs data.
 */
export const TIMESTAMP_SETTLE_MS = 30_000;

/**
 * Does the mirror already hold this page exactly as Graph describes it?
 *
 * True means the content fetch can be skipped. Every clause is a way that answering true
 * carelessly would serve superseded content as current:
 *
 * - `contentState` — a write already marked this stale, or it was never fetched.
 * - the stamp — same-clock equality against Graph's own string. A tolerance here would
 *   swallow a real edit, and the watermark then advances past it permanently.
 * - the title — `update_page_title` changes a title and no content, and every listing and
 *   by-name lookup matches on it.
 * - the section — page ids survive a move, so only the placement changed.
 * - the settle guard — see `TIMESTAMP_SETTLE_MS`.
 */
export function storedPageIsCurrent(
  stored: MirrorPage,
  live: PageStamp,
  sectionId: string,
  settleMs: number = TIMESTAMP_SETTLE_MS,
): boolean {
  if (stored.contentState !== 'present') return false;
  if (live.lastModifiedDateTime === '') return false;
  if (stored.lastModifiedDateTime !== live.lastModifiedDateTime) return false;
  if (stored.title !== live.title) return false;
  if (stored.sectionId !== sectionId) return false;

  const syncedAt = timestampToIso(stored.contentSyncedAt);
  if (syncedAt === null) return false;

  const synced = Date.parse(syncedAt);
  const modified = Date.parse(live.lastModifiedDateTime);
  if (Number.isNaN(synced) || Number.isNaN(modified)) return false;

  return synced - modified > settleMs;
}
```

- [ ] **Step 4: Run the predicate test**

Run: `npm test -- --test-name-pattern="storedPageIsCurrent"`
Expected: PASS

- [ ] **Step 5: Write the failing sync tests**

Add to `test/mirror-sync.test.ts`:

```ts
test('an unchanged page inside the overlap window costs no Graph request', async () => {
  const fetched: string[] = [];
  const stored = storedPage({
    id: 'p1',
    title: 'Page p1',
    lastModifiedDateTime: '2026-08-19T11:30:00Z',
    contentSyncedAt: '2026-08-19T11:35:00Z',
  });

  const h = harness(
    { changed: { 'sec-1': [summary('p1', '2026-08-19T11:30:00Z')] } },
    { sections: [section()], pages: new Map([['p1', stored]]) },
    {
      fetchRaw: (pageId) => {
        fetched.push(pageId);
        return Promise.resolve(rawHtml('<p>x</p>'));
      },
    },
  );

  const report = await runIncremental(h.deps, BUDGET);

  assert.deepEqual(fetched, [], 'the stamp did not move, so nothing needs re-reading');
  assert.deepEqual(h.storeCalls.puts, []);
  assert.equal(report.pagesSkipped, 1);
  assert.equal(report.graphRequests, 2, 'the tree and the listing, and nothing else');
  assert.deepEqual(h.storeCalls.watermarks, [{ sectionId: 'sec-1', watermark: NOW_ISO }]);
});

test('an unsettled copy is fetched even though the stamps match', async () => {
  const fetched: string[] = [];
  const stored = storedPage({
    id: 'p1',
    title: 'Page p1',
    lastModifiedDateTime: '2026-08-19T11:30:00Z',
    // Five seconds after the stamp: an edit inside that whole second is still possible.
    contentSyncedAt: '2026-08-19T11:30:05Z',
  });

  const h = harness(
    { changed: { 'sec-1': [summary('p1', '2026-08-19T11:30:00Z')] } },
    { sections: [section()], pages: new Map([['p1', stored]]) },
    {
      fetchRaw: (pageId) => {
        fetched.push(pageId);
        return Promise.resolve(rawHtml('<p>x</p>'));
      },
    },
  );

  const report = await runIncremental(h.deps, BUDGET);

  assert.deepEqual(fetched, ['p1']);
  assert.equal(report.pagesSkipped, 0);
});

test('a skipped page is read from Firestore once, not twice', async () => {
  const reads: string[] = [];
  const stored = storedPage({
    id: 'p1',
    title: 'Page p1',
    lastModifiedDateTime: '2026-08-19T11:30:00Z',
    contentSyncedAt: '2026-08-19T11:35:00Z',
  });

  const h = harness(
    { changed: { 'sec-1': [summary('p1', '2026-08-19T11:30:00Z')] } },
    { sections: [section()], pages: new Map([['p1', stored]]) },
  );
  const inner = h.deps.store.getPage;
  h.deps.store.getPage = (pageId) => {
    reads.push(pageId);
    return inner(pageId);
  };

  await runIncremental(h.deps, BUDGET);

  assert.deepEqual(reads, ['p1']);
});
```

- [ ] **Step 6: Run to see them fail**

Run: `npm test -- --test-name-pattern="overlap window costs no Graph"`
Expected: FAIL — `fetched` is `['p1']` and `report.pagesSkipped` is undefined

- [ ] **Step 7: Add the pre-check and thread the document through**

In `src/mirror-sync.ts`, replace `syncSection`'s page loop:

```ts
  for (const summary of changed) {
    // Checked before each page rather than only per section: 120 resvg renders on a
    // 1-CPU instance is what is most likely to overrun the wall clock.
    if (ctx.budget.exhausted) {
      ctx.tally.done = false;
      return;
    }

    // The saving this whole pass exists for. The listing already carried the stamp, and
    // an identical one means Graph believes nothing has changed since the copy in hand
    // was fetched — so the content request would spend a request from the hourly 400 to
    // learn what the listing already said. The document read here is handed to
    // `writePageFromRaw` so a page that is not skipped is still read from Firestore once.
    const stored = await ctx.deps.store.getPage(summary.id);
    if (stored !== null && storedPageIsCurrent(stored, summary, section.id)) {
      ctx.tally.pagesSkipped += 1;
      continue;
    }

    const written = await syncPage(ctx, section, summary, stored);
    if (written) logOverlapSave(section, summary);
  }
```

Thread the document down. `syncPage`:

```ts
async function syncPage(
  ctx: PassContext,
  section: MirrorSection,
  summary: PageSummary,
  stored?: MirrorPage | null,
): Promise<boolean> {
```

and its call into `storePage`:

```ts
    if (await storePage(ctx, section, summary, raw, stored)) {
```

`storePage`:

```ts
async function storePage(
  ctx: PassContext,
  section: MirrorSection,
  summary: PageSummary,
  raw: RawPageContent,
  stored?: MirrorPage | null,
): Promise<boolean> {
  return writePageFromRaw(
    { store: ctx.deps.store, blobs: ctx.deps.blobs },
    placementOf(section),
    summary,
    raw,
    () => {
      ctx.tally.pagesFailed += 1;
    },
    stored,
  );
}
```

`writePageFromRaw` gains a sixth parameter and stops re-reading when it is given one:

```ts
export async function writePageFromRaw(
  deps: PageWriteDeps,
  placement: PagePlacement,
  summary: PageSummary,
  raw: RawPageContent,
  onInkFailure: () => void = () => {},
  /**
   * The stored document, when the caller has already read it. `undefined` means "read it
   * here"; `null` means "there is none". The sync's pre-check has always just read it,
   * and a second read would double the Firestore cost of every page it does not skip.
   */
  known?: MirrorPage | null,
): Promise<boolean> {
  const html = pageHtml(raw) ?? '';
  const hash = createHash('sha256').update(html).digest('hex');

  const stored = known === undefined ? await deps.store.getPage(summary.id) : known;
```

Add `pagesSkipped: number` to `Tally` (initialised 0 in `runMode`), to `SyncReport`, to the report object, and to the `sync-completed` log line. Import `storedPageIsCurrent` from `./mirror-schema.ts`.

- [ ] **Step 8: Run the tests**

Run: `npm run typecheck && npm test`
Expected: PASS

- [ ] **Step 9: Commit**

```bash
git add src/mirror-schema.ts src/mirror-sync.ts test/
git commit -m "perf: skip the content fetch when Graph's page stamp has not moved

The changed-page listing already returns lastModifiedDateTime, and
syncPage then spent a Graph request on every page it named. The
content-hash short-circuit ran afterwards, so it saved the Firestore
write, the blob writes and the resvg render — not the request, which is
the resource under the 400/hour cap. With an hour of overlap every page
edited in the last hour was re-fetched by every run for the next hour.

Both sides of the equality are Graph's own string re-read from the same
collection, so it adds no assumption the \$filter listing does not already
make. The settle guard is a separate question: Graph stamps whole
seconds, so a fetch landing in the same second as a concurrent edit would
store older content under a stamp that never moves. Thirty seconds
because that guard crosses clocks; it only ever forces an extra fetch."
```

---

### Task 8: The section-scan overlap drops to fifteen minutes

**Goal:** A section edited once stops being re-listed on every run for an hour afterwards.

**Why now and not earlier:** with Task 7 in place the page-listing window costs nothing but bytes — the pages it surfaces are skipped rather than fetched — so there is no reason to narrow it. The section-scan window still costs one listing request per edited section per run, and that is the only remaining term worth cutting. Task 4 makes an over-narrow window recoverable within a night, Task 6 makes it measurable, and Task 3 means a re-activated notebook no longer depends on this window at all.

**Files:**
- Modify: `src/mirror-schema.ts` — add `SECTION_SCAN_OVERLAP_MS`
- Modify: `src/mirror-sync.ts:687` — `pickCandidates` passes it
- Modify: `README.md` — the request-budget table
- Test: `test/mirror-sync.test.ts`

**Acceptance Criteria:**
- [ ] `pickCandidates` uses a 15-minute cutoff and the page listing still uses an hour
- [ ] A section whose stamp is 20 minutes older than the scan watermark is not a candidate
- [ ] A section whose stamp is 10 minutes older still is
- [ ] A never-synced section is a candidate whatever its stamp
- [ ] A section in a wide-scan notebook is a candidate whatever its stamp

**Verify:** `npm run typecheck && npm test` → 0 failures

**Steps:**

- [ ] **Step 1: Write the failing test**

Add to `test/mirror-sync.test.ts`, beside the existing `pickCandidates` tests:

```ts
test('the section scan window is fifteen minutes, not an hour', () => {
  const state = { ...initialSyncState(), sectionsScannedThrough: '2026-08-19T12:00:00Z' };

  const inside = section({ id: 'inside', graphLastModifiedDateTime: '2026-08-19T11:50:00Z' });
  const outside = section({ id: 'outside', graphLastModifiedDateTime: '2026-08-19T11:40:00Z' });
  const never = section({
    id: 'never',
    graphLastModifiedDateTime: '2020-01-01T00:00:00Z',
    pagesSyncedThrough: null,
  });
  const widened = section({
    id: 'widened',
    notebookId: 'nb-2',
    graphLastModifiedDateTime: '2020-01-01T00:00:00Z',
  });

  // An empty live map means every section keeps its stored timestamp, which is what this
  // test is about — `withLiveMtimes` has its own test and is not the subject here.
  assert.deepEqual(
    pickCandidates(
      [inside, outside, never, widened],
      new Map(),
      state,
      true,
      new Set(['nb-2']),
    ).map((s) => s.id),
    ['inside', 'never', 'widened'],
  );
});

test('the page listing window is still an hour', () => {
  // The two windows are separate constants because they cost different things. Task 7
  // means a page the listing surfaces and that has not changed costs no request, so
  // there is nothing to gain by narrowing this one and a margin to lose.
  assert.equal(overlapFrom('2026-08-19T12:00:00Z'), '2026-08-19T11:00:00.000Z');
});
```

- [ ] **Step 2: Run to see it fail**

Run: `npm test -- --test-name-pattern="fifteen minutes"`
Expected: FAIL — `outside` is still a candidate under the hour-wide cutoff

- [ ] **Step 3: Add the constant**

Add to `src/mirror-schema.ts`, directly below `WATERMARK_OVERLAP_MS`:

```ts
/**
 * How far back the *section scan* reaches beyond `sectionsScannedThrough`.
 *
 * Separate from `WATERMARK_OVERLAP_MS` because the two windows cost different things. The
 * page-listing window pulls unchanged pages into a listing, and `storedPageIsCurrent`
 * skips them without a request — so widening it costs bytes and nothing else, and it
 * keeps the full hour of margin. This one keeps an edited section a candidate, at one
 * listing request per run for as long as it lasts, which is the only remaining term worth
 * cutting.
 *
 * Fifteen minutes rather than an hour. Both windows cover the same two things — Graph's
 * propagation lag and the gap between its clock and this process's — and neither was ever
 * measured; `sync-overlap-save` and `graph-clock-skew` in the logs are what will say
 * whether this is generous or tight. It is safe to be wrong by a margin because the
 * nightly sweep compares stamps and marks a drifted page stale, so a missed edit costs a
 * night rather than being lost. A notebook that was just mirrored or just activated does
 * not depend on this window at all — `notebooksNeedingWideScan` covers that case.
 */
export const SECTION_SCAN_OVERLAP_MS = 900_000;
```

- [ ] **Step 4: Use it**

In `src/mirror-sync.ts`, in `pickCandidates`:

```ts
  const since = overlapFrom(state.sectionsScannedThrough, SECTION_SCAN_OVERLAP_MS);
```

Add `SECTION_SCAN_OVERLAP_MS` to the `import` list from `./mirror-schema.ts`.

- [ ] **Step 5: Run the tests**

Run: `npm run typecheck && npm test`
Expected: PASS

- [ ] **Step 6: Update the README's cost table**

In `README.md`, under **What it costs**, replace the inline-refresh rows and add the notes:

```markdown
| Inline refresh, nothing changed | 1 — the expanded tree, and nothing else |
| Inline refresh, a few edited pages | 2–12; 12 is a hard budget |
| Nightly `/sync/sweep/full` over the active notebooks | ~70 here, plus one per changed page |
| `/sync/sweep` over sections whose timestamp moved | one per section visited |
| `/sync/sweep/all` over every mirrored notebook | ~202 here — half the hourly budget |
| One backfill run | 49–120 measured; it stops on whichever budget binds first |

**A page that has not changed costs no request.** The changed-page listing carries
`lastModifiedDateTime`, and a stored copy whose stamp is identical and that has been in
hand for `TIMESTAMP_SETTLE_MS` is not re-fetched — so the hour of page-listing overlap
costs bytes rather than requests, and a run's cost tracks pages that really changed rather
than pages edited in the last hour. `pagesSkipped` in the run report is how many that was.

**The two overlap windows are different widths on purpose.** `WATERMARK_OVERLAP_MS` is an
hour and governs which pages a section's listing returns; nothing it surfaces costs a
request. `SECTION_SCAN_OVERLAP_MS` is fifteen minutes and governs how long an edited
section keeps being visited, at one listing request per run. Both are margins against
Graph's propagation lag and clock skew, and `sync-overlap-save` and `graph-clock-skew` in
the logs are what say whether fifteen minutes is right.

**Editing the selection costs only the notebooks you edited.** Adding a notebook to
`notebookIds` or to `activeNotebookIds` writes only that notebook's structure documents
and makes only its sections eligible for a wide scan; every other notebook keeps its
watermark and is not listed. Removing one costs nothing at all.
```

- [ ] **Step 7: Commit**

```bash
git add src/mirror-schema.ts src/mirror-sync.ts test/ README.md
git commit -m "perf: narrow the section scan window to fifteen minutes

The overlap was one constant covering two windows that cost different
things. The page-listing window pulls unchanged pages into a listing, and
since the previous commit those cost no request — so it keeps the full
hour of margin. The section-scan window keeps an edited section a
candidate at one listing request per run, and an hour of that was the last
term worth cutting.

Both cover Graph's propagation lag and the gap between its clock and this
process's, and neither was ever measured. Being wrong is now recoverable:
the nightly sweep compares stamps and marks a drifted page stale, so a
missed edit costs a night. A re-activated notebook no longer depends on
this window at all. sync-overlap-save says whether fifteen minutes was
generous."
```

---

## After the plan: what to watch

These are operator steps, not implementation tasks. Do them once the change is deployed.

- [ ] Read `graph-clock-skew` for a week. If nothing is logged, the two clocks are inside five seconds and `TIMESTAMP_SETTLE_MS` could come down from 30 seconds. If it fires regularly, raise the settle guard rather than lowering it.
- [ ] Read the distribution of `ageMs` on `sync-overlap-save`. The largest value ever seen is the smallest safe `SECTION_SCAN_OVERLAP_MS`. If nothing is ever logged, the overlap has never caught anything and both windows could narrow further. If values above 15 minutes appear, widen it back.
- [ ] Read `pagesSkipped` against `pagesUpdated` in the `sync-completed` line. A run where `pagesSkipped` dominates is the change working; a run where it is always zero means either nothing is inside the overlap or the settle guard is never satisfied, and the second would show up as `graph-clock-skew` lines.
- [ ] After the first deploy, confirm `sync/state.wideScanNotebookIds` is `[]` and both `*IdsSeen` fields match the selection document. A set that never empties means runs are not completing.
- [ ] Confirm against `api-overview.md` whether page `lastModifiedDateTime` carries fractional seconds. All the recorded samples are whole seconds and they all come from section reads. If pages carry milliseconds, record it there — the settle guard exists only because of the one-second resolution.

## Open gaps this plan does not close

- Nothing measures how long Graph takes to make an edit visible in a filtered section listing. That lag is what both overlap windows exist for, and `sync-overlap-save` measures its effect rather than the lag itself.
- `src/mirror-store.ts` still has no automated test, so `#replaceCollection`'s merge and skip, `listPageDigestsInSection`'s projection and `putPageMetadata`'s merge are confirmed only by a live run. That is the standing rule for that file, not a new gap — but Task 2's change to `#replaceCollection` is the riskiest untested code in the plan, and the store fake in `test/mirror-sync.test.ts` deliberately re-implements its rules so at least the algorithm above it is exercised.
- Exact-match section candidacy — skipping a section whose live roll-up equals the value recorded at its last sync — is deliberately **not** in this plan. It would remove the remaining per-section listing request, and it would convert a page-listing lag from a delay into a permanent miss, because the repeated visits inside the overlap window are what give a lagging listing a second chance. Task 8 gets most of the same saving by narrowing the window instead, which keeps the retries. Revisit only if `sync-overlap-save` comes back empty over a long run.
- **`append_to_page`'s ink-clearance read is deliberately left alone.** It costs one Graph request per append (`src/write-tools.ts:567`), and the mirror holds both the raw HTML and the InkML in GCS, so it could be served from there when the page is `present` and the inline refresh reported `current`. It is not in this plan because the failure mode is the one the read exists to prevent: if the mirrored copy is behind, the padding is computed against the wrong ink bounds and the appended text renders across someone's handwriting, with a 204 and no error. Revisit after Tasks 1–8 have run against the real account long enough for `pagesSkipped` and `sync-overlap-save` to show the mirror keeping up.
- **Graph's `$batch` would not help this.** It reduces round trips and pressure on the 5-concurrent and 120-per-minute limits, and each inner request still counts against the 400 per hour — which is the limit every decision in this plan is sized against.
