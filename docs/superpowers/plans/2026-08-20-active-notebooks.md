# Active Notebooks Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers-extended-cc:subagent-driven-development (recommended) or superpowers-extended-cc:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Let an operator mark some mirrored notebooks "inactive" so they are backfilled once and then skipped by every sync, while writes and reads against them keep working.

**Architecture:** A second optional id list in the hand-edited Firestore selection document. The sync filters its section list by it in one new pure function; the read path turns "this answer came from an inactive notebook" into a third `source` value, `best-available`. Nothing about activity is stored on a notebook, section or page document, and nothing enters the structure hash.

**Tech Stack:** TypeScript on Node 22+ with native type stripping (`node --test`, no build), Express 5, Firestore, Google Cloud Storage, Microsoft Graph.

**Spec:** `docs/superpowers/specs/2026-08-20-active-notebooks-design.md`

---

## House rules for every task in this plan

Read these before Task 1. They are the repository's conventions, from `CLAUDE.md`, and they are not restated in each task.

- **Import specifiers carry `.ts`.** `import { isActive } from './mirror-schema.ts'`. Tests reach into source as `'../src/mirror-schema.ts'`.
- **The source must be erasable.** No `enum`, no `namespace`, no constructor parameter properties. Type-only imports are written `import type`.
- **Run both gates before every commit:** `npm run typecheck` then `npm test`. There is no ESLint.
- **No error message, log field, or thrown string may contain a notebook, section, or page name.** Counts are fine; ids and names are not. This repository's output can reach a public log.
- **Commit messages carry the reasoning.** What the implementation revealed that the plan did not anticipate goes in the commit message, not in a new document.
- Commit message trailer:
  ```
  Co-Authored-By: Claude Opus 5 (1M context) <noreply@anthropic.com>
  ```

## File structure

| File | Responsibility | Tasks |
|---|---|---|
| `src/mirror-schema.ts` | The selection shape, `isActive`, `activeSelectionHashOf`, the two new sync-state fields, `SyncMode` | 1, 4 |
| `src/mirror-sync.ts` | `splitByActivity`, the activation-change rescan, the three sweep entry points, the two new report fields | 2, 3, 4 |
| `src/sync-route.ts` | The fourth route | 4 |
| `src/tools.ts` | Binding `runSweepAll` | 4 |
| `src/mirror-reader.ts` | `MirrorSource`, `InactiveCoverage`, the activity snapshot, `sourceFor`, `readSourced` | 5 |
| `src/mcp-tools.ts` | `SOURCE_NOTE`, the one description of the three source values | 6 |
| `src/structure-tools.ts` | Wiring coverage into seven tools, `pagesActive`, `inactiveNotebooks` | 6 |
| `src/page-tools.ts` | Wiring coverage into `get_page_content` | 6 |
| `CLAUDE.md`, `README.md`, `project-spec.md` | Documentation | 7 |

---

### Task 1: The selection schema learns about activity

**Goal:** `readSelection` reads an optional `activeNotebookIds`, `isActive` answers from it, and the sync state carries the two fields later tasks need.

**Files:**
- Modify: `src/mirror-schema.ts`
- Test: `test/mirror-schema.test.ts`

**Acceptance Criteria:**
- [ ] `readSelection` returns `activeNotebookIds: null` when the field is absent or is not an array, and an array of cleaned ids when it is one.
- [ ] `readSelection` returns `activeNotebookIds: []` for an explicitly empty array — distinct from absent.
- [ ] `isActive` answers true for every notebook when `activeNotebookIds` is absent or null, and only for listed ids otherwise.
- [ ] `activeSelectionHashOf` gives a stable hash, equal for two selections with the same ids in different orders, and different for absent vs `[]`.
- [ ] `MirrorSyncState` carries `activeSelectionHash` and `unknownActiveNotebookIds`, defaulted by `initialSyncState` and `readSyncState`.
- [ ] Existing `NotebookSelection` literals such as `{ notebookIds: [NB] }` still compile.

**Verify:** `npm run typecheck && npm test` → both pass, no new failures.

**Steps:**

- [ ] **Step 1: Write the failing tests**

Append to `test/mirror-schema.test.ts`. Add `isActive` and `activeSelectionHashOf` to the existing import block from `'../src/mirror-schema.ts'`.

```ts
// ---------------------------------------------------------------------------
// The active set
// ---------------------------------------------------------------------------

test('an absent activeNotebookIds means every selected notebook is active', () => {
  const selection = readSelection({ notebookIds: ['nb-1', 'nb-2'] });

  assert.equal(selection.activeNotebookIds, null);
  assert.equal(isActive(selection, 'nb-1'), true);
  assert.equal(isActive(selection, 'nb-2'), true);
});

test('an activeNotebookIds that is not an array fails open, not closed', () => {
  // A malformed value must not silently freeze the mirror. Failing open costs Graph
  // requests; failing closed stops the mirror updating with nothing to say so.
  for (const raw of ['nb-1', 42, {}, null]) {
    const selection = readSelection({ notebookIds: ['nb-1'], activeNotebookIds: raw });
    assert.equal(selection.activeNotebookIds, null, `for ${JSON.stringify(raw)}`);
    assert.equal(isActive(selection, 'nb-1'), true);
  }
});

test('an explicitly empty activeNotebookIds means none are active', () => {
  const selection = readSelection({ notebookIds: ['nb-1'], activeNotebookIds: [] });

  assert.deepEqual(selection.activeNotebookIds, []);
  assert.equal(isActive(selection, 'nb-1'), false);
});

test('activeNotebookIds is cleaned the way notebookIds is', () => {
  const selection = readSelection({
    notebookIds: ['nb-1', 'nb-2'],
    activeNotebookIds: [' nb-1 ', 'nb-1', '', 7, null, 'nb-2'],
  });

  assert.deepEqual(selection.activeNotebookIds, ['nb-1', 'nb-2']);
});

test('an active id naming no selected notebook is not an error', () => {
  const selection = readSelection({ notebookIds: ['nb-1'], activeNotebookIds: ['nb-9'] });

  assert.deepEqual(selection.activeNotebookIds, ['nb-9']);
  assert.equal(isActive(selection, 'nb-1'), false);
  assert.equal(isActive(selection, 'nb-9'), true);
});

test('isActive treats an absent field and a null field identically', () => {
  assert.equal(isActive({ notebookIds: ['nb-1'] }, 'nb-1'), true);
  assert.equal(isActive({ notebookIds: ['nb-1'], activeNotebookIds: null }, 'nb-1'), true);
});

test('the active-selection hash ignores order and separates absent from empty', () => {
  const a = activeSelectionHashOf({ notebookIds: [], activeNotebookIds: ['nb-2', 'nb-1'] });
  const b = activeSelectionHashOf({ notebookIds: [], activeNotebookIds: ['nb-1', 'nb-2'] });
  const all = activeSelectionHashOf({ notebookIds: [], activeNotebookIds: null });
  const none = activeSelectionHashOf({ notebookIds: [], activeNotebookIds: [] });

  assert.equal(a, b);
  assert.notEqual(all, none);
  assert.notEqual(a, all);
  assert.notEqual(a, none);
});

test('the sync state defaults the activity fields', () => {
  const initial = initialSyncState();
  assert.equal(initial.activeSelectionHash, null);
  assert.equal(initial.unknownActiveNotebookIds, 0);

  const read = readSyncState({ activeSelectionHash: 'abc', unknownActiveNotebookIds: 3 });
  assert.equal(read.activeSelectionHash, 'abc');
  assert.equal(read.unknownActiveNotebookIds, 3);

  const junk = readSyncState({ activeSelectionHash: 7, unknownActiveNotebookIds: 'x' });
  assert.equal(junk.activeSelectionHash, null);
  assert.equal(junk.unknownActiveNotebookIds, 0);
});
```

- [ ] **Step 2: Run the tests to verify they fail**

```bash
npm test 2>&1 | tail -40
```

Expected: failures naming `isActive is not a function` and `activeSelectionHashOf is not a function`.

- [ ] **Step 3: Add the schema**

In `src/mirror-schema.ts`, replace the `NotebookSelection` interface and `readSelection` with:

```ts
/** What a human writes in the root document. The service never writes this. */
export interface NotebookSelection {
  /** Graph notebook ids whose *pages* are mirrored. */
  readonly notebookIds: readonly string[];
  /**
   * Which of those are still being edited, or absent/null for "all of them".
   *
   * A notebook that is mirrored but not active is backfilled once and then skipped by
   * every sync — see `splitByActivity` in ./mirror-sync.ts. Absent and null both mean
   * "everything is active", which is what makes this field a pure addition: a deployment
   * that has never heard of it behaves exactly as it did before.
   *
   * An explicitly empty array is *not* the same as absent. It means nothing is active,
   * which is a state an operator may legitimately want, and the type is nullable rather
   * than defaulted so those two cases stay distinguishable.
   *
   * Optional as well as nullable so that a `NotebookSelection` written as a literal —
   * every one in test/mirror-sync.test.ts — still compiles and still means "all active".
   */
  readonly activeNotebookIds?: readonly string[] | null;
}

/**
 * Read the selection document, tolerating anything a person might leave in it.
 *
 * A missing document, an absent `notebookIds`, a non-array, or an array holding
 * non-strings all resolve to "mirror nothing" rather than throwing. A human edits this
 * by hand in the Firestore console, and a half-finished edit must not take the sync down
 * — it should mean no pages are mirrored this run, which is visible and recoverable.
 * Non-string and empty entries are dropped individually, so one bad row does not discard
 * the good ones beside it.
 *
 * `activeNotebookIds` is read the same way with one difference in what a malformed value
 * means: a non-array resolves to null, which is "everything is active". That direction is
 * deliberate. Failing open costs Graph requests on notebooks that did not need checking;
 * failing closed stops the mirror updating and nothing in any result would say so.
 *
 * Duplicates are collapsed. Order is preserved, because the sync reports unmatched ids
 * back and matching that against what was typed is easier in the original order.
 */
export function readSelection(data: Record<string, unknown> | undefined): NotebookSelection {
  return {
    notebookIds: readIdArray(data?.['notebookIds']) ?? [],
    activeNotebookIds: readIdArray(data?.['activeNotebookIds']),
  };
}

/** Clean one hand-typed id array, or null when the value is not an array at all. */
function readIdArray(raw: unknown): string[] | null {
  if (!Array.isArray(raw)) return null;

  const seen = new Set<string>();
  const ids: string[] = [];
  for (const entry of raw) {
    if (typeof entry !== 'string') continue;
    const trimmed = entry.trim();
    if (trimmed === '' || seen.has(trimmed)) continue;
    seen.add(trimmed);
    ids.push(trimmed);
  }

  return ids;
}

/**
 * Is this notebook one the syncs still check?
 *
 * True whenever the operator has named no active set, which is the default and the
 * pre-feature behaviour. It does not consult `notebookIds`: a notebook that is not
 * mirrored never reaches a caller that asks this.
 */
export function isActive(selection: NotebookSelection, notebookId: string): boolean {
  const active = selection.activeNotebookIds;
  return active === undefined || active === null || active.includes(notebookId);
}

/**
 * A hash of the active set, for detecting that an operator changed it.
 *
 * Over the active ids alone rather than the whole selection: a change to `notebookIds`
 * already moves the structure hash and rewrites the structure, and this one exists to
 * trigger something different — the re-examination in ./mirror-sync.ts that a
 * re-activated notebook needs. Sorted, so reordering the array is not a change. The two
 * literal prefixes keep "no active set" and "an empty active set" apart, which are
 * opposite instructions.
 */
export function activeSelectionHashOf(selection: NotebookSelection): string {
  const active = selection.activeNotebookIds;
  const hash = createHash('sha256');

  if (active === undefined || active === null) {
    hash.update('active:all');
  } else {
    hash.update('active:subset');
    for (const id of [...active].sort()) hash.update(`\0${id}`);
  }

  return hash.digest('hex');
}
```

Add the import at the top of the file, beside the others:

```ts
import { createHash } from 'node:crypto';
```

- [ ] **Step 4: Add the two sync-state fields**

In `MirrorSyncState`, after `unknownNotebookIds`:

```ts
  /**
   * `activeSelectionHashOf` of the active set the last run saw.
   *
   * A change means an operator edited the active list, and the run that notices forces a
   * full re-examination — see `applyActivationChange` in ./mirror-sync.ts.
   */
  readonly activeSelectionHash: string | null;
  /** How many active notebook ids matched no notebook. A mistyped id is silent otherwise. */
  readonly unknownActiveNotebookIds: number;
```

In `initialSyncState()`, after `unknownNotebookIds: 0,`:

```ts
    activeSelectionHash: null,
    unknownActiveNotebookIds: 0,
```

In `readSyncState()`, after `unknownNotebookIds: numberOr(data['unknownNotebookIds'], 0),`:

```ts
    activeSelectionHash: stringOrNull(data['activeSelectionHash']),
    unknownActiveNotebookIds: numberOr(data['unknownActiveNotebookIds'], 0),
```

- [ ] **Step 5: Run the gates**

```bash
npm run typecheck && npm test 2>&1 | tail -20
```

Expected: typecheck silent, tests pass.

- [ ] **Step 6: Commit**

```bash
git add src/mirror-schema.ts test/mirror-schema.test.ts
git commit -F - <<'EOF'
Read an optional active set out of the selection document

`activeNotebookIds` names the mirrored notebooks the syncs still check. Absent
and null both mean all of them, so a deployment that has never heard of the
field behaves exactly as before; an explicitly empty array means none, which is
why the type is nullable rather than defaulted to `notebookIds`.

A malformed value resolves to "everything is active" rather than to "nothing
is". Failing open costs Graph requests on notebooks that did not need checking.
Failing closed stops the mirror updating and nothing in any tool result would
say so.

`activeSelectionHashOf` hashes the active ids alone. A change to `notebookIds`
already moves the structure hash; this one exists to trigger the re-examination
a re-activated notebook needs, which is a different thing.

Co-Authored-By: Claude Opus 5 (1M context) <noreply@anthropic.com>
EOF
```

---

### Task 2: The syncs skip inactive notebooks

**Goal:** The incremental pass backfills an inactive notebook once and never lists it again; the two existing sweeps drop it entirely; the report says how many sections were declined.

**Files:**
- Modify: `src/mirror-sync.ts`
- Test: `test/mirror-sync.test.ts`

**Acceptance Criteria:**
- [ ] `splitByActivity` is exported and pure, and `pickCandidates`'s signature is unchanged.
- [ ] With `includeBackfill` true, a section in an inactive notebook with `pagesSyncedThrough === null` is eligible and one with a watermark is not.
- [ ] With `includeBackfill` false, both are dropped.
- [ ] An inactive section is skipped by `runIncremental` once backfilled, and by `runSweep` and `runFullSweep` always.
- [ ] `SyncReport.sectionsSkippedInactive` counts what was declined.
- [ ] `unknownActiveNotebookIds` reaches the report and the sync state.
- [ ] `backfillComplete` still considers every mirrored section, active or not.
- [ ] `structureHashOf` is identical for two selections differing only in `activeNotebookIds`.
- [ ] `resyncPage` writes a page in an inactive notebook, unchanged.

**Verify:** `npm run typecheck && npm test` → pass.

**Steps:**

- [ ] **Step 1: Write the failing tests**

All of these go in `test/mirror-sync.test.ts`. Add `splitByActivity` to the existing import block from `'../src/mirror-sync.ts'`.

```ts
// ---------------------------------------------------------------------------
// splitByActivity
// ---------------------------------------------------------------------------

const ACTIVE_ONLY = { notebookIds: [NB, 'nb-cold'], activeNotebookIds: [NB] };

test('an inactive section that has never been synced is still backfilled', () => {
  const sections = [
    section({ id: 'warm', notebookId: NB, pagesSyncedThrough: null }),
    section({ id: 'cold-new', notebookId: 'nb-cold', pagesSyncedThrough: null }),
    section({ id: 'cold-done', notebookId: 'nb-cold', pagesSyncedThrough: '2026-08-01T00:00:00Z' }),
  ];

  const split = splitByActivity(sections, ACTIVE_ONLY, true);

  assert.deepEqual(split.eligible.map((s) => s.id), ['warm', 'cold-new']);
  assert.equal(split.skippedInactive, 1);
});

test('a sweep drops every inactive section, backfilled or not', () => {
  const sections = [
    section({ id: 'warm', notebookId: NB, pagesSyncedThrough: null }),
    section({ id: 'cold-new', notebookId: 'nb-cold', pagesSyncedThrough: null }),
    section({ id: 'cold-done', notebookId: 'nb-cold', pagesSyncedThrough: '2026-08-01T00:00:00Z' }),
  ];

  const split = splitByActivity(sections, ACTIVE_ONLY, false);

  assert.deepEqual(split.eligible.map((s) => s.id), ['warm']);
  assert.equal(split.skippedInactive, 2);
});

test('with no active set every section is eligible and nothing is skipped', () => {
  const sections = [section({ id: 'a', notebookId: NB }), section({ id: 'b', notebookId: 'nb-2' })];

  for (const includeBackfill of [true, false]) {
    const split = splitByActivity(sections, { notebookIds: [NB, 'nb-2'] }, includeBackfill);
    assert.deepEqual(split.eligible.map((s) => s.id), ['a', 'b']);
    assert.equal(split.skippedInactive, 0);
  }
});
```

Then the run-level tests. `harness` and `section` are the file's existing helpers; follow the shape of the tests already there. Add:

```ts
test('an incremental run backfills an inactive section once and then leaves it alone', async () => {
  const h = harness(
    {},
    {
      selection: ACTIVE_ONLY,
      sections: [
        section({ id: 's-cold', notebookId: 'nb-cold', pagesSyncedThrough: null }),
      ],
    },
  );

  const first = await runIncremental(h.deps, { requestBudget: 50 });
  assert.equal(first.sectionsVisited, 1);
  assert.equal(first.sectionsSkippedInactive, 0);

  const second = await runIncremental(h.deps, { requestBudget: 50 });
  assert.equal(second.sectionsVisited, 0);
  assert.equal(second.sectionsSkippedInactive, 1);
});

test('neither sweep visits an inactive section', async () => {
  for (const run of [runSweep, runFullSweep]) {
    const h = harness(
      {},
      {
        selection: ACTIVE_ONLY,
        sections: [
          section({ id: 's-cold', notebookId: 'nb-cold', pagesSyncedThrough: null }),
        ],
      },
    );

    const report = await run(h.deps, { requestBudget: 50 });
    assert.equal(report.sectionsVisited, 0);
    assert.equal(report.sectionsSkippedInactive, 1);
  }
});

test('an active id matching no notebook is counted, not thrown', async () => {
  const h = harness({}, { selection: { notebookIds: [NB], activeNotebookIds: [NB, 'nb-typo'] } });

  const report = await runIncremental(h.deps, { requestBudget: 50 });

  assert.equal(report.unknownActiveNotebookIds, 1);
});

test('activity does not enter the structure hash', () => {
  // If it did, an activation edit would move the hash, rewrite every structure document
  // through `putStructure`, and -- because `#replaceCollection` uses `batch.set` with no
  // merge and `buildStructure` emits a null watermark -- reset every section's
  // `pagesSyncedThrough`. That is a full re-backfill of the whole selection.
  const warm = structureHashOf(buildStructure(TREE, { notebookIds: [NB], activeNotebookIds: [NB] }));
  const cold = structureHashOf(buildStructure(TREE, { notebookIds: [NB], activeNotebookIds: [] }));

  assert.equal(warm, cold);
});

test('a write resyncs a page in an inactive notebook like any other', async () => {
  // The property most likely to be broken by someone adding an activity check to
  // `resyncPage` for symmetry. Activity governs what the *syncs* re-read; a write knows
  // exactly what it changed and there is nothing to save by not recording it.
  const h = resyncHarness({
    selection: { notebookIds: ['nb-cold'], activeNotebookIds: [] },
    sections: [section({ id: 's-cold', notebookId: 'nb-cold', mirrored: true })],
  });

  const outcome = await resyncPage(h.deps, 'p-cold', { sectionId: 's-cold' });

  assert.equal(outcome, 'updated');
  assert.notEqual(h.data.pages.get('p-cold'), undefined);
});
```

> `resyncHarness` is whatever this file already calls the helper its existing `resyncPage`
> tests use; reuse it rather than adding another. If those tests build `ResyncDeps` inline,
> build them inline here too, matching their shape.

> If `harness` in this file does not already accept a `selection` override, add one: it takes a `Partial<StoreState>` whose `selection` field is already `{ notebookIds: [NB] }`, so passing a different selection is a one-word change at the call site and needs no helper change.

- [ ] **Step 2: Run the tests to verify they fail**

```bash
npm test 2>&1 | tail -40
```

Expected: `splitByActivity is not a function`, and `sectionsSkippedInactive` undefined.

- [ ] **Step 3: Add `splitByActivity`**

In `src/mirror-sync.ts`, add `isActive` to the import block from `'./mirror-schema.ts'`, then add this function immediately above `pickCandidates`:

```ts
/**
 * Which sections this run is allowed to consider at all.
 *
 * A mirrored notebook the operator has not marked active is one they have asserted is no
 * longer being edited. Its pages are still held and still served; what stops is the
 * re-checking, which is the whole cost this feature removes — an archive notebook is
 * re-listed section by section every run and finds nothing.
 *
 * `includeBackfill` is true for the incremental pass, where a section in an inactive
 * notebook that has never been synced is still visited. That is the one-off fill: the
 * notebook is read once, completely, and from then on no incremental lists it again. It
 * is false for a sweep, which reconciles page ids against a mirror that has to be filled
 * already and never backfills anything.
 *
 * Separate from `pickCandidates` rather than folded into it for two reasons. The count of
 * what was declined has to reach the report, and `pickCandidates` returns only a list. And
 * `pickCandidates` has an early return for a distrusted section roll-up, so a filter
 * folded into it would have to be applied on both sides of that branch — the shape that
 * ends with one side wrong.
 */
export function splitByActivity(
  sections: readonly MirrorSection[],
  selection: NotebookSelection,
  includeBackfill: boolean,
): { eligible: MirrorSection[]; skippedInactive: number } {
  const eligible: MirrorSection[] = [];
  let skippedInactive = 0;

  for (const section of sections) {
    if (
      isActive(selection, section.notebookId) ||
      (includeBackfill && section.pagesSyncedThrough === null)
    ) {
      eligible.push(section);
    } else {
      skippedInactive += 1;
    }
  }

  return { eligible, skippedInactive };
}
```

- [ ] **Step 4: Count it in the tally and the report**

In the `Tally` interface, after `unknownNotebookIds: number;`:

```ts
  sectionsSkippedInactive: number;
  unknownActiveNotebookIds: number;
```

In `SyncReport`, after `readonly unknownNotebookIds: number;`:

```ts
  /** Sections this run declined to visit because their notebook is not active. */
  readonly sectionsSkippedInactive: number;
  /** Active notebook ids matching no notebook. A mistyped id is silent otherwise. */
  readonly unknownActiveNotebookIds: number;
```

In `runMode`, in the `tally` literal after `unknownNotebookIds: 0,`:

```ts
    sectionsSkippedInactive: 0,
    unknownActiveNotebookIds: 0,
```

In the `report` literal after `unknownNotebookIds: report.unknownNotebookIds`'s source line `unknownNotebookIds: tally.unknownNotebookIds,`:

```ts
    sectionsSkippedInactive: tally.sectionsSkippedInactive,
    unknownActiveNotebookIds: tally.unknownActiveNotebookIds,
```

In the `patchSyncState` call at the end of `runMode`, after `unknownNotebookIds: report.unknownNotebookIds,`:

```ts
    unknownActiveNotebookIds: report.unknownActiveNotebookIds,
```

- [ ] **Step 5: Filter in the incremental pass**

In `incrementalPass`, replace

```ts
  const sections = await ctx.deps.store.listSectionsToSync();
  const candidates = pickCandidates(sections, ctx.state, ctx.tally.treeRead && !structureChanged);
```

with

```ts
  const sections = await ctx.deps.store.listSectionsToSync();
  const { eligible, skippedInactive } = splitByActivity(sections, ctx.selection, true);
  ctx.tally.sectionsSkippedInactive += skippedInactive;
  const candidates = pickCandidates(eligible, ctx.state, ctx.tally.treeRead && !structureChanged);
```

Leave the `backfillComplete` computation reading `sections`, not `eligible`. The backfill is not complete while an inactive notebook is still filling, and `sections` is the list that knows.

- [ ] **Step 6: Filter in the sweep**

In `sweepPass`, replace

```ts
  const all = await ctx.deps.store.listSectionsToSync();
  const sections = unscoped ? all : pickCandidates(all, ctx.state, ctx.tally.treeRead);
```

with

```ts
  const all = await ctx.deps.store.listSectionsToSync();
  const { eligible, skippedInactive } = splitByActivity(all, ctx.selection, false);
  ctx.tally.sectionsSkippedInactive += skippedInactive;
  const sections = unscoped ? eligible : pickCandidates(eligible, ctx.state, ctx.tally.treeRead);
```

- [ ] **Step 7: Count the unknown active ids**

In `BuiltStructure`, after `readonly unknownNotebookIds: string[];`:

```ts
  /** Active notebook ids that matched no notebook in the tree. */
  readonly unknownActiveNotebookIds: string[];
```

In `buildStructure`'s return, after `unknownNotebookIds: selection.notebookIds.filter((id) => !seen.has(id)),`:

```ts
    unknownActiveNotebookIds: (selection.activeNotebookIds ?? []).filter((id) => !seen.has(id)),
```

In `reconcileStructure`, after the existing `unknownNotebookIds` block:

```ts
  ctx.tally.unknownActiveNotebookIds = built.unknownActiveNotebookIds.length;
  if (built.unknownActiveNotebookIds.length > 0) {
    // A count and never the ids. A notebook id is opaque, but the count is all an
    // operator needs to know their active list has a typo in it.
    logEvent('mirror-selection-unknown-active', { count: built.unknownActiveNotebookIds.length });
  }
```

`structureHashOf` is **not** changed. Activity must not enter the structure hash: `putStructure` replaces documents wholesale and `buildStructure` emits `pagesSyncedThrough: null`, so anything in that hash resets every section's watermark when it moves.

- [ ] **Step 8: Run the gates**

```bash
npm run typecheck && npm test 2>&1 | tail -20
```

Expected: pass. Existing `buildStructure` tests may need `unknownActiveNotebookIds: []` added to a `deepEqual` on the whole returned object; if so, add it.

- [ ] **Step 9: Commit**

```bash
git add src/mirror-sync.ts test/mirror-sync.test.ts
git commit -F - <<'EOF'
Skip inactive notebooks in every sync but the backfill

`splitByActivity` filters the section list before `pickCandidates` sees it.
`includeBackfill` is true for the incremental pass, so a section in an inactive
notebook that has never been synced is still read — once, completely — and is
then never listed again. It is false for a sweep, which never backfills.

Kept out of `pickCandidates` on purpose. The count of what was declined has to
reach the report and `pickCandidates` returns only a list, and `pickCandidates`
has an early return for a distrusted section roll-up that a folded-in filter
would have to be applied on both sides of.

`backfillComplete` still reads the unfiltered list: the backfill is not
complete while an inactive notebook is still filling.

Activity stays out of `structureHashOf`. `putStructure` replaces documents
wholesale and `buildStructure` emits a null watermark, so anything in that hash
resets every section's watermark when it changes -- an activation edit would
trigger a full re-backfill of the whole selection.

Co-Authored-By: Claude Opus 5 (1M context) <noreply@anthropic.com>
EOF
```

---

### Task 3: Re-activation forces a re-examination

**Goal:** Changing the active set makes the next run re-examine every eligible section, so a notebook edited while it was inactive is not skipped for ever.

**Files:**
- Modify: `src/mirror-sync.ts`
- Test: `test/mirror-sync.test.ts`

**Acceptance Criteria:**
- [ ] A run whose `activeSelectionHash` differs from the stored one patches `sectionsScannedThrough: null` and the new hash, and uses the nulled value for its own candidate pick.
- [ ] A run whose hash matches patches neither.
- [ ] Per-section `pagesSyncedThrough` values are untouched by the change, so no page is re-fetched that did not change.
- [ ] Both `incrementalPass` and `sweepPass` apply it.

**Verify:** `npm run typecheck && npm test` → pass.

**Steps:**

- [ ] **Step 1: Write the failing tests**

In `test/mirror-sync.test.ts`:

```ts
// ---------------------------------------------------------------------------
// Re-activation
// ---------------------------------------------------------------------------

test('changing the active set re-examines every eligible section on the next run', async () => {
  // The failure this prevents: tier 1 skips a section whose timestamp is older than the
  // last complete scan. A section edited three months ago, while its notebook was
  // inactive, is older than that cutoff for ever -- so without this, re-activating the
  // notebook would change nothing at all.
  const h = harness(
    {},
    {
      selection: { notebookIds: [NB], activeNotebookIds: [NB] },
      state: {
        ...initialSyncState(),
        sectionsScannedThrough: '2026-08-19T11:00:00.000Z',
        activeSelectionHash: 'a-different-hash',
      },
      sections: [
        section({
          id: 's-old',
          notebookId: NB,
          graphLastModifiedDateTime: '2026-05-01T00:00:00Z',
          pagesSyncedThrough: '2026-05-01T00:00:00Z',
        }),
      ],
    },
  );

  const report = await runIncremental(h.deps, { requestBudget: 50 });

  assert.equal(report.sectionsVisited, 1);
  assert.equal(
    h.data.state.activeSelectionHash,
    activeSelectionHashOf({ notebookIds: [NB], activeNotebookIds: [NB] }),
  );
});

test('an unchanged active set does not reset the scan watermark', async () => {
  const selection = { notebookIds: [NB], activeNotebookIds: [NB] };
  const h = harness(
    {},
    {
      selection,
      state: {
        ...initialSyncState(),
        sectionsScannedThrough: '2026-08-19T11:00:00.000Z',
        activeSelectionHash: activeSelectionHashOf(selection),
      },
      sections: [
        section({
          id: 's-old',
          notebookId: NB,
          graphLastModifiedDateTime: '2026-05-01T00:00:00Z',
          pagesSyncedThrough: '2026-05-01T00:00:00Z',
        }),
      ],
    },
  );

  const report = await runIncremental(h.deps, { requestBudget: 50 });

  assert.equal(report.sectionsVisited, 0);
  assert.equal(h.data.state.sectionsScannedThrough, '2026-08-19T11:00:00.000Z');
});

test('a re-activation does not disturb per-section watermarks', async () => {
  const h = harness(
    {},
    {
      selection: { notebookIds: [NB], activeNotebookIds: [NB] },
      state: { ...initialSyncState(), activeSelectionHash: 'a-different-hash' },
      sections: [
        section({ id: 's-1', notebookId: NB, pagesSyncedThrough: '2026-05-01T00:00:00Z' }),
      ],
    },
  );

  await runIncremental(h.deps, { requestBudget: 50 });

  // The section is re-listed, but from its own watermark: only pages that changed since
  // 2026-05-01 are fetched. A reset here would re-read the whole notebook.
  assert.equal(h.listedSince.get('s-1'), '2026-05-01T00:00:00.000Z');
});
```

> The last assertion needs the harness to record the `sinceIso` its fake `listPagesChangedSince` was called with. If the harness does not already expose that, add a `listedSince: Map<string, string>` to it and record in the fake. If recording it is awkward in the existing harness shape, assert instead that `overlapFrom` was honoured by checking the fetched page ids, and say so in the commit message.

- [ ] **Step 2: Run the tests to verify they fail**

```bash
npm test 2>&1 | tail -40
```

Expected: the first test fails with `sectionsVisited` 0, the third fails on the missing recorder.

- [ ] **Step 3: Add the activation check**

In `src/mirror-sync.ts`, add `activeSelectionHashOf` to the import block from `'./mirror-schema.ts'`, then add:

```ts
/**
 * Notice that the operator changed the active set, and force a re-examination.
 *
 * Returns the state this run should use, which is not always `ctx.state`.
 *
 * The failure this exists to prevent: tier 1 skips a section unless its
 * `graphLastModifiedDateTime` is newer than `overlapFrom(sectionsScannedThrough)`, and
 * that watermark advances on every completed run. A section edited three months ago,
 * while its notebook was inactive, is older than the cutoff — and stays older for ever.
 * Moving the notebook back into `activeNotebookIds` would therefore change nothing.
 *
 * Nulling `sectionsScannedThrough` makes every eligible section a candidate for one run.
 * That is one `listPagesChangedSince` per section, each against that section's own
 * per-section watermark, which is untouched — so only pages that genuinely changed are
 * fetched, and the run is budget-bounded and resumable like any other.
 *
 * The returned state matters as much as the patch. `ctx.state` is a snapshot taken at the
 * start of the run, so patching Firestore without also correcting the local copy would
 * defer the whole effect to the *next* run.
 */
async function applyActivationChange(ctx: PassContext): Promise<MirrorSyncState> {
  const hash = activeSelectionHashOf(ctx.selection);
  if (hash === ctx.state.activeSelectionHash) return ctx.state;

  await ctx.deps.store.patchSyncState({
    activeSelectionHash: hash,
    sectionsScannedThrough: null,
  });
  // Whether an active set exists at all, never which notebooks are in it.
  logEvent('mirror-activation-changed', {
    scope: ctx.selection.activeNotebookIds == null ? 'all' : 'subset',
  });

  return { ...ctx.state, activeSelectionHash: hash, sectionsScannedThrough: null };
}
```

- [ ] **Step 4: Use it in both passes**

In `incrementalPass`, after the empty-selection early return and before `reconcileStructure`:

```ts
  const state = await applyActivationChange(ctx);
```

and change the `pickCandidates` call to pass `state` rather than `ctx.state`:

```ts
  const candidates = pickCandidates(eligible, state, ctx.tally.treeRead && !structureChanged);
```

In `sweepPass`, after its own empty-selection early return:

```ts
  const state = await applyActivationChange(ctx);
```

and change its `pickCandidates` call to pass `state`:

```ts
  const sections = unscoped ? eligible : pickCandidates(eligible, state, ctx.tally.treeRead);
```

Both passes apply it because either may be the first to run after an edit, and clearing the
hash in one while the other still holds the old view would lose the re-examination. Clearing
it from a sweep is safe: `sectionsScannedThrough` only advances when a complete incremental
finishes, so the null the sweep wrote is still there for the next incremental to see.

- [ ] **Step 5: Run the gates**

```bash
npm run typecheck && npm test 2>&1 | tail -20
```

- [ ] **Step 6: Commit**

```bash
git add src/mirror-sync.ts test/mirror-sync.test.ts
git commit -F - <<'EOF'
Re-examine everything when the active set changes

Tier 1 skips a section unless its timestamp is newer than the last complete
scan, and that watermark advances every run. A section edited while its
notebook was inactive is older than the cutoff and stays older for ever, so
moving the notebook back into `activeNotebookIds` would otherwise change
nothing at all.

A run whose active-set hash differs from the stored one nulls
`sectionsScannedThrough`, which makes every eligible section a candidate for
exactly one run. Per-section watermarks are untouched, so each section is
re-listed from where it left off and only pages that really changed are
fetched.

The function returns the state the run should use rather than only patching
Firestore. `ctx.state` is a snapshot taken at the start of the run; patching
without correcting it would defer the whole effect to the next run.

Both passes apply it, because either may be the first to run after an edit.

Co-Authored-By: Claude Opus 5 (1M context) <noreply@anthropic.com>
EOF
```

---

### Task 4: `POST /sync/sweep/all`

**Goal:** A fourth route and a fourth sync mode that sweeps every mirrored section, active or not.

**Files:**
- Modify: `src/mirror-schema.ts`, `src/mirror-sync.ts`, `src/sync-route.ts`, `src/tools.ts`
- Test: `test/mirror-sync.test.ts`, `test/sync-route.test.ts`, `test/mirror-schema.test.ts`

**Acceptance Criteria:**
- [ ] `SyncMode` includes `'sweep-all'` and `syncModeOrNull` accepts it.
- [ ] `runSweepAll` visits sections in inactive notebooks; `runFullSweep` still does not.
- [ ] `POST /sync/sweep/all` reaches `runSweepAll` and no other mode.
- [ ] `GET /sync/sweep/all` answers 405 with `Allow: POST`.
- [ ] The route is absent when `MIRROR_SYNC_SECRET` is unset.
- [ ] A request with no `x-sync-secret` gets the route's own 401 with no `WWW-Authenticate` header.

**Verify:** `npm run typecheck && npm test` → pass.

**Steps:**

- [ ] **Step 1: Write the failing tests**

In `test/mirror-schema.test.ts`:

```ts
test('sweep-all is a sync mode a stored document can carry', () => {
  assert.equal(readSyncState({ runningMode: 'sweep-all' }).runningMode, 'sweep-all');
  assert.equal(readSyncState({ runningMode: 'sweep-sideways' }).runningMode, null);
});
```

In `test/mirror-sync.test.ts` (add `runSweepAll` to the import block):

```ts
test('sweep-all visits the sections the other sweeps decline', async () => {
  const h = harness(
    {},
    {
      selection: ACTIVE_ONLY,
      sections: [section({ id: 's-cold', notebookId: 'nb-cold', pagesSyncedThrough: null })],
    },
  );

  const report = await runSweepAll(h.deps, { requestBudget: 50 });

  assert.equal(report.mode, 'sweep-all');
  assert.equal(report.sectionsVisited, 1);
  assert.equal(report.sectionsSkippedInactive, 0);
});
```

In `test/sync-route.test.ts`, follow the file's existing shape for the other three modes:

```ts
test('POST /sync/sweep/all reaches runSweepAll and no other mode', async () => {
  const calls: string[] = [];
  const target = recordingTarget(calls);
  const server = await listen(syncRouter(SECRET, target));

  const response = await fetch(`${server.url}/sweep/all`, {
    method: 'POST',
    headers: { [SYNC_HEADER]: SECRET },
  });

  assert.equal(response.status, 200);
  assert.deepEqual(calls, ['runSweepAll']);

  await server.close();
});

test('GET /sync/sweep/all is 405 with an Allow header', async () => {
  const server = await listen(syncRouter(SECRET, recordingTarget([])));

  const response = await fetch(`${server.url}/sweep/all`, { method: 'GET' });

  assert.equal(response.status, 405);
  assert.equal(response.headers.get('allow'), 'POST');

  await server.close();
});
```

> `recordingTarget` and `listen` are this file's existing helpers; use whatever names it already has. `recordingTarget` must gain a `runSweepAll` member, since `SyncTarget` now requires one.

- [ ] **Step 2: Run the tests to verify they fail**

```bash
npm test 2>&1 | tail -40
```

- [ ] **Step 3: Widen `SyncMode`**

In `src/mirror-schema.ts`:

```ts
export type SyncMode = 'incremental' | 'sweep' | 'sweep-full' | 'sweep-all';
```

and in `syncModeOrNull`:

```ts
function syncModeOrNull(value: unknown): SyncMode | null {
  return value === 'incremental' ||
    value === 'sweep' ||
    value === 'sweep-full' ||
    value === 'sweep-all'
    ? value
    : null;
}
```

- [ ] **Step 4: Add the run mode**

In `src/mirror-sync.ts`, replace the three sweep entry points with:

```ts
export async function runSweep(deps: SyncDeps, options: SyncOptions): Promise<SyncReport> {
  return runMode('sweep', deps, options, (ctx) =>
    sweepPass(ctx, { unscoped: false, ignoreActivity: false }),
  );
}

export async function runFullSweep(deps: SyncDeps, options: SyncOptions): Promise<SyncReport> {
  return runMode('sweep-full', deps, options, (ctx) =>
    sweepPass(ctx, { unscoped: true, ignoreActivity: false }),
  );
}

/**
 * Every mirrored section, whether or not its notebook is active.
 *
 * The way an operator forces a re-check of a notebook they have marked inactive without
 * editing the selection document. `sweep-full` deliberately does not do this: it is the
 * weekly backstop and it runs on a schedule, so making it ignore activity would give back
 * every request the feature saves.
 */
export async function runSweepAll(deps: SyncDeps, options: SyncOptions): Promise<SyncReport> {
  return runMode('sweep-all', deps, options, (ctx) =>
    sweepPass(ctx, { unscoped: true, ignoreActivity: true }),
  );
}
```

Change `sweepPass`'s signature and its section selection:

```ts
async function sweepPass(
  ctx: PassContext,
  options: { unscoped: boolean; ignoreActivity: boolean },
): Promise<void> {
  if (ctx.selection.notebookIds.length === 0) return;

  const state = await applyActivationChange(ctx);

  await reconcileStructure(ctx);
  await learnNestedGroups(ctx);

  const all = await ctx.deps.store.listSectionsToSync();
  const { eligible, skippedInactive } = options.ignoreActivity
    ? { eligible: all, skippedInactive: 0 }
    : splitByActivity(all, ctx.selection, false);
  ctx.tally.sectionsSkippedInactive += skippedInactive;

  const sections = options.unscoped
    ? eligible
    : pickCandidates(eligible, state, ctx.tally.treeRead);
```

The rest of the function body is unchanged.

`sweepCursorSectionId` stays one field shared by all three sweep modes. A `sweep-all`
resuming onto a cursor a scoped `sweep` left behind finds no match, `findIndex` returns -1,
and `Math.max(0, -1)` restarts from the top — the existing behaviour, and it costs a repeat
rather than a wrong answer.

- [ ] **Step 5: Add the route**

In `src/sync-route.ts`, extend `SyncTarget`:

```ts
export interface SyncTarget {
  runIncremental(): Promise<SyncReport>;
  runSweep(): Promise<SyncReport>;
  runFullSweep(): Promise<SyncReport>;
  runSweepAll(): Promise<SyncReport>;
}
```

Register it and add it to the 405 list:

```ts
  router.post('/', run('runIncremental'));
  router.post('/sweep', run('runSweep'));
  router.post('/sweep/full', run('runFullSweep'));
  router.post('/sweep/all', run('runSweepAll'));

  // Only POST. A GET would let a link preview, or anything that crawls a URL, spend a
  // slice of the hourly Graph budget.
  for (const path of ['/', '/sweep', '/sweep/full', '/sweep/all']) {
```

Extend the header comment in that file — the block explaining "Three paths rather than one
path with a mode" — to say there are now four and what the fourth is for:

```
// **Four paths rather than one path with a mode.** ... `/sync/sweep/all` is the only one
// that visits notebooks the operator has marked inactive; it is the manual lever, not a
// scheduled job, which is why `/sync/sweep/full` still skips them.
```

- [ ] **Step 6: Bind it**

In `src/tools.ts`, add `runSweepAll` to the import block from `'./mirror-sync.ts'` and to the returned object in `createSyncTargetFor`:

```ts
  return {
    runIncremental: () => runIncremental(deps, options),
    runSweep: () => runSweep(deps, options),
    runFullSweep: () => runFullSweep(deps, options),
    runSweepAll: () => runSweepAll(deps, options),
  };
```

- [ ] **Step 7: Run the gates**

```bash
npm run typecheck && npm test 2>&1 | tail -20
```

- [ ] **Step 8: Commit**

```bash
git add src/mirror-schema.ts src/mirror-sync.ts src/sync-route.ts src/tools.ts test/
git commit -F - <<'EOF'
Add POST /sync/sweep/all, the only sweep that visits inactive notebooks

`sweep-full` deliberately keeps skipping them. It is the weekly scheduled
backstop, so making it ignore activity would give back every request the
feature saves. `sweep/all` is the manual lever an operator pulls when they want
an archive re-checked without editing the selection document.

The mode is the path, for the reason the file already gives: logging.ts records
the path and records neither the query string nor the body, and "which job ran,
and did it answer?" is the first question when the mirror looks wrong.

`sweepCursorSectionId` stays one field across all three sweep modes. A sweep-all
resuming onto a scoped sweep's cursor finds no match and restarts from the top,
which costs a repeat rather than a wrong answer.

Co-Authored-By: Claude Opus 5 (1M context) <noreply@anthropic.com>
EOF
```

---

### Task 5: `best-available`, and the reader that decides it

**Goal:** `MirrorSource` gains a third value, the reader can say how much of an answer came from an inactive notebook, and `readSourced` turns the two into the label.

**Files:**
- Modify: `src/mirror-reader.ts`
- Test: `test/mirror-reader.test.ts`

**Acceptance Criteria:**
- [ ] `sourceFor(confirmed, coverage)` is exported and satisfies the full six-row truth table.
- [ ] `MirrorReadStore` requires `getSelection`.
- [ ] `coverageOfSection`, `coverageOfPage` and `accountActivity` answer from one snapshot memoised for 30 s against an injectable clock.
- [ ] `coverageOfSection` and `coverageOfPage` answer `'some'` when the document is missing.
- [ ] `readSourced` calls `inactiveCoverage` only on a mirror hit, and a throw from it yields `'some'` rather than a failed read.
- [ ] A Graph-origin answer is still `onenote` and never calls `inactiveCoverage`.

**Verify:** `npm run typecheck && npm test` → pass.

**Steps:**

- [ ] **Step 1: Write the failing tests**

In `test/mirror-reader.test.ts`, add `sourceFor` and `type InactiveCoverage` to the import block, and add `getSelection` to the `fakeStore` fixture and store:

```ts
// in interface Fixture
  selection: NotebookSelection;

// in the data literal
    selection: { notebookIds: ['nb-1'] },

// in the store literal
    getSelection: () => guard('getSelection', data.selection),
```

Then the tests:

```ts
// ---------------------------------------------------------------------------
// The source label
// ---------------------------------------------------------------------------

test('the source truth table', () => {
  // `onenote` is a claim that the answer equals what OneNote holds. `best-available` is
  // a claim that everything not confirmed came from a notebook the operator says is not
  // being edited. `mirror` is the only one that means "this may be out of date".
  assert.equal(sourceFor(true, 'none'), 'onenote');
  assert.equal(sourceFor(false, 'none'), 'mirror');
  assert.equal(sourceFor(true, 'some'), 'best-available');
  assert.equal(sourceFor(false, 'some'), 'mirror');
  // `all` ignores confirmation: a refresh that failed was never going to check this
  // notebook, so its failure says nothing about this answer.
  assert.equal(sourceFor(true, 'all'), 'best-available');
  assert.equal(sourceFor(false, 'all'), 'best-available');
});

test('a section in an inactive notebook is full coverage, an active one is none', async () => {
  const { reader: r } = reader({
    selection: { notebookIds: ['nb-1', 'nb-cold'], activeNotebookIds: ['nb-1'] },
    notebooks: [notebook({ id: 'nb-1' }), notebook({ id: 'nb-cold' })],
    sections: [
      section({ id: 's-warm', notebookId: 'nb-1' }),
      section({ id: 's-cold', notebookId: 'nb-cold' }),
    ],
  });

  assert.equal(await r.coverageOfSection('s-warm'), 'none');
  assert.equal(await r.coverageOfSection('s-cold'), 'all');
});

test('an unknown section or page is "some", never "all"', async () => {
  // "all" would be a stronger claim than the data supports: it says every uncertain part
  // of this answer came from a notebook nobody edits, and a document the mirror cannot
  // find says nothing of the kind.
  const { reader: r } = reader({
    selection: { notebookIds: ['nb-1'], activeNotebookIds: [] },
  });

  assert.equal(await r.coverageOfSection('s-nope'), 'some');
  assert.equal(await r.coverageOfPage('p-nope'), 'some');
});

test('a page in an inactive notebook is full coverage', async () => {
  const { reader: r } = reader({
    selection: { notebookIds: ['nb-cold'], activeNotebookIds: [] },
    pages: [page({ id: 'p-1', notebookId: 'nb-cold' })],
  });

  assert.equal(await r.coverageOfPage('p-1'), 'all');
});

test('account activity reports the mix and the count', async () => {
  const mixed = reader({
    selection: { notebookIds: ['nb-1', 'nb-cold'], activeNotebookIds: ['nb-1'] },
    notebooks: [
      notebook({ id: 'nb-1', mirrored: true }),
      notebook({ id: 'nb-cold', mirrored: true }),
    ],
  }).reader;
  assert.deepEqual(await mixed.accountActivity(), { coverage: 'some', inactiveNotebooks: 1 });

  const allWarm = reader({
    selection: { notebookIds: ['nb-1'] },
    notebooks: [notebook({ id: 'nb-1', mirrored: true })],
  }).reader;
  assert.deepEqual(await allWarm.accountActivity(), { coverage: 'none', inactiveNotebooks: 0 });

  const allCold = reader({
    selection: { notebookIds: ['nb-1'], activeNotebookIds: [] },
    notebooks: [notebook({ id: 'nb-1', mirrored: true })],
  }).reader;
  assert.deepEqual(await allCold.accountActivity(), { coverage: 'all', inactiveNotebooks: 1 });
});

test('an unmirrored notebook is not counted as inactive', async () => {
  // `mirrored` and `active` are different questions. A notebook whose pages were never
  // mirrored contributes nothing to an answer, so it cannot make one best-available.
  const { reader: r } = reader({
    selection: { notebookIds: ['nb-1'], activeNotebookIds: ['nb-1'] },
    notebooks: [
      notebook({ id: 'nb-1', mirrored: true }),
      notebook({ id: 'nb-other', mirrored: false }),
    ],
  });

  assert.deepEqual(await r.accountActivity(), { coverage: 'none', inactiveNotebooks: 0 });
});

test('the activity snapshot is memoised and then expires', async () => {
  let clock = 0;
  const { store, data } = fakeStore({ selection: { notebookIds: ['nb-1'] } });
  let reads = 0;
  const counting: MirrorReadStore = {
    ...store,
    getSelection: () => {
      reads += 1;
      return store.getSelection();
    },
  };
  const r = new MirrorReader(counting, fakeBlobs(), undefined, () => clock);

  await r.accountActivity();
  await r.accountActivity();
  assert.equal(reads, 1);

  data.selection = { notebookIds: ['nb-1'], activeNotebookIds: [] };
  clock = ACTIVITY_MEMO_MS + 1;
  assert.deepEqual((await r.accountActivity()).coverage, 'all');
  assert.equal(reads, 2);
});

test('readSourced asks for coverage only on a mirror hit', async () => {
  let asked = 0;
  const answer = await readSourced({
    tool: 't',
    mirror: reader().reader,
    sync: undefined,
    fromMirror: () => Promise.resolve(null),
    fromGraph: () => Promise.resolve('from graph'),
    inactiveCoverage: () => {
      asked += 1;
      return Promise.resolve<InactiveCoverage>('all');
    },
  });

  assert.equal(answer.source, 'onenote');
  assert.equal(answer.data, 'from graph');
  assert.equal(asked, 0);
});

test('coverage that throws is "some", not a failed read', async () => {
  const answer = await readSourced({
    tool: 't',
    mirror: reader().reader,
    sync: { refresh: () => Promise.resolve('current' as const) },
    fromMirror: () => Promise.resolve('from mirror'),
    fromGraph: () => Promise.reject(new Error('graph must not be called')),
    inactiveCoverage: () => Promise.reject(new Error('firestore down')),
  });

  assert.equal(answer.source, 'best-available');
  assert.equal(answer.data, 'from mirror');
});
```

- [ ] **Step 2: Run the tests to verify they fail**

```bash
npm test 2>&1 | tail -40
```

- [ ] **Step 3: Widen the source type**

In `src/mirror-reader.ts`, replace the `MirrorSource` declaration and its comment:

```ts
/**
 * What a tool result claims about the answer it carries.
 *
 * `onenote` is a claim that the answer equals what OneNote holds: either it came from
 * OneNote directly, or it came from the mirror after a refresh that finished, or it is a
 * page document no write has marked stale.
 *
 * `mirror` is the weaker claim — a stored copy that may be behind, with `mirroredAt`
 * saying how far. It is what a refresh that did not finish leaves behind: the budget ran
 * out, the scheduler held the lease, Graph refused, Firestore was unreachable.
 *
 * `best-available` sits between them and is not a degraded answer. Everything in it is
 * either confirmed current or came from a notebook the operator has marked inactive —
 * that is, has asserted is no longer being edited. Without this value an unscoped
 * `search_pages` would report `mirror` from the moment one notebook went inactive, for
 * ever, which would train a calling model to ignore the label that matters.
 *
 * Deliberately **not** the same axis as which store answered. A model has no use for
 * "Firestore or Graph", and reporting the store would make the common case — a mirror
 * read the inline sync just brought current — read as second-hand data.
 */
export type MirrorSource = 'onenote' | 'best-available' | 'mirror';

/**
 * How much of one answer came from notebooks the operator marked inactive.
 *
 * `'some'` is also the value for "cannot tell", because it is the pessimistic one: it
 * can never produce `onenote`, and it cannot claim `best-available` on an answer whose
 * refresh did not finish.
 */
export type InactiveCoverage = 'none' | 'some' | 'all';
```

- [ ] **Step 4: Add the decision function**

Immediately above `readSourced` in the same file:

```ts
/**
 * The label, from whether the answer was confirmed and how much of it is inactive.
 *
 * Exported so the six rows are asserted directly rather than through a fake read.
 *
 * The `'all'` rows are the ones that look wrong and are not. A read confined to an
 * inactive notebook reports `best-available` even when the refresh failed outright,
 * because that refresh was never going to check that notebook — its failure says nothing
 * about this answer. The claim rests on two things that hold regardless: the operator's
 * assertion that the notebook is not edited, and the fact that a write through this
 * server marks its page stale or holds its section listing, and a held listing is a
 * mirror miss that goes to Graph.
 */
export function sourceFor(confirmed: boolean, coverage: InactiveCoverage): MirrorSource {
  if (coverage === 'all') return 'best-available';
  if (!confirmed) return 'mirror';
  return coverage === 'some' ? 'best-available' : 'onenote';
}
```

- [ ] **Step 5: Teach the reader about activity**

Add to the imports from `'./mirror-schema.ts'`: `isActive`, and `type NotebookSelection`.
Add to the imports from `'./read-sync.ts'`: `INLINE_SYNC_MIN_INTERVAL_MS`.

Add `getSelection` to `MirrorReadStore`:

```ts
export interface MirrorReadStore {
  getSelection(): Promise<NotebookSelection>;
  listNotebooks(): Promise<MirrorNotebook[]>;
  // ... the rest unchanged
}
```

Add the memo constant near the other module constants:

```ts
/**
 * How long the activity snapshot is reused.
 *
 * Matched to the inline refresh's own interval. The whole consequence of the memo is a
 * `source` label up to 30 seconds behind an operator's edit to the selection document;
 * without it every covered tool call would add a Firestore document read and a notebook
 * query.
 */
export const ACTIVITY_MEMO_MS = INLINE_SYNC_MIN_INTERVAL_MS;
```

Add to `MirrorReader`, after the existing private fields:

```ts
  readonly #now: () => number;
  #activity: { at: number; value: ActivitySnapshot } | null = null;
```

Change the constructor to accept the clock, keeping the existing parameters and order:

```ts
  constructor(
    store: MirrorReadStore,
    blobs: MirrorBlobReader,
    maxInkBytes?: number,
    now: () => number = Date.now,
  ) {
    this.#store = store;
    this.#blobs = blobs;
    this.#maxInkBytes = maxInkBytes;
    this.#now = now;
  }
```

Add the snapshot type near the other interfaces in the file:

```ts
/** The selection and the mirrored notebooks, read together and reused for a short while. */
interface ActivitySnapshot {
  readonly selection: NotebookSelection;
  readonly mirrored: number;
  readonly inactive: number;
}
```

Add the methods:

```ts
  /**
   * The selection document and the mirrored notebook count, memoised.
   *
   * Read together because every caller needs both, and reused because they answer a
   * question about configuration rather than about data — an operator edits the selection
   * by hand, in a console, not thirty times a minute.
   */
  async #activitySnapshot(): Promise<ActivitySnapshot> {
    const now = this.#now();
    const cached = this.#activity;
    if (cached !== null && now - cached.at < ACTIVITY_MEMO_MS) return cached.value;

    const [selection, notebooks] = await Promise.all([
      this.#store.getSelection(),
      this.#store.listNotebooks(),
    ]);

    // Only mirrored notebooks are counted. `mirrored` and `active` are different
    // questions: a notebook whose pages were never mirrored contributes nothing to any
    // answer, so it cannot make one best-available.
    const held = notebooks.filter((notebook) => notebook.mirrored);
    const value: ActivitySnapshot = {
      selection,
      mirrored: held.length,
      inactive: held.filter((notebook) => !isActive(selection, notebook.id)).length,
    };

    this.#activity = { at: now, value };
    return value;
  }

  /** How much of an account-wide answer comes from inactive notebooks, and how many. */
  async accountActivity(): Promise<{ coverage: InactiveCoverage; inactiveNotebooks: number }> {
    const { mirrored, inactive } = await this.#activitySnapshot();

    const coverage: InactiveCoverage =
      inactive === 0 ? 'none' : inactive === mirrored ? 'all' : 'some';

    return { coverage, inactiveNotebooks: inactive };
  }

  /**
   * Coverage for an answer confined to one section.
   *
   * A section the mirror cannot find is `'some'` rather than `'all'`: `'all'` claims
   * every unconfirmed part of the answer came from a notebook nobody edits, and a missing
   * document says nothing of the kind.
   */
  async coverageOfSection(sectionId: string): Promise<InactiveCoverage> {
    const section = await this.#store.getSection(sectionId);
    if (section === null) return 'some';
    return this.#coverageOfNotebook(section.notebookId);
  }

  /** Coverage for an answer confined to one page. */
  async coverageOfPage(pageId: string): Promise<InactiveCoverage> {
    const page = await this.#store.getPage(pageId);
    if (page === null) return 'some';
    return this.#coverageOfNotebook(page.notebookId);
  }

  async #coverageOfNotebook(notebookId: string): Promise<InactiveCoverage> {
    const { selection } = await this.#activitySnapshot();
    return isActive(selection, notebookId) ? 'none' : 'all';
  }
```

- [ ] **Step 6: Use it in `readSourced`**

Add the member to `SourcedRead`:

```ts
  /**
   * How much of this answer comes from notebooks the operator marked inactive.
   *
   * `data` is the mirror's answer, already in hand: the two `_by_name` reading tools
   * resolve a section id inside `fromMirror`, and that id is the only thing that says
   * which notebook the answer came from. Every other tool ignores it and uses the id its
   * own handler already has.
   *
   * Absent means `'none'`, which is what keeps `list_notebooks` and `list_sections` out
   * of this entirely — they answer from structure, and structure is read for the whole
   * account in one request whether or not a notebook is active.
   */
  inactiveCoverage?(reader: MirrorReader, data: M): Promise<InactiveCoverage>;
```

Replace the final return of `readSourced` with:

```ts
  const stamp =
    read.mirroredAt === undefined ? null : await read.mirroredAt(mirror).catch(() => null);

  // A failure here must not fail the read and must not strengthen the claim. `'some'` is
  // the pessimistic value: it cannot produce `onenote`, and it cannot report a refresh
  // that did not finish as anything but `mirror`.
  const coverage: InactiveCoverage =
    read.inactiveCoverage === undefined
      ? 'none'
      : await read.inactiveCoverage(mirror, hit).catch(() => 'some' as const);

  return {
    data: hit,
    origin: 'mirror',
    source: sourceFor(freshness === 'current' || read.staleTracked === true, coverage),
    ...(stamp === null ? {} : { mirroredAt: stamp }),
  };
```

Update the doc comment on `readSourced`, step 3, to name the third value.

- [ ] **Step 7: Run the gates**

```bash
npm run typecheck && npm test 2>&1 | tail -30
```

Expected: `test/mirror-tools.test.ts` fails to typecheck because its `MirrorReadStore` fake has no `getSelection`. Add it there too:

```ts
    getSelection: () => Promise.resolve({ notebookIds: ['nb-1'] }),
```

- [ ] **Step 8: Commit**

```bash
git add src/mirror-reader.ts test/mirror-reader.test.ts test/mirror-tools.test.ts
git commit -F - <<'EOF'
A third source value for an answer that skipped an inactive notebook

`best-available` means everything returned is either confirmed current or came
from a notebook the operator has marked as no longer edited. Without it an
unscoped search_pages would report `mirror` from the moment one notebook went
inactive, for ever, conflating a deliberate skip with a refresh that failed and
training a model to ignore the label that matters.

Coverage `all` ignores whether the refresh finished. A refresh that failed was
never going to check an inactive notebook, so its failure says nothing about an
answer confined to one. The claim rests instead on the operator's assertion and
on the write path, which marks a page stale or holds a section listing before it
touches OneNote -- and a held listing is a mirror miss.

A missing section or page document is `some`, not `all`. `all` claims every
unconfirmed part of the answer came from a notebook nobody edits, and a document
the mirror cannot find says nothing of the kind. A throw from `inactiveCoverage`
lands in the same place rather than failing the read.

The activity snapshot is memoised for the inline refresh's own interval. The
whole consequence is a label up to 30 seconds behind an edit to the selection
document.

Co-Authored-By: Claude Opus 5 (1M context) <noreply@anthropic.com>
EOF
```

---

### Task 6: Wire coverage into the tools

**Goal:** Each covered tool reports the right coverage, `list_notebooks` says which notebooks are active, unscoped `search_pages` says how many it skipped, and every covered description explains the three source values.

**Files:**
- Modify: `src/mcp-tools.ts`, `src/structure-tools.ts`, `src/page-tools.ts`, `src/mirror-reader.ts`
- Test: `test/mirror-tools.test.ts`

**Acceptance Criteria:**
- [ ] `get_page_content`, `list_pages`, `list_pages_by_name`, `find_page_by_name` and scoped `search_pages` report `best-available` for an inactive notebook and `onenote` for an active one.
- [ ] Unscoped `search_pages` reports `onenote`, `best-available` or `mirror` by the mix, and carries `inactiveNotebooks`.
- [ ] `list_notebooks` and `list_sections` pass no coverage and are unchanged apart from `pagesActive`.
- [ ] Mirror-origin `list_notebooks` entries carry `pagesActive`; Graph-origin entries carry neither flag.
- [ ] All eight covered tool descriptions end with `SOURCE_NOTE`.

**Verify:** `npm run typecheck && npm test` → pass.

**Steps:**

- [ ] **Step 1: Write the failing tests**

In `test/mirror-tools.test.ts`, following the file's existing shape for building tools against a fake reader:

```ts
test('a read confined to an inactive notebook is best-available', async () => {
  const tools = toolsFor({
    selection: { notebookIds: ['nb-1'], activeNotebookIds: [] },
    sections: [section({ id: 's-1', notebookId: 'nb-1' })],
    pages: [page({ id: 'p-1', sectionId: 's-1', notebookId: 'nb-1' })],
  });

  const listed = await call(tools, 'list_pages', { sectionId: 's-1' });
  assert.equal(listed.source, 'best-available');

  const content = await call(tools, 'get_page_content', { pageId: 'p-1' });
  assert.equal(content.source, 'best-available');
});

test('a read confined to an active notebook is unchanged', async () => {
  const tools = toolsFor({
    selection: { notebookIds: ['nb-1'], activeNotebookIds: ['nb-1'] },
    sections: [section({ id: 's-1', notebookId: 'nb-1' })],
    pages: [page({ id: 'p-1', sectionId: 's-1', notebookId: 'nb-1' })],
  });

  assert.equal((await call(tools, 'list_pages', { sectionId: 's-1' })).source, 'onenote');
  assert.equal((await call(tools, 'get_page_content', { pageId: 'p-1' })).source, 'onenote');
});

test('an unscoped search reports the mix and how many notebooks it skipped', async () => {
  const tools = toolsFor({
    selection: { notebookIds: ['nb-1', 'nb-cold'], activeNotebookIds: ['nb-1'] },
    notebooks: [
      notebook({ id: 'nb-1', mirrored: true }),
      notebook({ id: 'nb-cold', mirrored: true }),
    ],
  });

  const found = await call(tools, 'search_pages', { query: 'x' });

  assert.equal(found.source, 'best-available');
  assert.equal(found.inactiveNotebooks, 1);
});

test('list_notebooks says which notebooks are still checked', async () => {
  const tools = toolsFor({
    selection: { notebookIds: ['nb-1', 'nb-cold'], activeNotebookIds: ['nb-1'] },
    notebooks: [
      notebook({ id: 'nb-1', mirrored: true }),
      notebook({ id: 'nb-cold', mirrored: true }),
    ],
  });

  const listed = await call(tools, 'list_notebooks', {});

  assert.equal(listed.source, 'onenote');
  assert.deepEqual(
    listed.notebooks.map((n: { id: string; pagesActive: boolean }) => [n.id, n.pagesActive]),
    [['nb-1', true], ['nb-cold', false]],
  );
});
```

> `toolsFor`, `call`, `notebook`, `section` and `page` are this file's existing helpers; use the names it already has. `toolsFor` needs to pass the fixture's `selection` through to the fake store.

- [ ] **Step 2: Run the tests to verify they fail**

```bash
npm test 2>&1 | tail -40
```

- [ ] **Step 3: Add the shared description note**

In `src/mcp-tools.ts`, beside the other exported helpers:

```ts
/**
 * What `source` means, appended to every tool description that reports one.
 *
 * One constant rather than eight copies: the values are a contract with the calling
 * model, and eight copies of a contract drift. Written for a model rather than for a
 * person, like every other description in this repository.
 */
export const SOURCE_NOTE =
  ' Every result carries a source field. "onenote" means the data is confirmed current ' +
  'with OneNote. "best-available" means part of it came from a local copy of a notebook ' +
  'the operator has marked as no longer edited — treat it as current. "mirror" means ' +
  'part of it came from a local copy that could not be confirmed on this call and may be ' +
  'out of date; mirroredAt says how old that copy is.';
```

- [ ] **Step 4: Append it to the eight descriptions**

In `src/structure-tools.ts`, import `SOURCE_NOTE` from `'./mcp-tools.ts'` and append it to the `description` of `list_notebooks`, `list_sections`, `list_pages`, `search_pages`, `find_page_by_name` and `list_pages_by_name` — each is a concatenated string literal, so the change is `+ SOURCE_NOTE` at the end of each.

In `src/page-tools.ts`, do the same for `get_page_content`.

- [ ] **Step 5: Report `pagesActive`**

In `src/mirror-reader.ts`, extend `MirroredNotebook`:

```ts
/** One notebook as `list_notebooks` reports it from the mirror. */
export interface MirroredNotebook extends Notebook {
  /** Is this notebook's page content mirrored, or only its name? */
  readonly pagesMirrored: boolean;
  /**
   * Is it also still being checked?
   *
   * False for a notebook the operator has marked inactive: its pages are held and served
   * but no sync re-reads them, so a model choosing where to look knows which of its
   * answers will carry `best-available`.
   */
  readonly pagesActive: boolean;
}
```

and change `listNotebooks`:

```ts
  async listNotebooks(): Promise<MirroredNotebook[] | null> {
    const notebooks = await this.#store.listNotebooks();
    if (notebooks.length === 0) return null;

    const { selection } = await this.#activitySnapshot();

    return notebooks.map((notebook) => ({
      id: notebook.id,
      displayName: notebook.displayName,
      pagesMirrored: notebook.mirrored,
      pagesActive: notebook.mirrored && isActive(selection, notebook.id),
    }));
  }
```

In `src/structure-tools.ts`, the `list_notebooks` Graph fallback maps `pagesMirrored: false`; add `pagesActive: false` beside it so the shapes stay assignable. The result mapping already strips both fields on a Graph-origin answer and needs no change.

- [ ] **Step 6: Pass coverage from each tool**

In `src/structure-tools.ts`:

`list_pages` — add to its `readSourced` call:

```ts
          inactiveCoverage: (reader) => reader.coverageOfSection(sectionId),
```

`search_pages` — add:

```ts
          inactiveCoverage: async (reader) =>
            sectionId === undefined
              ? (await reader.accountActivity()).coverage
              : reader.coverageOfSection(sectionId),
```

and in its mirror branch, before the `jsonResult` call:

```ts
        // `best-available` says a skip happened; this says how large it was. The reader
        // memoises the snapshot, so asking a second time costs nothing.
        const inactiveNotebooks =
          sectionId === undefined && mirror !== undefined
            ? (await mirror.accountActivity()).inactiveNotebooks
            : undefined;
```

then spread it into the result object:

```ts
          notebooksInAccount: found.notebooksInAccount,
          ...(inactiveNotebooks === undefined ? {} : { inactiveNotebooks }),
          note: mirrorSearchNote(found, matches.length),
```

`resolveAndList` — add to its `readSourced` call, which covers both `find_page_by_name` and `list_pages_by_name`:

```ts
    inactiveCoverage: (reader, data) => reader.coverageOfSection(data.resolved.section.id),
```

`list_notebooks` and `list_sections` — pass nothing. Structure is read for the whole account
in one request whether or not a notebook is active, so those two answers are as current as
they ever were.

In `src/page-tools.ts`, add to the `get_page_content` `readSourced` call:

```ts
          inactiveCoverage: (reader) => reader.coverageOfPage(pageId),
```

- [ ] **Step 7: Run the gates**

```bash
npm run typecheck && npm test 2>&1 | tail -30
```

Expected: `test/structure-tools.test.ts` and `test/page-tools.test.ts` still pass untouched — they build tools with no mirror argument, so no coverage is ever asked for.

- [ ] **Step 8: Commit**

```bash
git add src/mcp-tools.ts src/structure-tools.ts src/page-tools.ts src/mirror-reader.ts test/mirror-tools.test.ts
git commit -F - <<'EOF'
Report best-available from the tools that can answer out of an archive

Six tools pass coverage; list_notebooks and list_sections pass none. Structure
is read for the whole account in one Graph request whether or not a notebook is
active, so those two answers are exactly as current as they were before this
feature existed and marking them would be a false weakening.

An unscoped search_pages also reports `inactiveNotebooks`. The label says a skip
happened and the count says how large it was, which is the same idiom as
`stoppedEarly` and `notebooksSearched`: a bounded result that omits its bound
reads as a complete one.

`list_notebooks` gains `pagesActive` beside `pagesMirrored`, so a model choosing
where to look can tell which of its answers will come back best-available.

`SOURCE_NOTE` is one exported constant appended to all eight descriptions. The
three values are a contract with the calling model, and eight copies of a
contract drift.

Co-Authored-By: Claude Opus 5 (1M context) <noreply@anthropic.com>
EOF
```

---

### Task 7: Documentation

**Goal:** `CLAUDE.md`, `README.md` and `project-spec.md` describe the feature, so the next agent does not undo it.

**Files:**
- Modify: `CLAUDE.md`, `README.md`, `project-spec.md`

**Acceptance Criteria:**
- [ ] `CLAUDE.md` carries four new convention blocks, in the style of the ones around them.
- [ ] `README.md` documents the fourth route, the new selection field, and where the four jobs sit in the schedule.
- [ ] `project-spec.md`'s mirror section describes activity.
- [ ] No document names a real notebook, a tenant, or a secret.

**Verify:** `npm run typecheck && npm test` → pass (documentation only; the gates confirm nothing was broken). Read each edited section back once.

**Steps:**

- [ ] **Step 1: Add the `CLAUDE.md` conventions**

In the `## Conventions` section, near the existing mirror blocks, add:

```markdown
**A notebook can be mirrored without being checked, and `activeNotebookIds` is the
switch.** The root selection document's second, optional list names the mirrored notebooks
the syncs still re-read. Absent or null means all of them, so a deployment that has never
set the field behaves exactly as before; an explicitly empty array means none, which is why
`NotebookSelection.activeNotebookIds` is nullable rather than defaulted to `notebookIds`. A
malformed value reads as "everything is active": failing open costs Graph requests on
notebooks that did not need checking, and failing closed stops the mirror updating with
nothing in any tool result to say so. An inactive notebook is still backfilled — a section
with `pagesSyncedThrough === null` is eligible whatever its notebook's state — and after
that no incremental and neither scheduled sweep lists it again.

**Activity is never stored on a notebook, section or page document, and never enters
`structureHashOf`.** `MirrorStore.putStructure` replaces documents wholesale — `batch.set`
with no merge — and `buildStructure` emits `pagesSyncedThrough: null` and `pageCount: 0`. So
anything inside the structure hash resets every section's watermark when it moves, and an
activation edit would trigger a full re-backfill of the whole selection, hours of the request
budget. The active set is read from the selection document at the start of each run instead,
and `splitByActivity` filters the section list in memory.

**Changing the active set nulls `sectionsScannedThrough`, and that is not belt-and-braces.**
Tier 1 skips a section unless its `graphLastModifiedDateTime` is newer than
`overlapFrom(sectionsScannedThrough)`, and that watermark advances on every completed run. A
section edited while its notebook was inactive is older than the cutoff and stays older for
ever, so re-activating the notebook would otherwise change nothing at all.
`applyActivationChange` compares `activeSelectionHashOf(selection)` against the stored hash,
patches the null, and **returns the corrected state** — `ctx.state` is a snapshot taken when
the run started, so patching Firestore alone would defer the effect to the next run. Both
passes call it, because either may be the first to run after an edit. Per-section watermarks
are untouched, so each section is re-listed from where it left off.

**`source` has three values and `best-available` is not a degraded answer.** `onenote` means
the data is confirmed current. `mirror` means part of it came from a copy that could not be
confirmed on this call — the refresh failed, ran out of budget, or found the lease held — and
may be behind. `best-available` means everything unconfirmed came from a notebook the
operator marked inactive. `sourceFor` in `src/mirror-reader.ts` is the one place the six rows
are written, and the two `'all'` rows ignore whether the refresh finished: a refresh that
failed was never going to check an inactive notebook, so its failure says nothing about an
answer confined to one. Without the third value an unscoped `search_pages` would report
`mirror` from the moment one notebook went inactive, for ever, which teaches a model to
ignore the label that matters. A missing section or page document is `'some'`, never
`'all'` — `'all'` claims every unconfirmed part came from a notebook nobody edits, and a
document the mirror cannot find says nothing of the kind.

**`list_notebooks` and `list_sections` pass no coverage, and that is correct rather than an
omission.** Both answer from structure, and `getExpandedTree()` reads the whole account in
one request whether or not a notebook is active — `buildStructure` stores every notebook and
section it returns, which is what stops `list_notebooks` answering partially. Marking those
two `best-available` would be a false weakening of a claim that still holds.

**`POST /sync/sweep/all` is the only sweep that visits an inactive notebook, and
`/sync/sweep/full` deliberately still does not.** `full` is the weekly scheduled backstop, so
making it ignore activity would give back every request the feature saves. `all` is the
manual lever for re-checking an archive without editing the selection document. The mode is
the path for the reason the other three are: `src/logging.ts` records the path and records
neither the query string nor the body.
```

- [ ] **Step 2: Update `README.md`**

Find the section documenting the selection document and the sync schedule. Add
`activeNotebookIds` to the example document, add the fourth route to the schedule table with
"manual, not scheduled" against it, and state the operational rule: a notebook is moved out
of `activeNotebookIds` when it stops being edited, and back in when it starts again — and
moving it back costs one run that re-lists every eligible section.

- [ ] **Step 3: Update `project-spec.md`**

In the mirror section, describe the mirrored-versus-active distinction, the one-off backfill,
and the three `source` values.

- [ ] **Step 4: Run the gates and read the edits back**

```bash
npm run typecheck && npm test 2>&1 | tail -10
git diff --stat
```

- [ ] **Step 5: Commit**

```bash
git add CLAUDE.md README.md project-spec.md
git commit -F - <<'EOF'
Document active notebooks

Four conventions, each recording a decision whose obvious-looking alternative is
wrong: activity stays out of the structure hash because putStructure replaces
documents wholesale; a changed active set nulls sectionsScannedThrough because
tier 1 would otherwise skip a re-activated notebook for ever; best-available
exists so a deliberate skip is not reported the same way as a failed refresh;
and list_notebooks and list_sections pass no coverage because structure is read
for the whole account regardless.

Co-Authored-By: Claude Opus 5 (1M context) <noreply@anthropic.com>
EOF
```

---

## Manual verification, after Task 7

None of this is covered by an automated test, for the reasons `CLAUDE.md` gives about
`src/mirror-store.ts`: there is no Firestore emulator on this machine and an in-memory fake
would assert the fake.

- [ ] Add `activeNotebookIds` to the real selection document naming one notebook, leaving at least one other selected notebook out of it.
- [ ] `POST /sync` and read the response: `sectionsSkippedInactive` should be non-zero once the inactive notebook's backfill has finished, and zero while it is still filling.
- [ ] `POST /sync/sweep` — `sectionsSkippedInactive` non-zero, `sectionsVisited` covering only the active notebook.
- [ ] `POST /sync/sweep/all` — `sectionsVisited` covering both.
- [ ] Call `list_pages` on a section in the inactive notebook: `source` is `best-available`.
- [ ] Call `list_pages` on a section in the active notebook: `source` is `onenote`.
- [ ] Call `search_pages` with no `sectionId`: `source` is `best-available` and `inactiveNotebooks` matches the count you left out.
- [ ] `append_to_page` against a page in the inactive notebook, then `get_page_content` on it: the appended text is there.
- [ ] Move the notebook back into `activeNotebookIds`, `POST /sync`, and confirm `sectionsVisited` covers its sections again.

## Not covered by any test in this plan

- Whether Firestore accepts a selection document carrying two arrays. It is two string arrays
  in a document that already holds one; the risk is nil, and only a live write settles it.
- Whether a calling model reads `best-available` as intended. Nothing confirms that until a
  real client runs against the deployed URL.
- Whether 30 seconds is the right memo window for the activity snapshot. It is matched to
  `INLINE_SYNC_MIN_INTERVAL_MS` by argument, not measured.
