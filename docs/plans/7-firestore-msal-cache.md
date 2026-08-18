# Firestore-backed MSAL cache plugin with transactional read-modify-write

**Issue:** #7
**Goal:** Implement MSAL's `ICachePlugin` against one Firestore document so the rotating refresh token survives Cloud Run's ephemeral filesystem, with the write half wrapped in a transaction that cannot lose an overlapping instance's update.
**Date:** 2026-08-18

## Status

| Phase | Status | Commit |
|-------|--------|--------|
| Phase 1: `src/token-cache.ts` — plugin, factory, pure-function unit tests | ☑ | this commit |
| Phase 2: Documentation — `README.md`, `CLAUDE.md` | ☑ | this commit |

## Acceptance Criteria

> Reproduced from issue #7 so the plan is self-contained. The six "Tasks" checkboxes and both sentences of "Acceptance" are treated as acceptance criteria.

- **AC-1:** Implement MSAL's `ICachePlugin` (`beforeCacheAccess` / `afterCacheAccess`) against a single Firestore document, path from `FIRESTORE_CACHE_DOC` (default `tokencache/msal`).
- **AC-2:** `afterCacheAccess` writes only when `tokenCacheContext.cacheHasChanged` is true.
- **AC-3:** Wrap the read-modify-write in a Firestore transaction.
- **AC-4:** Store the serialized cache as a single string field plus an `updatedAt` timestamp. Do not model MSAL's internal structure.
- **AC-5:** Handle "document does not exist" as an empty cache, not an error — that is the pre-bootstrap state.
- **AC-6:** Unit tests against the Firestore emulator, including a concurrent-write test that proves the transaction serialises. ⚠️ Deviation: **no emulator-dependent test is committed.** The emulator is not installed here and cannot be installed by the agent — `gcloud` is the Debian package build with its component manager disabled, so `gcloud components install cloud-firestore-emulator` is refused and directs the operator to `sudo apt-get install google-cloud-cli-firestore-emulator`. Committing tests that cannot run in this environment, and that no one has watched pass, buys nothing. Instead: the decoder is unit-tested (Phase 1), and the Firestore-touching behaviour is verified by the operator running the manual procedure documented below. No substitute Firestore is built either.
- **AC-7:** A test writes a cache, reads it back through a second plugin instance, and gets byte-identical content. ⚠️ Deviation: same as AC-6 — verified by step 2 of the manual procedure, not by a committed test.
- **AC-8:** A concurrent-write test produces one coherent final document rather than a partial merge. ⚠️ Deviation: same as AC-6 — verified by steps 3 and 4 of the manual procedure, not by a committed test.

## Context

### What exists today

Issues #5 and #6 are complete. `src/` holds `index.ts`, `server.ts`, `config.ts`, `bootstrap.ts`, `version.ts`. `test/` holds `config.test.ts`, `server.test.ts`, `version.test.ts`. There is no Firestore code anywhere in `src/` — `@google-cloud/firestore@9.0.0` is a declared dependency that nothing imports yet.

`src/config.ts` already carries the two variables this issue needs, in the `firestore` group:

- `FIRESTORE_CACHE_DOC` (`src/config.ts:50`), optional, fallback `'tokencache/msal'`, validated by `checkDocumentPath` — rejects empty segments and an odd segment count, so a collection path cannot reach the plugin.
- `GOOGLE_CLOUD_PROJECT` (`src/config.ts:58`), optional, undefined on Cloud Run where the client infers the project.

`loadConfig(['firestore'])` returns `FirestoreConfig { cacheDocumentPath: string; projectId: string | undefined }` (`src/config.ts:96`). No config change is needed for this issue.

`src/bootstrap.ts` is still the placeholder from #5 and already requests `['graph', 'firestore']`. It is the future writer of the cache (issue #9); the deployed server is the future reader and refresher (issue #12). Neither is in scope here.

### Design specs consulted

There is no `docs/design/` directory. `project-spec.md` is authoritative.

- `project-spec.md:173` — "Cloud Run containers are ephemeral and stateless, so the token cache cannot live on the container filesystem." The reason this module exists.
- `project-spec.md:186` — "MSAL token cache | Firestore in Native mode, one document". One document, not a collection per entity.
- `project-spec.md:225` "Token cache: Firestore, not Secret Manager" — the cache is rewritten on every refresh because Microsoft issues rotating refresh tokens; `doc.set()` overwrites in place and Firestore has transactions, Secret Manager has neither.
- `project-spec.md:259` Constraint 1 — "Every refresh returns a new refresh token and the service must write the updated cache back to Firestore. `--max-instances=1` removes most of the concurrent-write risk; use a Firestore transaction for the read-modify-write anyway, since revision transitions can briefly overlap instances." This is AC-3's justification and the reason the transaction body must re-read rather than blind-write.
- `project-spec.md:264` Constraint 2 — "Bootstrap sign-in happens locally, never on Cloud Run." The plugin must therefore work identically from a laptop and from Cloud Run, which it does: same class, same document, different ambient credentials.
- `project-spec.md:394` and `project-spec.md:402` Repo hygiene — the MSAL token cache is secret and must never appear in source, an issue, or a workflow log. Every cache blob in a test fixture is fabricated: fake GUIDs, a fake tenant, no real tokens.

### Issue Assessment

Findings from checking the issue against the current repository and the installed dependency tree. Each claim below was checked against a real file or a real command.

**Nothing stale.** Every AC refers to something that exists. `FIRESTORE_CACHE_DOC` with the `tokencache/msal` fallback is live at `src/config.ts:50`. `@azure/msal-node@5.5.0` and `@google-cloud/firestore@9.0.0` are both installed.

**Verified: the interface to implement.** `node_modules/@azure/msal-common/types/cache/interface/ICachePlugin.d.ts` declares exactly two members, both `(tokenCacheContext: TokenCacheContext) => Promise<void>`. `TokenCacheContext` exposes `cacheHasChanged: boolean` and `tokenCache: ISerializableTokenCache` (`serialize(): string` / `deserialize(cache: string): void`). `@azure/msal-node` re-exports `ICachePlugin`, `TokenCacheContext`, and `ISerializableTokenCache` from its index (`node_modules/@azure/msal-node/types/index.d.ts:53`), so `src/` and `test/` import all three from `'@azure/msal-node'` and never reach into `@azure/msal-common` directly.

**Verified: MSAL's `deserialize` merges, it does not replace.** `TokenCache.deserialize` in `node_modules/@azure/msal-node/types/cache/TokenCache.d.ts` documents `mergeState` / `mergeUpdates` / `mergeRemovals` as its implementation. That is what makes the transaction body's recovery path correct: on detecting a competing write, the plugin feeds the stored blob back through `deserialize` and re-serialises, and the result contains both sides rather than one clobbering the other.

**Verified: the emulator is honoured by the client, but is not installed.** `node_modules/@google-cloud/firestore/build/src/index.js:604` reads `FIRESTORE_EMULATOR_HOST` and rewrites the endpoint, so no special client construction is needed for tests. `gcloud emulators firestore start --help` succeeds (the command exists), but `gcloud components list --filter="id:cloud-firestore-emulator"` reports `Not Installed` and `gcloud components install` is refused because this is the Debian-packaged CLI. `java` is on `PATH`, so the emulator will run once `sudo apt-get install google-cloud-cli-firestore-emulator` has been done.

**Scope gap 1 — the issue does not say what happens when the document exists but is malformed.** AC-5 covers only the absent document. A document whose `cache` field is missing or is not a string is a different situation: treating it as empty would make MSAL report no accounts, and the next `afterCacheAccess` would then overwrite whatever was actually there. The plugin therefore throws a `TokenCacheError` naming the document path and the type it found. A hard failure at startup is recoverable by re-running bootstrap; a silent overwrite of a real cache is not.

**Scope gap 2 — the transaction has to survive the split between the two callbacks.** MSAL calls `beforeCacheAccess`, then does its work, then calls `afterCacheAccess`. A Firestore transaction cannot be held open across that gap; `runTransaction` owns its own callback scope. So "wrap the read-modify-write in a transaction" (AC-3) is implemented as: `beforeCacheAccess` records the exact string it handed to MSAL, and `afterCacheAccess` opens a transaction that re-reads the document, compares it to that recorded string, and folds in any competing write before serialising. The transaction gives the compare-and-set; the recorded string is what there is to compare against.

**Scope excess — none.** No AC duplicates existing functionality or contradicts a project convention.

**No substitute Firestore, and no unrunnable tests.** An earlier draft of this plan proposed an in-memory `FirestoreLike` fake so the plugin's behaviour could be asserted without the emulator, which in turn required the plugin to accept a narrow structural interface instead of a `Firestore`. A later draft dropped the fake but still committed an emulator-gated test file that self-skipped when `FIRESTORE_EMULATOR_HOST` was unset. Both are cut. The plugin takes a real `Firestore`, and the repository commits no test that needs a backend. What is automated is the pure decoder; what needs Firestore is a written manual procedure the operator runs once, with its output recorded in Findings.

**Deliberately out of scope.** Nothing wires the plugin into a running MSAL client in this issue. The bootstrap CLI that first populates the document is issue #9; the Graph client that refreshes through it is issue #12. `src/index.ts`, `src/server.ts`, and `src/bootstrap.ts` are not modified.

### Design decisions

**Document shape.** Exactly two fields, matching AC-4:

| Field | Type | Meaning |
|---|---|---|
| `cache` | string | The blob from `ISerializableTokenCache.serialize()`, stored verbatim |
| `updatedAt` | Firestore `Timestamp` | Written as `FieldValue.serverTimestamp()`; diagnostic only, never read by the plugin |

Nothing else. No parsing of the blob, no per-entity documents, no schema version field — MSAL owns that blob's structure and changing its shape between library versions is MSAL's business, not this module's.

**`tx.set` replaces, it does not merge.** The call is `tx.set(ref, {cache, updatedAt})` with no `{merge: true}`. The merge that matters happens inside MSAL's `deserialize`, before serialising. A Firestore-level field merge on a single opaque string field would achieve nothing.

**The recorded-read string is per-instance, not global.** Each plugin instance holds the last blob it read or wrote. Two instances in the same process — which is what the concurrency test builds — therefore behave like two Cloud Run instances: neither knows what the other read.

## Phase 1: `src/token-cache.ts` — plugin, factory, pure-function unit tests

**Goal:** A `FirestoreTokenCachePlugin` implementing `ICachePlugin`, a factory that builds it from `FirestoreConfig`, and unit coverage of the snapshot decoder, which is the only part that runs without a backend.
**Addresses:** AC-1, AC-2, AC-3, AC-4, AC-5

**Files:**
- Create: `src/token-cache.ts`
- Create: `test/token-cache.test.ts`

**Steps:**

1. In `src/token-cache.ts`, import `Firestore`, `FieldValue` as values and `type DocumentReference` from `@google-cloud/firestore`; import `type ICachePlugin`, `type TokenCacheContext` from `@azure/msal-node`; import `type { FirestoreConfig }` from `./config.ts`.

2. Add `export class TokenCacheError extends Error`, setting `this.name = 'TokenCacheError'`, following the shape of `ConfigError` in `src/config.ts:13`. It carries the document path in the message.

3. Add the field-name constants and the decoder:

   ```ts
   const CACHE_FIELD = 'cache';
   const UPDATED_AT_FIELD = 'updatedAt';

   /** The slice of a Firestore DocumentSnapshot this module reads. Declared narrowly so
    *  the decoder is callable from a test with a plain object. */
   export interface CacheSnapshot {
     readonly exists: boolean;
     data(): Record<string, unknown> | undefined;
   }

   export function readCache(snapshot: CacheSnapshot, documentPath: string): string | null
   ```

   `readCache` returns `null` when `snapshot.exists` is false or `snapshot.data()` is `undefined` (AC-5 — the pre-bootstrap state, not an error); returns the value when `data()[CACHE_FIELD]` is a string; throws `TokenCacheError` otherwise, with the message `Token cache document <path> has no usable "cache" field (found <typeof>). Re-run the bootstrap CLI to recreate it.` A document that exists with `cache: ''` is a valid empty cache and returns `''`, not `null`.

   `DocumentSnapshot` satisfies `CacheSnapshot` structurally, so no cast is needed at the call sites in step 5.

4. Implement the class. It is constructed with a `Firestore` and a document path, resolves the `DocumentReference` once in the constructor, and keeps one piece of mutable state:

   ```ts
   export class FirestoreTokenCachePlugin implements ICachePlugin {
     readonly #firestore: Firestore;
     readonly #documentPath: string;
     readonly #ref: DocumentReference;
     /** The blob last handed to, or written on behalf of, MSAL. null means "document absent". */
     #lastKnown: string | null = null;
     ...
   }
   ```

   Use `#private` fields, not `private` modifiers or constructor parameter properties — the source must stay erasable (`erasableSyntaxOnly`).

5. `beforeCacheAccess(context: TokenCacheContext): Promise<void>`:
   - `const serialized = readCache(await this.#ref.get(), this.#documentPath)`
   - `this.#lastKnown = serialized`
   - if `serialized !== null && serialized !== ''`, call `context.tokenCache.deserialize(serialized)`; otherwise leave MSAL's in-memory cache untouched, which is the empty state.

   Note in a comment that this read is deliberately outside a transaction: it is a plain read, and the compare-and-set that protects the write happens in `afterCacheAccess`.

6. `afterCacheAccess(context: TokenCacheContext): Promise<void>`:
   - Return immediately when `context.cacheHasChanged` is false (AC-2). No read, no write, no transaction.
   - Otherwise run the whole read-modify-write inside `this.#firestore.runTransaction` (AC-3):

     ```ts
     let written: string | undefined;
     await this.#firestore.runTransaction(async (tx) => {
       const stored = readCache(await tx.get(this.#ref), this.#documentPath);
       if (stored !== null && stored !== '' && stored !== this.#lastKnown) {
         // Another instance wrote between our read and now. MSAL's deserialize merges
         // into the in-memory cache rather than replacing it, so folding the stored
         // blob back in keeps both sides' entries.
         context.tokenCache.deserialize(stored);
       }
       written = context.tokenCache.serialize();
       tx.set(this.#ref, { [CACHE_FIELD]: written, [UPDATED_AT_FIELD]: FieldValue.serverTimestamp() });
     });
     this.#lastKnown = written ?? this.#lastKnown;
     ```

     `#lastKnown` is assigned after `runTransaction` resolves, never inside the callback: Firestore re-runs the callback on contention, and a value assigned by an aborted attempt would be a blob that was never committed.

   - `serialize()` is called inside the callback on every attempt, so a retry re-reads and re-merges rather than committing a blob computed before the conflict was known.

7. Add the factory:

   ```ts
   export function createFirestoreTokenCachePlugin(config: FirestoreConfig): FirestoreTokenCachePlugin
   ```

   It constructs `new Firestore(config.projectId === undefined ? {} : { projectId: config.projectId })` — passing `projectId: undefined` explicitly is avoided so the client's own inference path (Cloud Run metadata) is used — and returns `new FirestoreTokenCachePlugin(firestore, config.cacheDocumentPath)`.

**Tests added/updated** (`test/token-cache.test.ts`, no Firestore involved):

- `an absent document reads as an empty cache` — AC-5; `readCache({exists: false, data: () => undefined}, 'tokencache/msal')` returns `null`.
- `a document with a cache field returns it verbatim` — AC-1, AC-4; the returned string is `assert.equal`-identical to the input, including whitespace.
- `an empty cache field is an empty cache, not an absent one` — `cache: ''` returns `''` rather than `null`, which is what stops `beforeCacheAccess` treating a deliberately emptied document as pre-bootstrap.
- `a non-string cache field throws TokenCacheError` — scope gap 1; checks `cache: 42`, `cache: null`, and a document with no `cache` field, and asserts the message contains the document path it was given.
- `extra fields on the document are ignored` — AC-4; a document carrying `cache`, `updatedAt`, and an unknown field still decodes.

`beforeCacheAccess`, `afterCacheAccess`, the transaction, and the factory are not unit-tested — they need a backend, and this repository commits no test that needs one. They are covered by the manual verification procedure below.

**Verification:**
- [ ] `npm test` passes
- [ ] `npm run typecheck` passes (this is the lint gate)
- [ ] `grep -n "as any\|@ts-expect-error\|@ts-ignore" src/token-cache.ts` returns nothing
- [ ] `npm run build` emits `dist/token-cache.js`

## Phase 2: Documentation

**Goal:** Record the new module, state plainly that its Firestore-touching half has no automated coverage, and give the operator the exact manual procedure that stands in for it.
**Addresses:** AC-6, AC-7, AC-8

**Files:**
- Modify: `CLAUDE.md`
- Modify: `README.md`

**Steps:**

1. `CLAUDE.md` directory tree: add `src/token-cache.ts` with the comment `# Firestore-backed MSAL ICachePlugin (issue #7)` and `test/token-cache.test.ts`.
2. `CLAUDE.md` prose under the tree: one paragraph stating that `test/token-cache.test.ts` covers only `readCache`, the pure decoder, and that `beforeCacheAccess`, `afterCacheAccess`, the transaction, and `createFirestoreTokenCachePlugin` have no automated test because the Firestore emulator is not installable here. Point at the manual procedure section of this plan by path. Say explicitly that a future agent must not add an in-memory Firestore fake to close that gap — it would assert the fake's behaviour, not Firestore's.
3. `CLAUDE.md` Conventions: add two entries, each stating the mechanism then the consequence.
   - **The token cache is one opaque string in one document.** The plugin stores `serialize()`'s output verbatim in `cache` plus an `updatedAt` server timestamp. Do not parse the blob or split it across documents; MSAL owns its shape and it changes between library versions.
   - **`afterCacheAccess` re-reads inside the transaction.** `--max-instances=1` does not prevent two instances existing during a revision transition. The transaction re-reads the document, and a blob differing from the one this instance last saw is fed back through `deserialize` before serialising, because MSAL's deserialize merges. Removing the re-read turns an overlap into a lost refresh token, which wedges the cache until bootstrap is re-run.
4. `README.md`, a short subsection after the Scripts table titled "Token cache": what `src/token-cache.ts` is, the fact that `npm test` covers only its decoder, and that the rest is checked by hand against an emulator — with the `sudo apt-get install google-cloud-cli-firestore-emulator` prerequisite, why `gcloud components install` does not work on the Debian-packaged CLI, and that `java` must be on `PATH`.
5. `README.md` Configuration section: no table change — both variables are already documented at `README.md:95` and `README.md:96`. Add one sentence stating that `FIRESTORE_CACHE_DOC` is the document the MSAL cache plugin reads and writes, and that its value must be a document path (even segment count), which `loadConfig` enforces.

**Tests added/updated:**
- None. This phase changes only Markdown. The behaviour it describes is covered by Phase 1's unit tests and by the manual procedure below.

**Verification:**
- [ ] `npm run typecheck` passes
- [ ] `npm test` passes
- [ ] Every path named in `CLAUDE.md`'s tree exists
- [ ] The README's stated prerequisite command is the one that actually appears in `gcloud components install`'s refusal message

## Manual verification procedure

This is what stands in for AC-6, AC-7, and AC-8. It is run by the operator, once, on a machine with the emulator installed. Its output goes into Findings.

**Prerequisites:** `sudo apt-get install google-cloud-cli-firestore-emulator` and `java` on `PATH`. `gcloud components install cloud-firestore-emulator` will not work — this is the Debian-packaged CLI and its component manager is disabled.

**Setup.** In one terminal:

```bash
gcloud emulators firestore start --host-port=127.0.0.1:8081
```

In a second terminal, write a throwaway script under the scratchpad — not in the repository, it is not committed:

```bash
export FIRESTORE_EMULATOR_HOST=127.0.0.1:8081
export GOOGLE_CLOUD_PROJECT=onenote-mcp-emulator   # the client needs a project id and
                                                   # must not fall through to the
                                                   # metadata server
```

The script imports `createFirestoreTokenCachePlugin` and `FirestoreTokenCachePlugin` from `src/token-cache.ts`, uses `TokenCacheContext` from `@azure/msal-node`, and drives them with a fifteen-line `ISerializableTokenCache` stand-in whose `deserialize` merges one level deep (mirroring MSAL's `mergeState`) so a change can be made without acquiring a real token. Every cache blob is fabricated: fake account ids, the fake tenant `00000000-0000-0000-0000-000000000000`, no real tokens.

**Step 1 — absent document reads as empty (AC-5 against a real backend).** Point a plugin at a document path that does not exist and call `beforeCacheAccess`. It must resolve without throwing and leave the in-memory cache empty. This is the pre-bootstrap state; a `get()` on a missing document resolves with `exists === false` rather than rejecting, which is the assumption Phase 1's decoder is built on.

**Step 2 — byte-identical round-trip (AC-7).** Instance A writes a fabricated blob through `afterCacheAccess` with `cacheHasChanged` true. A separately constructed instance B, pointed at the same path, calls `beforeCacheAccess`. Assert the string B received is `===` to the string A wrote, and that `Buffer.byteLength` matches on both. Then read the raw document through the client and assert `Object.keys(data()).sort()` is `['cache', 'updatedAt']`, that `updatedAt` is a `Timestamp`, and that `toMillis()` is within 60s of now (AC-4).

**Step 3 — two overlapping instances, neither write lost (AC-3, AC-8).** A and B both `beforeCacheAccess` on the same empty document. A adds account `acct-a` and commits. B, which still holds the pre-A view, adds `acct-b` and commits. The final `cache` field must parse as valid JSON in one piece and contain both `acct-a` and `acct-b` under `Account`. If B's write had been a blind overwrite, `acct-a` would be gone — that is the lost refresh token this whole design exists to prevent.

**Step 4 — five concurrent writers, one coherent document (AC-8).** Five plugin instances seeded from the same starting blob, all five `afterCacheAccess` calls started together and awaited with `Promise.all`. The final `cache` field must parse and contain all five account ids. Firestore aborts and retries the losers; the re-read inside the transaction callback is what folds each retry's view forward.

**Step 5 — `cacheHasChanged` false writes nothing (AC-2).** Call `afterCacheAccess` with a context constructed as `new TokenCacheContext(cache, false)` against an absent document. The document must still be absent afterwards.

**Step 6 — malformed document fails loudly (scope gap 1).** Write `{cache: 42}` directly through the client, then call `beforeCacheAccess`. It must reject with `TokenCacheError` and the message must name the document path.

**Step 7 — MSAL drives it end to end (AC-1, AC-6).** Seed the document with a fabricated blob holding one `Account` entry. Construct `new PublicClientApplication({auth: {clientId, authority}, cache: {cachePlugin}})` with a fake client id and `https://login.microsoftonline.com/common`, then call `getTokenCache().getAllAccounts()` and assert it returns that account. Then `removeAccount()` it and assert the document's `cache` field no longer contains the account id and `updatedAt` advanced. This is expected to be offline — authority metadata is resolved on token acquisition, not on cache access. If it turns out to make a network call, record that in Findings; steps 1 through 6 still cover the plugin.

**Cleanup.** Stop the emulator. Its data is in-memory and disappears with the process; nothing to delete. Delete the throwaway script.

**Recording.** Paste the script's output into the Findings section of this plan and tick the manual box in Final Verification. An untouched Findings section means AC-6, AC-7, and AC-8 are unverified, and the plan should not be treated as complete.

## Final Verification

- [ ] `npm test` passes; no test is skipped, because no committed test needs a backend
- [ ] `npm run typecheck` passes
- [ ] `npm run build` succeeds and `dist/token-cache.js` exists
- [ ] `bash scripts/test/run.sh` passes (unchanged by this issue, run to prove it was not broken)
- [ ] `git ls-files test/ | xargs grep -ln "FIRESTORE_EMULATOR_HOST"` returns nothing — no committed test depends on an emulator
- [ ] Manual verification procedure completed by the operator and its output recorded in Findings. Until this box is ticked, AC-6, AC-7, and AC-8 are unverified.
- [ ] `git grep -n -i "onmicrosoft\|tenant" src/ test/` finds no real tenant name or id
- [ ] Acceptance criteria traceability:
  - AC-1 (`ICachePlugin` against one document, path from `FIRESTORE_CACHE_DOC`): Phase 1 — `FirestoreTokenCachePlugin` implements both callbacks against the path given to the constructor, and `createFirestoreTokenCachePlugin` takes that path from `FirestoreConfig.cacheDocumentPath`, whose default `tokencache/msal` lives in `src/config.ts:50`. Exercised end to end by manual step 7.
  - AC-2 (write only when `cacheHasChanged`): Phase 1 — the early return at the top of `afterCacheAccess`. Verified by manual step 5.
  - AC-3 (transactional read-modify-write): Phase 1 — the `runTransaction` body with its re-read and conditional `deserialize`. Verified by manual step 3.
  - AC-4 (single string field plus `updatedAt`, no modelling of MSAL internals): Phase 1 — the two-field `tx.set` and a decoder that never parses the blob; unit-tested by `extra fields on the document are ignored`. The server timestamp resolving to a real `Timestamp` is verified by manual step 2.
  - AC-5 (absent document is an empty cache): Phase 1 — `readCache` returns `null` for a missing document, unit-tested by `an absent document reads as an empty cache`. Confirmed against a real backend by manual step 1.
  - AC-6 (emulator-backed tests including concurrency): manual procedure, documented in Phase 2. Deviation recorded above — the emulator cannot be installed by the agent, so no emulator-dependent test is committed and no substitute Firestore is built.
  - AC-7 (byte-identical round-trip through a second instance): manual step 2.
  - AC-8 (concurrent writes produce one coherent document): manual steps 3 and 4.

## Findings

> Filled in during implementation: what the work revealed that planning missed, and the pasted output of the manual verification procedure.

### Phase 1

- **`TokenCacheError` carries the path as a field, not only in the message.** The plan said the message carries the document path. The implementation adds `readonly documentPath: string`, mirroring `ConfigError`'s structured `missing`/`invalid` fields in `src/config.ts:14`. The test asserts the path by field rather than by matching prose, so rewording the message cannot silently break it. The message text is unchanged from the plan.
- **`exists === true` is not sufficient to decode.** `readCache` returns `null` when `snapshot.exists` is true but `data()` is `undefined`, not just when `exists` is false. The plan's prose said this; it is recorded here because the branch condition is two checks, not one. Covered by the first test.
- **No typing workaround was needed.** `DocumentSnapshot` satisfies `CacheSnapshot` structurally — `readonly exists: boolean` and `data(): DocumentData | undefined`, where `DocumentData`'s index signature is assignable to `Record<string, unknown>` — so neither `this.#ref.get()` nor `tx.get(this.#ref)` needs a cast, and `tx.set` accepted the two-field computed-key object without a generic annotation. `src/token-cache.ts` contains no `as any`, `@ts-ignore`, or `@ts-expect-error`.
- **Verified:** `npm run typecheck` exit 0; `npm test` 24 pass / 0 fail / 0 skipped; `npm run build` emits `dist/token-cache.js`.
