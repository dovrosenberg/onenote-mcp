# Firestore-backed MSAL cache plugin with transactional read-modify-write

**Issue:** #7
**Goal:** Implement MSAL's `ICachePlugin` against one Firestore document so the rotating refresh token survives Cloud Run's ephemeral filesystem, with the write half wrapped in a transaction that cannot lose an overlapping instance's update.
**Date:** 2026-08-18

## Status

| Phase | Status | Commit |
|-------|--------|--------|
| Phase 1: `src/token-cache.ts` — plugin, factory, pure-function unit tests | ☑ | this commit |
| Phase 2: Emulator-backed tests (round-trip, concurrency) | ☐ | — |
| Phase 3: Documentation — `README.md`, `CLAUDE.md` | ☐ | — |

## Acceptance Criteria

> Reproduced from issue #7 so the plan is self-contained. The six "Tasks" checkboxes and both sentences of "Acceptance" are treated as acceptance criteria.

- **AC-1:** Implement MSAL's `ICachePlugin` (`beforeCacheAccess` / `afterCacheAccess`) against a single Firestore document, path from `FIRESTORE_CACHE_DOC` (default `tokencache/msal`).
- **AC-2:** `afterCacheAccess` writes only when `tokenCacheContext.cacheHasChanged` is true.
- **AC-3:** Wrap the read-modify-write in a Firestore transaction.
- **AC-4:** Store the serialized cache as a single string field plus an `updatedAt` timestamp. Do not model MSAL's internal structure.
- **AC-5:** Handle "document does not exist" as an empty cache, not an error — that is the pre-bootstrap state.
- **AC-6:** Unit tests against the Firestore emulator, including a concurrent-write test that proves the transaction serialises. ⚠️ Deviation: the emulator is not installed in this environment and cannot be installed by the agent. `gcloud` here is the Debian package build with its component manager disabled — `gcloud components install cloud-firestore-emulator` fails with "You cannot perform this action because the Google Cloud CLI component manager is disabled", directing the operator to `sudo apt-get install google-cloud-cli-firestore-emulator`. The tests are written and committed; they skip with a printed reason when `FIRESTORE_EMULATOR_HOST` is unset, and are run manually by the operator against an emulator they start themselves. No substitute Firestore is built.
- **AC-7:** A test writes a cache, reads it back through a second plugin instance, and gets byte-identical content.
- **AC-8:** A concurrent-write test produces one coherent final document rather than a partial merge.

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

**No substitute Firestore.** An earlier draft of this plan proposed an in-memory `FirestoreLike` fake so the plugin's behaviour could be asserted without the emulator, which in turn required the plugin to accept a narrow structural interface instead of a `Firestore`. That is cut. The plugin takes a real `Firestore`. Everything that needs a backend is verified by the emulator tests in Phase 2, run manually. What remains automated in Phase 1 is the one piece that needs no backend at all: the pure function that decodes a document snapshot.

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

`beforeCacheAccess`, `afterCacheAccess`, the transaction, and the factory are not unit-tested here — they need a backend. They are covered by Phase 2's emulator tests, run manually.

**Verification:**
- [ ] `npm test` passes
- [ ] `npm run typecheck` passes (this is the lint gate)
- [ ] `grep -n "as any\|@ts-expect-error\|@ts-ignore" src/token-cache.ts` returns nothing
- [ ] `npm run build` emits `dist/token-cache.js`

## Phase 2: Emulator-backed tests

**Goal:** Exercise the plugin against real Firestore: byte-identical round-trip through a second instance, and concurrent writers producing one coherent document.
**Addresses:** AC-6, AC-7, AC-8

**Files:**
- Create: `test/token-cache.emulator.test.ts`

**Steps:**

1. The emulator is started by hand. There is no wrapper script and no npm script — two terminals, documented in Phase 3:

   ```bash
   # terminal 1 — needs: sudo apt-get install google-cloud-cli-firestore-emulator
   gcloud emulators firestore start --host-port=127.0.0.1:8081

   # terminal 2
   FIRESTORE_EMULATOR_HOST=127.0.0.1:8081 GOOGLE_CLOUD_PROJECT=onenote-mcp-emulator npm test
   ```

   `GOOGLE_CLOUD_PROJECT` is required: the client needs a project id and must not fall through to a metadata-server lookup.

2. Write `test/token-cache.emulator.test.ts`. The `npm test` glob (`test/**/*.test.ts`) picks the file up unconditionally, so it gates itself: read `process.env['FIRESTORE_EMULATOR_HOST']` once at module scope and pass `{skip: 'FIRESTORE_EMULATOR_HOST is not set — see README, Token cache tests'}` to every `test()` when it is absent. `node --test` prints the skip reason, so a run that checks nothing says so instead of reporting a pass.

3. Fixtures live at the top of this one file, not in a shared helpers directory:
   - `sampleCache(accountId: string): string` — a JSON blob shaped like MSAL's serialized cache (`Account`, `IdToken`, `AccessToken`, `RefreshToken`, `AppMetadata` top-level objects) keyed by `accountId`, with fabricated GUIDs and the fake tenant `00000000-0000-0000-0000-000000000000`. No real tenant, no real token, per the repo hygiene rules.
   - `class TestTokenCache implements ISerializableTokenCache` — roughly fifteen lines: `serialize()` returns the current blob, `deserialize(blob)` merges one level deep (for each top-level key, `{...existing[key], ...incoming[key]}`), mirroring MSAL's `mergeState`. This stands in for MSAL's own `TokenCache` so a test can make a change without acquiring a token. The real `TokenCache` is exercised by the `PublicClientApplication` test below.
   - Contexts are `new TokenCacheContext(cache, hasChanged)`, imported from `@azure/msal-node`.

4. Each test uses its own document path — `` `emulator-${randomUUID()}/msal` `` from `node:crypto` — so tests never share state. An `after()` hook deletes the documents the file created, so a long-lived emulator does not accumulate them.

5. Build the plugin through `createFirestoreTokenCachePlugin({cacheDocumentPath, projectId: 'onenote-mcp-emulator'})`, so the factory is covered rather than only the class.

6. Note for the implementer: `new PublicClientApplication({auth: {clientId, authority}, cache: {cachePlugin}})` followed by `getTokenCache().getAllAccounts()` is expected to be offline — authority metadata is resolved on token acquisition, not on cache access. Verify that when writing the test. If it does attempt a network call, drop that one test and record the finding in the Findings section; the remaining tests still cover AC-6, AC-7, and AC-8.

**Tests added/updated** (`test/token-cache.emulator.test.ts`):

- `an absent document reads as an empty cache against real Firestore` — AC-5, AC-6; confirms a `get()` on a missing document resolves with `exists === false` rather than rejecting, which is what Phase 1's decoder assumes.
- `a cache written by one instance is byte-identical when read by a second` — AC-7; instance A writes `sampleCache('acct-1')`; a freshly constructed instance B calls `beforeCacheAccess`; `assert.equal(bCache.state, aBlob)` on the exact string, plus a `Buffer.byteLength` comparison.
- `afterCacheAccess writes nothing when cacheHasChanged is false` — AC-2; read the document's `updatedAt` before and after, assert it is unchanged and the document is still absent when it started absent.
- `the stored document has exactly cache and updatedAt, and updatedAt is a Timestamp` — AC-4; `Object.keys(data()).sort()` deep-equals `['cache', 'updatedAt']`, `data().updatedAt instanceof Timestamp`, and `toMillis()` is within 60s of `Date.now()`.
- `MSAL itself drives the plugin end to end` — AC-1, AC-6; seed the document with a fabricated blob containing one `Account` entry, construct a `PublicClientApplication` with `cache: {cachePlugin}`, assert `getAllAccounts()` returns that account, then `removeAccount()` it and assert the document's `cache` field no longer contains the account id and `updatedAt` advanced.
- `two overlapping instances both survive: neither write is lost` — AC-3, AC-8; A and B both `beforeCacheAccess` on the same empty document, A adds account `acct-a` and commits, B adds `acct-b` and commits; the final document parses as valid JSON in one piece and contains both ids under `Account`.
- `five concurrent writers produce one coherent document` — AC-8; five plugin instances seeded from the same starting blob, all five `afterCacheAccess` calls started together and awaited with `Promise.all`; the final `cache` field parses and contains all five account ids. Firestore aborts and retries the losers, and the callback's re-read is what folds each retry's view forward.
- `a document holding a non-string cache field fails loudly` — scope gap 1 on the real backend; seed `{cache: 42}` directly through the client and assert `beforeCacheAccess` rejects with `TokenCacheError`.

**Verification:**
- [ ] `npm test` with no emulator running: passes, and the emulator tests report as skipped with their reason — not as passes
- [ ] `npm run typecheck` passes
- [ ] Manual run with the emulator started as in step 1: every test in `test/token-cache.emulator.test.ts` passes, and the output is pasted into the Findings section of this plan as the record that AC-6, AC-7, and AC-8 were actually exercised
- [ ] During the manual run, `updatedAt` on the document written by the round-trip test is a real server timestamp, confirmed by reading the document back through the client

## Phase 3: Documentation

**Goal:** Record the new module and the exact commands for the manual emulator run, so the next agent and the next operator do not have to rediscover them.
**Addresses:** AC-6

**Files:**
- Modify: `CLAUDE.md`
- Modify: `README.md`

**Steps:**

1. `CLAUDE.md` directory tree: add `src/token-cache.ts` with the comment `# Firestore-backed MSAL ICachePlugin (issue #7)`, plus `test/token-cache.test.ts` and `test/token-cache.emulator.test.ts`.
2. `CLAUDE.md` prose under the tree: state that `test/token-cache.emulator.test.ts` skips itself with a printed reason unless `FIRESTORE_EMULATOR_HOST` is set, that the emulator is started by hand rather than by a wrapper script, and that the plugin's Firestore-touching behaviour has no automated coverage without it — the same "explicit skip, never silent" rule as `RUN_DOCKER_TESTS`.
3. `CLAUDE.md` Conventions: add two entries, each stating the mechanism then the consequence.
   - **The token cache is one opaque string in one document.** The plugin stores `serialize()`'s output verbatim in `cache` plus an `updatedAt` server timestamp. Do not parse the blob or split it across documents; MSAL owns its shape and it changes between library versions.
   - **`afterCacheAccess` re-reads inside the transaction.** `--max-instances=1` does not prevent two instances existing during a revision transition. The transaction re-reads the document, and a blob differing from the one this instance last saw is fed back through `deserialize` before serialising, because MSAL's deserialize merges. Removing the re-read turns an overlap into a lost refresh token, which wedges the cache until bootstrap is re-run.
4. `README.md`, a short subsection after the Scripts table titled "Token cache tests": the two commands from Phase 2 step 1 verbatim, the `sudo apt-get install google-cloud-cli-firestore-emulator` prerequisite and why `gcloud components install` will not work here, the fact that `java` must be on `PATH`, and that without the emulator those tests skip with a printed reason rather than failing.
5. `README.md` Configuration section: no table change — both variables are already documented at `README.md:95` and `README.md:96`. Add one sentence stating that `FIRESTORE_CACHE_DOC` is the document the MSAL cache plugin reads and writes, and that its value must be a document path (even segment count), which `loadConfig` enforces.

**Tests added/updated:**
- None. This phase changes only Markdown; the behaviour it describes is covered by Phases 1 and 2. Documentation accuracy is checked in verification below by running the commands the docs name.

**Verification:**
- [ ] `npm run typecheck` passes
- [ ] `npm test` passes
- [ ] Every path named in `CLAUDE.md`'s tree exists
- [ ] The two commands in the README's "Token cache tests" subsection are copy-pasted and run during the Phase 2 manual verification, unedited

## Final Verification

- [ ] `npm test` passes with no emulator present; the emulator tests report as skipped with their reason
- [ ] `npm run typecheck` passes
- [ ] `npm run build` succeeds and `dist/token-cache.js` exists
- [ ] `bash scripts/test/run.sh` passes (unchanged by this issue, run to prove it was not broken)
- [ ] Manual emulator run completed and its output recorded in Findings
- [ ] `git grep -n -i "onmicrosoft\|tenant" src/ test/` finds no real tenant name or id; every fixture GUID is fabricated
- [ ] Acceptance criteria traceability:
  - AC-1 (`ICachePlugin` against one document, path from `FIRESTORE_CACHE_DOC`): Phase 1 — `FirestoreTokenCachePlugin` implements both callbacks against the injected path, and `createFirestoreTokenCachePlugin` takes the path from `FirestoreConfig.cacheDocumentPath`, whose default `tokencache/msal` lives in `src/config.ts:50`. Exercised by `MSAL itself drives the plugin end to end` (Phase 2).
  - AC-2 (write only when `cacheHasChanged`): Phase 1 — the early return in `afterCacheAccess`; verified by `afterCacheAccess writes nothing when cacheHasChanged is false` (Phase 2).
  - AC-3 (transactional read-modify-write): Phase 1 — the `runTransaction` body with its re-read and merge; verified by `two overlapping instances both survive` (Phase 2).
  - AC-4 (single string field plus `updatedAt`, no modelling of MSAL internals): Phase 1 — the two-field `tx.set` and the opaque-blob decoder, unit-tested by `extra fields on the document are ignored`; verified on a real backend by `the stored document has exactly cache and updatedAt` (Phase 2).
  - AC-5 (absent document is an empty cache): Phase 1 — `readCache` returns `null` for a missing document, unit-tested by `an absent document reads as an empty cache`; repeated against real Firestore in Phase 2.
  - AC-6 (emulator-backed tests including concurrency): Phase 2 — `test/token-cache.emulator.test.ts`, documented in Phase 3. Deviation recorded above: the emulator cannot be installed by the agent, so the tests are committed but run manually, and no substitute Firestore is built.
  - AC-7 (byte-identical round-trip through a second instance): Phase 2 — `a cache written by one instance is byte-identical when read by a second`.
  - AC-8 (concurrent writes produce one coherent document): Phase 2 — `two overlapping instances both survive` and `five concurrent writers produce one coherent document`.

## Findings

> Filled in during implementation: what the work revealed that planning missed, and the pasted output of the manual emulator run.

### Phase 1

- **`TokenCacheError` carries the path as a field, not only in the message.** The plan said the message carries the document path. The implementation adds `readonly documentPath: string`, mirroring `ConfigError`'s structured `missing`/`invalid` fields in `src/config.ts:14`. The test asserts the path by field rather than by matching prose, so rewording the message cannot silently break it. The message text is unchanged from the plan.
- **`exists === true` is not sufficient to decode.** `readCache` returns `null` when `snapshot.exists` is true but `data()` is `undefined`, not just when `exists` is false. The plan's prose said this; it is recorded here because the branch condition is two checks, not one. Covered by the first test.
- **No typing workaround was needed.** `DocumentSnapshot` satisfies `CacheSnapshot` structurally — `readonly exists: boolean` and `data(): DocumentData | undefined`, where `DocumentData`'s index signature is assignable to `Record<string, unknown>` — so neither `this.#ref.get()` nor `tx.get(this.#ref)` needs a cast, and `tx.set` accepted the two-field computed-key object without a generic annotation. `src/token-cache.ts` contains no `as any`, `@ts-ignore`, or `@ts-expect-error`.
- **Verified:** `npm run typecheck` exit 0; `npm test` 24 pass / 0 fail / 0 skipped; `npm run build` emits `dist/token-cache.js`.
