# Graph auth module: silent token acquisition, named error on dead refresh token

**Issue:** #8
**Goal:** Give the server one module that turns the Firestore-backed MSAL cache into a Graph access token silently, and that fails with an error naming the bootstrap CLI rather than letting a bare 401 or a raw MSAL error reach the caller.

**Date:** 2026-08-18

## Status

| Phase | Status | Commit |
|-------|--------|--------|
| Phase 1: `src/graph-auth.ts` — scopes, named error, silent acquisition, unit tests | ☑ | this commit |
| Phase 2: Documentation — `README.md`, `CLAUDE.md` | ☐ | — |

## Acceptance Criteria

> Reproduced from issue #8 so the plan is self-contained. The six "Tasks" checkboxes are AC-1 to AC-6; the two sentences under "Acceptance" are AC-7 and AC-8.

- **AC-1:** `PublicClientApplication` configured with `ONENOTE_CLIENT_ID`, `ONENOTE_AUTHORITY`, and the Firestore cache plugin.
- **AC-2:** `getAccessToken()`: `getTokenCache().getAllAccounts()` → `acquireTokenSilent`. If there are no accounts, or silent acquisition fails, throw a named error.
- **AC-3:** The error message must say what actually happened and what to do: the refresh token is expired or revoked, and the local bootstrap CLI needs re-running. A bare 401 propagated to the caller is the failure mode to avoid.
- **AC-4:** Do not call `acquireTokenByDeviceCode` anywhere in the server path. It is bootstrap-only.
- **AC-5:** Cache the in-process token and let MSAL handle expiry; do not hand-roll expiry logic.
- **AC-6:** Scopes: `Notes.Read`, `Notes.ReadWrite`.
- **AC-7:** With a seeded cache, `getAccessToken()` returns a token and the Firestore document's `updatedAt` advances after a refresh. ⚠️ Deviation: **no automated test.** A seeded cache requires the device-code sign-in from issue #9, run against the real Entra app registration by issue #10. Neither exists yet, and no credential that could seed one may be committed. Verified instead by the manual procedure below, which the operator runs after #10. A mock cannot substitute: what is being checked is that Microsoft accepts the stored refresh token and that the plugin writes the rotated one back.
- **AC-8:** With a deliberately corrupted cache, the thrown error names the bootstrap CLI.

## Context

### What exists today

Issues #5, #6 and #7 are complete. `src/` holds `index.ts`, `server.ts`, `config.ts`, `bootstrap.ts`, `token-cache.ts`, `version.ts`. Nothing in `src/` constructs an MSAL client — `@azure/msal-node@5.5.0` is imported only for its types, by `src/token-cache.ts`.

The two pieces this issue joins already exist:

- `loadConfig(['graph'])` returns `GraphConfig { clientId: string; authority: string }` from `ONENOTE_CLIENT_ID` and `ONENOTE_AUTHORITY` (`src/config.ts:96`, specs at `src/config.ts:36` and `src/config.ts:42`). `ONENOTE_AUTHORITY` is already validated as an absolute https URL. No config change is needed for this issue.
- `createFirestoreTokenCachePlugin(config: FirestoreConfig): FirestoreTokenCachePlugin` (`src/token-cache.ts:131`) returns an `ICachePlugin` ready to hand to MSAL's `cache.cachePlugin` option.

`src/server.ts` builds the Express app and mounts only `/healthz`. This issue adds no route and no wiring into `createApp` — the first consumer is the Graph structure client (issue #11), and the MCP transport lands in #14. Keeping the module unwired is deliberate: wiring it into `createApp` now would mean `createApp` constructs a Firestore client, and `test/server.test.ts` drives `createApp` with stub config on every `npm test` run.

`src/bootstrap.ts` is still the #5 placeholder. It is the only file that will ever call `acquireTokenByDeviceCode` (issue #9).

### Design specs consulted

There is no `docs/design/` directory. `project-spec.md` is authoritative.

- `project-spec.md:69` and `project-spec.md:71` — MSAL Node device-code flow, delegated auth, "Required scopes: `Notes.Read` and `Notes.ReadWrite`." This is AC-6.
- `project-spec.md:328`–`project-spec.md:337`, the two-OAuth-layers table — Layer 2 is "Device code once at bootstrap, then silent refresh forever." This module is the "silent refresh forever" half; the device-code half is issue #9.
- `project-spec.md:264` Constraint 2 — "Bootstrap sign-in happens locally, never on Cloud Run." This is AC-4's justification.
- `project-spec.md:169` — Graph's OneNote endpoints do not support app-only auth, so there is no client-credentials fallback when the refresh token dies. The only recovery is a human re-running bootstrap, which is why AC-3 demands the error say so.
- `project-spec.md:470` (Appendix A, the validated script) — `const SCOPES = ["https://graph.microsoft.com/Notes.Read"]`, the fully-qualified scope form.
- `project-spec.md:394`, `project-spec.md:402` and the repo-hygiene section of `CLAUDE.md` — the token cache is secret, and the tenant name and ID must never appear in source or logs. This constrains what the error messages may include; see the Issue Assessment.

### Issue Assessment

Every claim below was checked against a file in this repository or in `node_modules`.

**Nothing stale.** `ONENOTE_CLIENT_ID` and `ONENOTE_AUTHORITY` are live at `src/config.ts:36` and `src/config.ts:42`. `createFirestoreTokenCachePlugin` is live at `src/token-cache.ts:131`. `@azure/msal-node@5.5.0` is installed.

**Verified: the API surface.** `node_modules/@azure/msal-node/types/client/IPublicClientApplication.d.ts` declares `getTokenCache(): TokenCache` and `acquireTokenSilent(request: SilentFlowRequest): Promise<AuthenticationResult>` — non-nullable, unlike `acquireTokenByRefreshToken` and `acquireTokenByDeviceCode` which both return `Promise<AuthenticationResult | null>`. `ITokenCache` (`node_modules/@azure/msal-node/types/cache/ITokenCache.d.ts`) declares `getAllAccounts(): Promise<AccountInfo[]>`, so the issue's `getTokenCache().getAllAccounts()` is an async call, not the synchronous browser-MSAL one. `SilentFlowRequest` requires `account: AccountInfo` and `scopes: string[]`. `CacheOptions.cachePlugin?: ICachePlugin` is where the plugin goes (`node_modules/@azure/msal-node/types/config/Configuration.d.ts:43`).

**Scope gap 1: a corrupted cache does not fail inside `acquireTokenSilent`.** AC-2 names two failure paths — no accounts, and silent acquisition failing. AC-8's "deliberately corrupted cache" is a third, and it fires earlier: `getAllAccounts()` triggers `beforeCacheAccess`, so a blob that is a valid string but not valid MSAL JSON makes MSAL's `deserialize` throw from inside the `getAllAccounts()` call. A non-string `cache` field throws `TokenCacheError` from `readCache` at the same point. If only `acquireTokenSilent` is wrapped, AC-8's error reaches the caller as a raw `SyntaxError`. The `getAllAccounts()` call is therefore wrapped too, giving three named reasons rather than two.

**Scope gap 2: the error must not name the user or the tenant.** The obvious diagnostic detail — `account.username`, or `homeAccountId`, which is `<oid>.<tenantId>` — is exactly what `CLAUDE.md`'s hygiene rule forbids in logs that may reach an issue or an Actions log. The messages therefore carry the reason, the document path, and the underlying MSAL error code, and no account identifier. The account count is safe and is included.

**Scope excess: none.** All six task bullets describe work that does not exist yet.

**AC-5 read narrowly.** "Cache the in-process token" is satisfied by constructing `PublicClientApplication` once per process and reusing it: MSAL's in-memory token cache lives on that instance, and `acquireTokenSilent` returns a cached token without a network call while one is valid. Storing `accessToken` and comparing `expiresOn` in this module would be the hand-rolled expiry logic AC-5 forbids. No single-flight de-duplication of concurrent `getAccessToken()` calls is added either: overlapping refreshes are already handled by the transaction in `afterCacheAccess` (`src/token-cache.ts:104`), and adding a promise-sharing layer here would be untested concurrency code with no failure it prevents.

## Phase 1: `src/graph-auth.ts` — scopes, named error, silent acquisition, unit tests

**Goal:** One module that exports the scope list, a named error type, a testable silent-acquisition function over a narrow client interface, and a factory that builds the real `PublicClientApplication`.

**Addresses:** AC-1, AC-2, AC-3, AC-4, AC-5, AC-6, AC-8

**Files:**
- Create: `src/graph-auth.ts`
- Test: `test/graph-auth.test.ts`

**Steps:**

1. **`GRAPH_SCOPES`.** Export `export const GRAPH_SCOPES: readonly string[] = ['https://graph.microsoft.com/Notes.Read', 'https://graph.microsoft.com/Notes.ReadWrite']`. Fully-qualified, matching `project-spec.md:470`, so the bootstrap CLI (#9) imports this same constant and the account it seeds carries the scopes this module asks for. A comment records why: MSAL keys cached access tokens by scope string, so the short form here and the long form in bootstrap would look like two different grants.

2. **`GraphAuthError`.** Export a class extending `Error`, `name = 'GraphAuthError'`, with `readonly reason: GraphAuthErrorReason` and `readonly documentPath: string`, mirroring `TokenCacheError`'s structured field at `src/token-cache.ts:19`. `export type GraphAuthErrorReason = 'cache-unreadable' | 'no-account' | 'silent-failed'`. The constructor takes `(reason, documentPath, options?: { cause?: unknown })` and builds the message from the reason, so the tests assert on `reason` rather than on prose. Each message states the mechanism, then the fix:
   - `cache-unreadable`: the token cache at `<documentPath>` could not be read or decoded; it is absent, empty, or its contents are not a cache MSAL recognises; re-run `npm run bootstrap` on a machine with a browser to recreate it.
   - `no-account`: the token cache at `<documentPath>` was read but holds no signed-in account, so there is no refresh token to exchange; run `npm run bootstrap`.
   - `silent-failed`: silent token acquisition failed for the account in `<documentPath>`; the stored refresh token is expired or revoked, and the deployed server never signs in interactively, so it cannot recover on its own; re-run `npm run bootstrap`.
   Every message ends with the literal string `npm run bootstrap`. When a `cause` is present, append `Underlying error: <name>: <message>` built from the cause with a helper that handles a non-`Error` throw. Do not append the cause's stack. Do not include `account.username` or `homeAccountId` — see Issue Assessment, scope gap 2.

3. **`SilentTokenSource`.** Export a narrow structural interface covering only what this module calls, the same technique `CacheSnapshot` uses at `src/token-cache.ts:35`:
   ```ts
   export interface SilentTokenSource {
     getTokenCache(): { getAllAccounts(): Promise<AccountInfo[]> };
     acquireTokenSilent(request: SilentFlowRequest): Promise<AuthenticationResult | null>;
   }
   ```
   `PublicClientApplication` satisfies it structurally, so the factory passes one with no cast, and a test passes a hand-written fake. Import `AccountInfo`, `AuthenticationResult`, `SilentFlowRequest` with `import type` (`verbatimModuleSyntax`).

4. **`acquireGraphToken(client, documentPath)`.** Exported async function returning `Promise<string>`:
   - `let accounts: AccountInfo[]`; `try { accounts = await client.getTokenCache().getAllAccounts() } catch (err) { throw new GraphAuthError('cache-unreadable', documentPath, { cause: err }) }`. This is the AC-8 path: `getAllAccounts()` is what triggers `beforeCacheAccess`, so a `TokenCacheError` from `readCache` or a `SyntaxError` from MSAL's `deserialize` both surface here and both become a named error.
   - `if (accounts.length === 0) throw new GraphAuthError('no-account', documentPath)`.
   - Take `accounts[0]`. A comment records that bootstrap seeds exactly one account, and that if a second ever appears the first is used rather than guessing — with more than one account the count appears in no message, only in the `accounts.length === 0` branch.
   - `try { result = await client.acquireTokenSilent({ account, scopes: [...GRAPH_SCOPES] }) } catch (err) { throw new GraphAuthError('silent-failed', documentPath, { cause: err }) }`. The spread produces the mutable `string[]` `SilentFlowRequest` requires without exporting a mutable constant.
   - `if (result === null || result.accessToken === '') throw new GraphAuthError('silent-failed', documentPath)`. `IPublicClientApplication` types this non-nullable, but the interface in step 3 widens it so a fake can return `null`, and an empty `accessToken` is the shape a caller would otherwise send to Graph and get a 401 for — the exact failure mode AC-3 exists to prevent.
   - `return result.accessToken`.

5. **`GraphAuth`.** Export a class holding the `SilentTokenSource` and the document path, with one method `getAccessToken(): Promise<string>` that returns `acquireGraphToken(this.#client, this.#documentPath)`. It stores no token and no expiry — MSAL's in-memory cache on the retained client is the in-process cache (AC-5).

6. **`createGraphAuth(graph: GraphConfig, firestore: FirestoreConfig): GraphAuth`.** Builds `createFirestoreTokenCachePlugin(firestore)`, then `new PublicClientApplication({ auth: { clientId: graph.clientId, authority: graph.authority }, cache: { cachePlugin } })`, then `new GraphAuth(pca, firestore.cacheDocumentPath)` (AC-1). This is the only `PublicClientApplication` construction in `src/`, and the module imports no device-code request type (AC-4).

7. **Tests** in `test/graph-auth.test.ts`, following the fake-object style of `test/token-cache.test.ts`. A `fakeSource(overrides)` helper builds a `SilentTokenSource` whose `getAllAccounts` and `acquireTokenSilent` are supplied per test, and a `fakeAccount()` helper returns an `AccountInfo` with fabricated values — fake `homeAccountId`, tenant `00000000-0000-0000-0000-000000000000`, `username: 'nobody@example.invalid'`. No real tenant, no real token.

**Tests added/updated:**

- `GRAPH_SCOPES is exactly Notes.Read and Notes.ReadWrite, fully qualified` — asserts the array contents (AC-6).
- `a token is returned and the request carries the scopes and the cached account` — a fake returning one account and `{ accessToken: 'fake-token-value' }`; asserts the return value and captures the `SilentFlowRequest`, asserting `request.account` is the account from the cache and `request.scopes` deep-equals `GRAPH_SCOPES` (AC-2, AC-6).
- `an empty cache throws GraphAuthError with reason no-account` — `getAllAccounts` resolves `[]`; asserts `err instanceof GraphAuthError`, `err.name === 'GraphAuthError'`, `err.reason === 'no-account'`, `err.documentPath`, and `/npm run bootstrap/` in the message (AC-2, AC-3).
- `a cache that will not decode throws GraphAuthError with reason cache-unreadable` — two cases: `getAllAccounts` rejects with `new TokenCacheError(...)`, and rejects with `new SyntaxError('Unexpected token')`. Both must produce `reason === 'cache-unreadable'`, name `npm run bootstrap`, and expose the original on `err.cause` (AC-8).
- `a dead refresh token throws GraphAuthError with reason silent-failed and does not leak the account` — `acquireTokenSilent` rejects with an error shaped like MSAL's `InteractionRequiredAuthError` (`errorCode: 'invalid_grant'`). Asserts `reason === 'silent-failed'`, that the message contains `expired or revoked` and `npm run bootstrap`, that `err.cause` is the original, and that the message contains neither the account's `username` nor its `homeAccountId` (AC-3, scope gap 2).
- `a null or blank result throws rather than returning an unusable token` — `acquireTokenSilent` resolving `null`, and resolving `{ accessToken: '' }`; both throw with `reason === 'silent-failed'` (AC-3).
- `no module under src/ calls acquireTokenByDeviceCode except bootstrap.ts` — reads every `*.ts` under `src/` with `readdir`/`readFile` and asserts the string `acquireTokenByDeviceCode` appears in none of them but `src/bootstrap.ts`. This is a source-text guard, not a behavioural test, and its comment says so: the AC is about a call site that must never be added to the server path, and no runtime assertion can observe a call that is absent (AC-4).

**Verification:**
- [ ] `npm test` passes, no test skipped
- [ ] `npm run typecheck` passes (this is the lint gate — `CLAUDE.md`, Commands)
- [ ] `npm run build` succeeds and `dist/graph-auth.js` exists
- [ ] `git grep -n "acquireTokenByDeviceCode" -- src/` returns nothing
- [ ] `git grep -n -iE "onmicrosoft|[a-z0-9]+\.sharepoint" src/ test/` returns nothing — no real tenant reached the fixtures

## Phase 2: Documentation

**Goal:** Record what the module is, what its tests do and do not cover, and the two conventions a future agent could plausibly break.

**Addresses:** AC-3, AC-4, AC-5, AC-7

**Files:**
- Modify: `CLAUDE.md`
- Modify: `README.md`

**Steps:**

1. `CLAUDE.md` directory tree: add `src/graph-auth.ts` with the comment `# Layer-2 Graph auth: silent token acquisition (issue #8)` and `test/graph-auth.test.ts` in the `test/` block.
2. `CLAUDE.md`, in the paragraph after the tree: one sentence stating that `test/graph-auth.test.ts` drives `acquireGraphToken` through a hand-written `SilentTokenSource` and never constructs a `PublicClientApplication`, so `createGraphAuth` has no automated test — it needs a seeded cache and a real Entra app registration, which arrive with issues #9 and #10, and it is covered by the manual procedure in this plan.
3. `CLAUDE.md` Conventions: add two entries, mechanism then consequence.
   - **The server never signs in interactively.** `acquireTokenByDeviceCode` belongs to `src/bootstrap.ts` alone. The deployed service acquires silently from the Firestore cache; a device-code call in a request path would block on a human who is not there, and Cloud Run would time the request out instead. A test in `test/graph-auth.test.ts` scans `src/` for the call.
   - **Graph auth failures are `GraphAuthError`, never a raw MSAL error.** `acquireGraphToken` wraps both the `getAllAccounts()` call and the `acquireTokenSilent` call, so a decode failure, an empty cache, and a dead refresh token all arrive as one error type whose message ends in `npm run bootstrap`. Propagating the MSAL error instead surfaces to the operator as a bare 401 from Graph, which does not say that a human has to re-run the CLI.
4. `README.md`: new `## Graph auth` section after `## Token cache`. What `src/graph-auth.ts` does, the scopes it requests, the fact that the deployed server never signs in interactively, and the three `GraphAuthError` reasons with the operator action for each — all three being `npm run bootstrap`. State that nothing wires it into `createApp` yet; the first consumer is issue #11.
5. `README.md`, `## Token cache`: correct the last sentence of the first paragraph. It currently reads "that is issue #9 for the bootstrap CLI and issue #12 for the server" — #12 is the ink pipeline. The server-side consumer is this issue, #8. Change it to name #9 and #8.
6. `README.md` `## Configuration`: after the table, one sentence stating that `ONENOTE_CLIENT_ID` and `ONENOTE_AUTHORITY` are the public-client app registration that `src/graph-auth.ts` presents to Entra, and that there is deliberately no client secret for Layer 2.

**Tests added/updated:**
- None. This phase changes only Markdown; the behaviour it describes is covered by Phase 1's tests and the manual procedure below. The one factual claim it adds that Phase 1 does not check — that every path in `CLAUDE.md`'s tree exists — is in this phase's verification.

**Verification:**
- [ ] `npm run typecheck` passes
- [ ] `npm test` passes
- [ ] Every path named in `CLAUDE.md`'s directory tree exists on disk
- [ ] `README.md` no longer attributes the server-side token-cache consumer to issue #12

## Manual verification procedure

This stands in for AC-7. It cannot run until issue #10 has seeded the cache with a real device-code sign-in, so it is run by the operator after #10, not during this issue. It touches the real Entra app registration and the real Firestore document; it is not a test and nothing it produces is committed.

**Prerequisites:** issues #9 and #10 complete, so `tokencache/msal` (or whatever `FIRESTORE_CACHE_DOC` names) holds a seeded cache. Application Default Credentials for the project on the machine running it. `ONENOTE_CLIENT_ID`, `ONENOTE_AUTHORITY`, `GOOGLE_CLOUD_PROJECT` and `FIRESTORE_CACHE_DOC` set in the environment.

Write a throwaway script under the scratchpad — not in the repository, and never committed. It calls `loadConfig(['graph', 'firestore'])`, then `createGraphAuth(config.graph, config.firestore)`, then `getAccessToken()`.

**Step 1 — a token comes back (AC-7, first half).** Run it. It must print a non-empty token. Print only `token.length` and the first eight characters, never the whole token — the output of this step goes into Findings, and Findings is committed.

**Step 2 — `updatedAt` advances after a refresh (AC-7, second half).** Read the document's `updatedAt` before and after. A first call within the access token's lifetime may serve from MSAL's in-memory cache and write nothing, which is AC-5 working correctly, so force a refresh: construct a second `GraphAuth` in a fresh process, which starts with an empty in-memory cache and must go to the token endpoint. Assert the second run's `updatedAt` is strictly greater than the first's. Record both timestamps.

**Step 3 — a corrupted cache produces the named error against a real backend (AC-8).** Copy the document's `cache` field to a scratchpad file first — this is the live refresh token and overwriting it without a copy means re-running bootstrap. Then set `cache` to `'not a cache'` and run the script: it must throw `GraphAuthError` with `reason === 'cache-unreadable'` and a message ending in `npm run bootstrap`. Restore the saved blob and re-run step 1 to confirm the cache still works.

**Step 4 — no interactive prompt (AC-4).** Confirm that no step printed a device code or a verification URL.

**Recording.** Paste the script's output — with the token itself redacted — into Findings and tick the manual box in Final Verification. An untouched Findings entry means AC-7 is unverified.

## Final Verification

- [ ] `npm test` passes; no test skipped
- [ ] `npm run typecheck` passes
- [ ] `npm run build` succeeds and `dist/graph-auth.js` exists
- [ ] `bash scripts/test/run.sh` passes (unchanged by this issue; run to prove it was not broken)
- [ ] `git grep -n "acquireTokenByDeviceCode" -- src/` returns nothing
- [ ] `git grep -n -iE "onmicrosoft|[a-z0-9]+\.sharepoint" src/ test/` returns nothing
- [ ] Manual verification procedure completed by the operator after issue #10, output recorded in Findings. Until this box is ticked, AC-7 is unverified.
- [ ] Acceptance criteria traceability:
  - AC-1 (`PublicClientApplication` with client id, authority, cache plugin): Phase 1 step 6 — `createGraphAuth` is the single construction site and takes all three from `GraphConfig` and `createFirestoreTokenCachePlugin(FirestoreConfig)`. Exercised end to end by manual step 1.
  - AC-2 (`getAllAccounts()` → `acquireTokenSilent`, named error on no account or failure): Phase 1 step 4 — tested by `a token is returned and the request carries the scopes and the cached account`, `an empty cache throws GraphAuthError with reason no-account`, and `a dead refresh token throws GraphAuthError with reason silent-failed and does not leak the account`.
  - AC-3 (message says what happened and what to do; no bare 401): Phase 1 step 2 — every `GraphAuthError` message ends with `npm run bootstrap`, asserted in each error test; `a null or blank result throws rather than returning an unusable token` covers the path that would otherwise produce the 401. Documented in Phase 2 step 4.
  - AC-4 (no `acquireTokenByDeviceCode` in the server path): Phase 1 step 7 — the source-scan test over `src/`, plus the `git grep` check above. Documented as a convention in Phase 2 step 3.
  - AC-5 (in-process token cache, MSAL owns expiry): Phase 1 steps 5 and 6 — `GraphAuth` retains one `PublicClientApplication` and stores no token or expiry. Observed in manual step 2, where a same-process second call does not necessarily rewrite the document.
  - AC-6 (`Notes.Read`, `Notes.ReadWrite`): Phase 1 step 1 — `GRAPH_SCOPES`, tested by `GRAPH_SCOPES is exactly Notes.Read and Notes.ReadWrite, fully qualified` and by the scope assertion in the success test.
  - AC-7 (seeded cache returns a token, `updatedAt` advances): manual steps 1 and 2. Deviation recorded above — needs issues #9 and #10 first.
  - AC-8 (corrupted cache names the bootstrap CLI): Phase 1 step 4 — `a cache that will not decode throws GraphAuthError with reason cache-unreadable`. Confirmed against the real document by manual step 3.

## Findings

> Filled in during implementation: what the work revealed that planning missed, and the recorded output of the manual verification procedure.

### Phase 1

- **`noUncheckedIndexedAccess` collapses the empty-cache check into one branch.** The plan had a `accounts.length === 0` check followed by `accounts[0]`. Under `noUncheckedIndexedAccess` (`tsconfig.json:14`) that index is `AccountInfo | undefined`, so it would need a cast. The implementation destructures — `const [account] = accounts` — and throws `no-account` when it is `undefined`. One check, no cast; `src/graph-auth.ts` contains no `as`, `@ts-ignore`, or `@ts-expect-error`.
- **`cause` is assigned, not passed to `super`.** `GraphAuthError`'s constructor takes `options: { cause?: unknown }` and sets `this.cause` only when the key is present. Passing `options` straight to `super` would give every error an own `cause` property of `undefined` on the two paths that have no underlying error (`no-account`, and the null-or-blank-result branch of `silent-failed`), which the tests then could not distinguish from a wrapped error.
- **The source-scan test asserts it scanned something.** `no module under src/ calls acquireTokenByDeviceCode except bootstrap.ts` first asserts that the file list includes `graph-auth.ts`. Without that, a `readdir` returning an empty list — a moved directory, a changed `import.meta.dirname` — would make the test pass while checking nothing.
- **Verified:** `npm run typecheck` exit 0; `npm test` 31 pass / 0 fail / 0 skipped; `npm run build` emits `dist/graph-auth.js`; `git grep acquireTokenByDeviceCode -- src/` returns nothing.
