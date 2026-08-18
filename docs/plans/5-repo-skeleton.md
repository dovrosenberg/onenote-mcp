# Repo skeleton: TypeScript ESM project, config validation, .gitignore

**Issue:** #5
**Goal:** Stand up the TypeScript ESM project — dependencies, build, test runner, fail-fast env-var validation, and a `/healthz` entrypoint that binds `process.env.PORT` — so every later issue has a compiling, testable base.
**Date:** 2026-08-18

## Status

| Phase | Status | Commit |
|-------|--------|--------|
| Phase 1: Project skeleton, dependencies, build and test wiring | ☑ | `e6f00dc` |
| Phase 2: `src/config.ts` — grouped env schema and fail-fast validation | ☐ | — |
| Phase 3: `src/index.ts` entrypoint, `/healthz`, bootstrap placeholder | ☐ | — |

## Acceptance Criteria

> Reproduced from issue #5 so the plan is self-contained. The "Tasks" checkboxes, the "Acceptance" line, and the "Assumption to confirm" are all treated as acceptance criteria.

- **AC-1:** `.gitignore` covering `node_modules/`, `output/`, `dist/`, `.env`, `*.token-cache.json`, `.msal-cache*`.
- **AC-2:** `package.json`: ESM (`"type": "module"`), Node >= 20, scripts for `build`, `dev`, `start`, `bootstrap`, `test`. ⚠️ Deviation: `engines.node` is `">=24"`, not `">=20"` — see Issue Assessment finding 1.
- **AC-3:** TypeScript + `tsconfig.json` targeting Node 20, `moduleResolution: nodenext`. ⚠️ Deviation: targets Node 24 (`target: es2024`), not Node 20, for the same reason. `moduleResolution: nodenext` is unchanged.
- **AC-4:** Dependencies: `@modelcontextprotocol/sdk`, `@azure/msal-node`, `@google-cloud/firestore`, `fast-xml-parser`, `@resvg/resvg-js`, an HTTP framework (express or hono). Resolved to **Express 5** — see Issue Assessment finding 2.
- **AC-5:** `src/config.ts` — read and validate every env var at startup, fail fast with a named error listing what is missing. Bind to `process.env.PORT`, never a hardcoded port.
- **AC-6:** Test runner wired up (`node --test` is sufficient; no framework).
- **AC-7:** `README.md` stub pointing at `project-spec.md`.
- **AC-8:** `npm ci && npm run build` succeeds from a clean checkout.
- **AC-9:** `npm start` exits with a clear message listing missing env vars rather than a stack trace.
- **AC-10:** TypeScript rather than the plain `.mjs` of Appendix A. Confirmed — the plan is written in TypeScript.

## Context

### What exists today

`LICENSE`, `project-spec.md` (870 lines), `docs/plans/2-gcp-bootstrap-script.md`, and `scripts/` containing `gcp-bootstrap.sh` plus its stub-based shell test harness (`scripts/test/run.sh`, `scripts/test/bootstrap.test.sh`, `scripts/test/stubs/gcloud`). There is no `package.json`, no `.gitignore`, no `src/`, no TypeScript, and no JavaScript test runner. Phase 0 (issues #1–#4) is closed, so the GCP project, Firestore database, Artifact Registry repo, WIF provider, and the GitHub repo variables and secrets all exist already.

This issue is the first one that writes application code. Nothing in the repo constrains its structure, so the conventions established here become the project's conventions.

### Design specs consulted

There is no `docs/design/` directory. `project-spec.md` is authoritative. Relevant sections:

- `project-spec.md:160` — "Config via environment variables: Azure client ID, tenant/authority, the Layer-1 OAuth client ID and secret, the access-token signing key, the Firestore document path for the token cache, and the bind port (`PORT`). Cloud Run sets `PORT` itself; the server must bind to it rather than a hardcoded value." This is the complete env-var list for AC-5.
- `project-spec.md:197` "Where each piece of configuration lives" — fixes the exact env var *names* the service receives.
- `project-spec.md:280` "Workflow shape" — the `env_vars:` block of the deploy step is the definitive list of what Cloud Run will actually set: `ONENOTE_CLIENT_ID`, `ONENOTE_AUTHORITY`, `MCP_OAUTH_CLIENT_ID`, `MCP_OAUTH_CLIENT_SECRET`, `MCP_TOKEN_SIGNING_KEY`, `FIRESTORE_CACHE_DOC=tokencache/msal`. `config.ts` must not require anything outside this set, or the first deploy (#25) fails at startup.
- `project-spec.md:392` "Repo hygiene" — "`.gitignore` must exclude the token cache and any output/scratch directories. Real page dumps contain personal notes." This is the reason AC-1 lists `output/` and the cache-file globs, and the repo is public.
- `project-spec.md:422` Appendix A `package.json` — the validated dependency set (`@azure/msal-node`, `@resvg/resvg-js`, `fast-xml-parser`) and `"type": "module"`.

### Issue Assessment

Findings from checking the issue against the current package ecosystem. Every claim below was verified by running the command, not inferred.

**Scope adjustments — deviations from the issue text:**

1. **Node >= 20 is not satisfiable with the current dependency set.** `@google-cloud/firestore@9.0.0` declares `engines.node: ">=22"`. Holding the line at Node 20 means pinning Firestore to `^8.7.1`, a trailing major. The user directed a review of whether Node 24 is viable instead; it is, and it is the better target:
   - Node 24 is Active LTS as of 2026-08. Node 26 is Current and does not reach LTS until 2026-10.
   - The full dependency tree installs with `exit=0` under `npm install --engine-strict` with `engines.node: ">=24"`.
   - `@resvg/resvg-js@2.6.2` is N-API with `engines.node: ">= 10"`; its `@resvg/resvg-js-linux-x64-gnu` prebuild declares `libc: ["glibc"]`. It was loaded on Node 24.13.0 and rendered a test SVG to a 79-byte PNG. glibc means the Docker runtime base in #6 must stay Debian (`-slim`), not alpine.
   - `typescript@7.0.2` (the current `latest`) compiled a file importing all seven packages under `strict` + `moduleResolution: nodenext` with `exit=0`.

   **This changes an AC on issue #6**, which specifies a `node:20-slim` runtime stage. That must become `node:24-slim`. Flag it on #6 rather than silently diverging.

2. **HTTP framework resolved to Express 5** (`express@5.2.1` + `@types/express@5.0.6`). The MCP TypeScript SDK's `StreamableHTTPServerTransport` takes Node's `IncomingMessage`/`ServerResponse`, which is what Express hands a middleware. Hono's handlers take Web-standard `Request`/`Response` and would need an adapter shim in #14. Every SDK example for Streamable HTTP is written against Express, so #14 and #21–#23 can follow the SDK docs without translation.

3. **`node --test` runs TypeScript directly on Node 24 — no build step for tests.** Native type stripping is unflagged in Node 24; `node --test "src/**/*.test.ts"` was run and passed with no warning. This is worth locking in deliberately, because it constrains the source: type stripping erases types but cannot transform syntax, so `enum`, `namespace`, and constructor parameter properties are forbidden, and type-only imports must be written `import type`. Two compiler options enforce that mechanically rather than by discipline — `erasableSyntaxOnly: true` (verified: rejects `enum E { A }` with `error TS1294`) and `verbatimModuleSyntax: true`.

4. **The `.ts` extension in import specifiers needs two compiler options.** Under `moduleResolution: nodenext` the usual style is to import `./config.js` from `config.ts`. That breaks `node --test` on source, because Node looks for a real `./config.js` that does not exist until after a build. The fix is to write `import { … } from './config.ts'` and set `allowImportingTsExtensions: true` with `rewriteRelativeImportExtensions: true`. Verified: `tsc` emits `import { NAME } from './config.js';` into `dist/`, `node dist/index.js` runs, and `node --test` on `src/` also runs. Both paths work from one source form.

**Scope gap — added beyond what the issue lists:**

5. **Config validation must be grouped, not all-or-nothing.** The issue says "validate every env var at startup," but two entrypoints need different subsets. `src/index.ts` (the server) needs the Layer-1 OAuth vars; `src/bootstrap.ts` (issue #9, the local device-code CLI) does not, and must not — the operator running it locally has no reason to hold `MCP_OAUTH_CLIENT_SECRET`. A single flat `loadConfig()` that demands all of them would make `npm run bootstrap` unrunnable. Phase 2 therefore builds `loadConfig()` over named groups (`graph`, `firestore`, `server`, `oauth`) and each entrypoint requests the groups it needs.

6. **`npm run bootstrap` would be a dangling script reference.** AC-2 requires the script, but `src/bootstrap.ts` is issue #9's deliverable. Phase 3 creates a minimal `src/bootstrap.ts` that loads the `graph` and `firestore` config groups, prints "not implemented — see issue #9", and exits 1. That keeps the script honest, exercises the grouped loader, and does not do #9's work.

**Scope excess — one item deliberately left out.**

`MCP_PUBLIC_URL` (or equivalent) will be needed by issue #21 for the `issuer` field in the OAuth metadata documents. It is not added here. It is absent from the `env_vars:` block in `project-spec.md:280`, so making it required now would fail startup on the first deploy. Issue #21 adds it to both `config.ts` and the workflow together.

**Nothing stale.** No AC references a deleted or renamed file — there is no implementation code yet.

### Lint

The repo has no JavaScript linter and this issue does not ask for one, so adding ESLint would be scope creep. The verification step named "lint" in each phase below is `npm run typecheck` (`tsc --noEmit`) — a full-program static check that fails on the same class of error a linter would catch here. If a linter is wanted it belongs in its own issue.

---

## Phase 1: Project skeleton, dependencies, build and test wiring

**Goal:** A clean checkout can run `npm ci && npm run build && npm test` successfully.
**Addresses:** AC-1, AC-2, AC-3, AC-4, AC-6, AC-7, AC-8, AC-10

**Files:**
- Create: `.gitignore`
- Create: `package.json`
- Create: `tsconfig.json`
- Create: `README.md`
- Create: `src/version.ts`
- Test: `src/version.test.ts`

**Steps:**

1. Write `.gitignore` with exactly these entries, one per line, plus a trailing `npm-debug.log*` and `.DS_Store`:
   ```
   node_modules/
   dist/
   output/
   .env
   .env.*
   *.token-cache.json
   .msal-cache*
   npm-debug.log*
   .DS_Store
   ```
   `.env.*` is included alongside `.env` because `.env.local` is the conventional second file and the repo is public.

2. Write `package.json`:
   ```json
   {
     "name": "onenote-mcp",
     "version": "0.1.0",
     "private": true,
     "type": "module",
     "engines": { "node": ">=24" },
     "scripts": {
       "build": "tsc -p tsconfig.json",
       "typecheck": "tsc -p tsconfig.json --noEmit",
       "dev": "node --watch src/index.ts",
       "start": "node dist/index.js",
       "bootstrap": "node src/bootstrap.ts",
       "test": "node --test \"src/**/*.test.ts\""
     },
     "dependencies": {
       "@azure/msal-node": "^5.5.0",
       "@google-cloud/firestore": "^9.0.0",
       "@modelcontextprotocol/sdk": "^1.30.0",
       "@resvg/resvg-js": "^2.6.2",
       "express": "^5.2.1",
       "fast-xml-parser": "^5.11.0"
     },
     "devDependencies": {
       "@types/express": "^5.0.6",
       "@types/node": "^24.0.0",
       "typescript": "^7.0.2"
     }
   }
   ```
   `start` runs the built output because that is what the Docker runtime stage in #6 executes. `dev` and `bootstrap` run TypeScript source directly via native type stripping.

3. Write `tsconfig.json`:
   ```json
   {
     "compilerOptions": {
       "target": "es2024",
       "lib": ["es2024"],
       "module": "nodenext",
       "moduleResolution": "nodenext",
       "types": ["node"],
       "rootDir": "src",
       "outDir": "dist",
       "strict": true,
       "noUncheckedIndexedAccess": true,
       "exactOptionalPropertyTypes": true,
       "verbatimModuleSyntax": true,
       "erasableSyntaxOnly": true,
       "allowImportingTsExtensions": true,
       "rewriteRelativeImportExtensions": true,
       "declaration": true,
       "sourceMap": true,
       "skipLibCheck": true,
       "noEmitOnError": true
     },
     "include": ["src/**/*.ts"],
     "exclude": ["dist", "node_modules"]
   }
   ```
   `"types": ["node"]` is required explicitly — without it TypeScript 7.0.2 emits `TS2591: Cannot find name 'node:test'` even with `@types/node` installed. `skipLibCheck` is on because `@google-cloud/firestore` ships generated protobuf declarations that are slow and not ours to fix.

4. Run `npm install` to generate `package-lock.json`. Commit the lockfile — AC-8 requires `npm ci`, which fails without one.

5. Create `src/version.ts` exporting `export const SERVICE_NAME = 'onenote-mcp';` and `export const VERSION = '0.1.0';`. This gives Phase 1 a real module to compile and a real thing to test, and `/healthz` in Phase 3 reports both.

6. Write `README.md`: the project name, one sentence saying it is an MCP server exposing Microsoft OneNote via Microsoft Graph, a prominent pointer to `project-spec.md` as the authoritative design document, a Requirements line stating Node >= 24, a Quick start block (`npm ci`, `npm run build`, `npm test`), and a Configuration section listing the env var names from Phase 2 with which are required and which have defaults. Do not put the Entra tenant name or ID in it — `project-spec.md:405` forbids that in a public repo.

**Tests added/updated:**
- `src/version.test.ts` — asserts `SERVICE_NAME === 'onenote-mcp'` and that `VERSION` matches `/^\d+\.\d+\.\d+$/`. Its real job is proving the `node --test` + type-stripping + `./version.ts` import path works end to end before any logic depends on it.

**Verification:**
- [ ] `rm -rf node_modules dist && npm ci && npm run build` exits 0 and produces `dist/version.js` (AC-8)
- [ ] `npm test` exits 0 with 1 passing test (AC-6)
- [ ] `npm run typecheck` exits 0 (lint)
- [ ] `git status --porcelain` shows no `node_modules/` or `dist/` entries after a build (AC-1)
- [ ] `node -e "const p=require('./package.json'); if(p.type!=='module') process.exit(1)"` confirms ESM (AC-2)

---

## Phase 2: `src/config.ts` — grouped env schema and fail-fast validation

**Goal:** One module that reads every environment variable the service uses, validates it, and throws a single named error naming every missing variable at once.
**Addresses:** AC-5

**Files:**
- Create: `src/config.ts`
- Test: `src/config.test.ts`

**Steps:**

1. Define and export `class ConfigError extends Error`. Set `this.name = 'ConfigError'` in the constructor. It takes `missing: string[]` and `invalid: string[]`, stores both as readonly fields, and composes a message that lists each missing variable on its own line with its purpose. Two arrays, not one, because "you did not set `MCP_TOKEN_SIGNING_KEY`" and "`PORT` is not a number" need different remedies.

2. Define the variable table as a module-level constant. Each entry: `name`, `group`, `required` or a `default`, a one-line `purpose` string used in the error message, and an optional `parse` function. The entries:

   | Name | Group | Required | Default | Purpose |
   |---|---|---|---|---|
   | `ONENOTE_CLIENT_ID` | `graph` | yes | — | Azure app registration client ID (public client) |
   | `ONENOTE_AUTHORITY` | `graph` | yes | — | Entra ID authority URL for the tenant |
   | `FIRESTORE_CACHE_DOC` | `firestore` | no | `tokencache/msal` | Firestore document path holding the MSAL token cache |
   | `GOOGLE_CLOUD_PROJECT` | `firestore` | no | — | GCP project; the Firestore client infers it on Cloud Run, so it is only needed locally |
   | `MCP_OAUTH_CLIENT_ID` | `oauth` | yes | — | Layer-1 OAuth client ID that Claude presents |
   | `MCP_OAUTH_CLIENT_SECRET` | `oauth` | yes | — | Layer-1 OAuth client secret |
   | `MCP_TOKEN_SIGNING_KEY` | `oauth` | yes | — | Key used to sign issued access tokens |
   | `PORT` | `server` | no | `8080` | Bind port; Cloud Run sets this |

   Every required name here appears in the `env_vars:` block at `project-spec.md:280`, and nothing required is absent from it.

3. Validation rules beyond presence:
   - `PORT` must parse as an integer in 1–65535. A non-numeric `PORT` goes into `invalid`, not `missing`.
   - `ONENOTE_AUTHORITY` must parse as a `https:` URL via `new URL()`.
   - `FIRESTORE_CACHE_DOC` must have an even number of `/`-separated segments, all non-empty — a Firestore *document* path is collection/doc pairs, and `tokencache` alone would be a collection reference that fails at write time in issue #7.
   - `MCP_TOKEN_SIGNING_KEY` must be at least 32 characters. A short signing key is a real weakness and this is the only place it can be caught cheaply.
   - Values are trimmed; a variable set to whitespace counts as missing, because that is what an empty GitHub secret produces.

4. Export `loadConfig(groups: ConfigGroup[], env: NodeJS.ProcessEnv = process.env): Config`. It walks the table, keeps only entries whose `group` is in `groups`, accumulates *all* missing and invalid names rather than throwing on the first, and throws one `ConfigError` at the end if either array is non-empty. Taking `env` as an injectable parameter defaulting to `process.env` is what makes the tests hermetic — they never mutate the real environment.

5. Export a `Config` type whose shape is grouped (`{ graph?: {...}, firestore?: {...}, oauth?: {...}, server?: {...} }`), with only the requested groups populated. Declare it with `interface`/`type` only — `erasableSyntaxOnly` forbids `enum`, so `ConfigGroup` is `type ConfigGroup = 'graph' | 'firestore' | 'oauth' | 'server'`.

6. `config.ts` must not read `process.env` at module scope and must not call `process.exit`. Only `loadConfig()` reads the environment, and only the entrypoint decides to exit. Otherwise importing the module from a test has side effects.

**Tests added/updated:**
- `src/config.test.ts`, all using an injected fake env object:
  - a complete env for all four groups loads and returns every parsed value
  - an env missing `MCP_OAUTH_CLIENT_ID` and `MCP_TOKEN_SIGNING_KEY` throws one `ConfigError` whose `missing` array contains **both** names — the accumulate-don't-short-circuit behavior AC-5 requires
  - the thrown error has `name === 'ConfigError'` and its message contains every missing name
  - requesting only `['graph','firestore']` succeeds against an env with no `MCP_*` variables set (the issue #9 bootstrap case)
  - `PORT` unset yields `8080`; `PORT='3000'` yields the number `3000`; `PORT='abc'` and `PORT='70000'` each land in `invalid`, not `missing`
  - `FIRESTORE_CACHE_DOC` unset yields `tokencache/msal`; `FIRESTORE_CACHE_DOC='tokencache'` is invalid (odd segment count)
  - `ONENOTE_AUTHORITY='not-a-url'` and `ONENOTE_AUTHORITY='http://x'` are both invalid
  - `MCP_TOKEN_SIGNING_KEY='short'` is invalid
  - a variable set to `'   '` is reported as missing

**Verification:**
- [ ] `npm test` exits 0, all `config.test.ts` cases pass
- [ ] `npm run typecheck` exits 0 (lint)
- [ ] `npm run build` exits 0
- [ ] `grep -n 'process.exit\|process\.env' src/config.ts` shows `process.env` only as the `loadConfig` default parameter value, and `process.exit` only inside `exitOnConfigError` (pulled forward from Phase 3 step 4 so both entrypoints share one implementation)

---

## Phase 3: `src/index.ts` entrypoint, `/healthz`, bootstrap placeholder

**Goal:** `npm start` either binds `process.env.PORT` and serves `/healthz`, or exits 1 with a readable list of missing variables and no stack trace.
**Addresses:** AC-5, AC-9

**Files:**
- Create: `src/index.ts`
- Create: `src/server.ts`
- Create: `src/bootstrap.ts`
- Test: `src/server.test.ts`

**Steps:**

1. `src/server.ts` exports `createApp(config)` returning a configured Express 5 `Application` with one route: `GET /healthz` responding `200` with `{ status: 'ok', service: SERVICE_NAME, version: VERSION }`. It does not listen and does not read the environment. Keeping construction separate from listening is what lets `server.test.ts` exercise the route without binding a port, and it is the seam issue #14 mounts the MCP transport onto.

   `/healthz` returns no configuration values. Issue #6's acceptance criterion is `curl localhost:8080/healthz` returning 200, which this satisfies, and echoing config from an `--allow-unauthenticated` service would leak it.

2. `src/index.ts` is the entrypoint:
   - `import { loadConfig, ConfigError } from './config.ts'` (the `.ts` specifier — `rewriteRelativeImportExtensions` turns it into `./config.js` in `dist/`)
   - call `loadConfig(['graph', 'firestore', 'oauth', 'server'])` inside `try`
   - `catch (err)`: if `err instanceof ConfigError`, write `err.message` to `process.stderr` and `process.exit(1)`. Do not re-throw and do not print `err.stack` — AC-9 is specifically about not showing a stack trace. Any other error is re-thrown, because an unexpected failure should not be disguised as a config problem.
   - on success, `createApp(config).listen(config.server.port, ...)` and log `listening on port N`. The port comes only from `config.server.port`; no numeric literal appears in `index.ts`.
   - register a `SIGTERM` handler calling `server.close()`. Cloud Run sends `SIGTERM` before terminating an instance, and without this the container is killed after the grace period on every revision change.

3. `src/bootstrap.ts` is a placeholder for issue #9: it calls `loadConfig(['graph', 'firestore'])` with the same `ConfigError` handling as `index.ts`, then writes `bootstrap CLI not implemented yet — see issue #9` to stderr and exits 1. It exists so `npm run bootstrap` from AC-2 resolves to a real file and so the grouped loader has a second real consumer.

4. Factor the `ConfigError` catch into a small exported helper in `src/config.ts` (`exitOnConfigError(err): never`) rather than duplicating the block in both entrypoints. It is the one place `process.exit` is allowed to appear in that module, and it is only ever called from an entrypoint.

**Tests added/updated:**
- `src/server.test.ts`:
  - `createApp()` with a stub config, driven through `app.listen(0)` on an ephemeral port, then `fetch('/healthz')` — asserts status 200 and a JSON body whose `service` is `onenote-mcp` and whose `version` matches `/^\d+\.\d+\.\d+$/`. Closes the server in an `after` hook. Port 0 rather than a fixed port so the test cannot collide with a running dev server.
  - `GET /` returns 404 — confirms no accidental catch-all is exposed on a service that will be `--allow-unauthenticated`.
  - the `/healthz` body has no key whose name matches `/secret|key|client_id|token/i` — a regression guard against a later phase adding config echo to the health endpoint.
- `src/config.test.ts` gains one case: `exitOnConfigError` is not unit-tested directly (it calls `process.exit`); instead the entrypoint behavior is verified by the Phase 3 command-line check below.

**Verification:**
- [ ] `npm test` exits 0, all cases pass
- [ ] `npm run typecheck` exits 0 (lint)
- [ ] `npm run build && env -u ONENOTE_CLIENT_ID -u ONENOTE_AUTHORITY -u MCP_OAUTH_CLIENT_ID -u MCP_OAUTH_CLIENT_SECRET -u MCP_TOKEN_SIGNING_KEY npm start` exits non-zero, prints all five missing names, and prints no line matching `` `at .*\(` `` (AC-9)
- [ ] with all required vars set and `PORT=8081`, `npm start` logs `listening on port 8081` and `curl -s -o /dev/null -w '%{http_code}' localhost:8081/healthz` returns `200` (AC-5)
- [ ] with all required vars set and `PORT` unset, the same check against `localhost:8080` returns `200` (default path)
- [ ] `grep -rnE '\b(8080|3000)\b' src/index.ts src/server.ts` returns nothing — the only `8080` in the tree is the documented default inside the `config.ts` table (AC-5, "never a hardcoded port")
- [ ] `npm run bootstrap` with the graph and firestore vars set prints the issue-#9 message and exits 1

---

## Final Verification

- [ ] `rm -rf node_modules dist && npm ci && npm run build` exits 0 from a clean checkout
- [ ] `npm test` exits 0 — all unit tests across `version`, `config`, and `server` pass
- [ ] `npm run typecheck` exits 0 (lint substitute; see Context → Lint)
- [ ] No E2E tests exist yet; the first end-to-end check is issue #6's container run and issue #25's deploy
- [ ] `git status --porcelain` is clean after a full build and test run
- [ ] Acceptance criteria traceability:
  - **AC-1:** covered by Phase 1 — `.gitignore` written with all six required entries plus `.env.*`; verified by `git status --porcelain` being clean after a build
  - **AC-2:** covered by Phase 1 — `package.json` with `"type": "module"`, `engines.node: ">=24"` (deviation, finding 1), and all five required scripts
  - **AC-3:** covered by Phase 1 — `tsconfig.json` with `moduleResolution: nodenext` and `target: es2024` (deviation, finding 1)
  - **AC-4:** covered by Phase 1 — all five named packages plus Express 5 (finding 2); install verified under `--engine-strict`
  - **AC-5:** covered by Phase 2 (the `loadConfig` table, accumulated-error validation, and `ConfigError`) and Phase 3 (port sourced only from `config.server.port`, verified by the `grep` for hardcoded ports)
  - **AC-6:** covered by Phase 1 — `npm test` = `node --test "src/**/*.test.ts"`, running TypeScript source directly (finding 3); every subsequent phase adds tests to it
  - **AC-7:** covered by Phase 1 — `README.md` pointing at `project-spec.md`
  - **AC-8:** covered by Phase 1 — `npm ci && npm run build` from a clean checkout is an explicit verification step, and the lockfile is committed
  - **AC-9:** covered by Phase 3 — the `ConfigError` catch in `index.ts` prints the message and exits 1 without a stack; verified by the `env -u …` check asserting no `at …(` line in the output
  - **AC-10:** covered by Phase 1 — the project is TypeScript throughout; no `.mjs` file is created

## Follow-ups this plan creates

- **Issue #6:** ✅ Done — its Dockerfile AC was edited to `node:24-slim`, and a "Base image constraints" section was added recording that the base must stay Debian `-slim` and must not become Alpine (the resvg prebuild is glibc-only, no musl build exists).
- **Issue #21:** will need to add a public-base-URL env var to both `config.ts` and the deploy workflow's `env_vars:` block in the same change. Deliberately omitted here (Scope excess).
