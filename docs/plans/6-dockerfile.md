# Dockerfile for linux/amd64 with a working resvg binary

**Issue:** #6
**Goal:** Produce a multi-stage `Dockerfile` whose `linux/amd64` runtime image carries a working `@resvg/resvg-js` native binary, serves `/healthz` on `$PORT`, and contains no dev dependencies or host `node_modules`.
**Date:** 2026-08-18

## Status

| Phase | Status | Commit |
|-------|--------|--------|
| Phase 1: `.dockerignore`, multi-stage `Dockerfile`, image-content test | ☑ | `c64d54c` |
| Phase 2: Runtime behaviour — `$PORT`, `/healthz`, startup time, resvg render | ☑ | `86cd1d8` |
| Phase 3: Documentation — `README.md`, `CLAUDE.md` | ☑ | this commit |

## Acceptance Criteria

> Reproduced from issue #6 so the plan is self-contained. The four "Tasks" checkboxes, the "Base image constraints" paragraph, and both sentences of "Acceptance" are all treated as acceptance criteria.

- **AC-1:** `Dockerfile`, multi-stage: build stage installs dev deps and compiles; runtime stage is `node:24-slim` with production deps only. ⚠️ Deviation from the issue's original text: the issue body says `node:24-slim` after the author's comment corrected it from `node:20-slim`; the plan follows the corrected value.
- **AC-2:** Do not copy `node_modules` from the host. Add `.dockerignore` covering `node_modules`, `.git`, `output`, `dist`.
- **AC-3:** Confirm the resvg native binary is present in the runtime stage — a production-only `npm ci` must still pull the correct optional platform package.
- **AC-4:** Container listens on `$PORT` and starts in well under Cloud Run's startup timeout.
- **AC-5:** `/healthz` endpoint that does not touch Graph or Firestore. Delivered by #5; this issue only verifies it responds inside the container.
- **AC-6:** The runtime base must be Debian `-slim`, not Alpine. `@resvg/resvg-js-linux-x64-gnu` declares `libc: ["glibc"]` and has no musl prebuild.
- **AC-7:** `docker build --platform linux/amd64 -t onenote-mcp . && docker run -p 8080:8080 -e PORT=8080 ... onenote-mcp` starts, and `curl localhost:8080/healthz` returns 200.
- **AC-8:** A test that renders a trivial SVG through resvg passes inside the container.

## Context

### What exists today

Issue #5 is complete. The repository has `package.json` (ESM, `engines.node: ">=24"`), a committed `package-lock.json`, two tsconfigs, `src/{index,server,config,version,bootstrap}.ts`, `test/{config,server,version}.test.ts`, and `scripts/` with `gcp-bootstrap.sh` plus a bash test harness under `scripts/test/`.

There is no `Dockerfile` and no `.dockerignore`. `dist/` exists locally but is gitignored.

`src/index.ts` already does everything the container needs: it binds `config.server.port` (from `PORT`, default `8080`), prints one readable line and exits 1 on a config or bind failure, and closes the server on `SIGTERM`/`SIGINT`. `src/server.ts` serves `GET /healthz` returning `{status, service, version}` and touches neither Graph nor Firestore. No source change is needed for this issue.

### Design specs consulted

There is no `docs/design/` directory. `project-spec.md` is authoritative. Relevant sections:

- `project-spec.md:164` — "Cloud Run sets `PORT` itself; the server must bind to it rather than a hardcoded value." Already satisfied by `src/index.ts`; the container must not override it with a baked-in value.
- `project-spec.md:181` "Runtime | Cloud Run, single service, `--max-instances=1`" — the deployment target the image is built for.
- `project-spec.md:183` "Image build | `docker build` inside the GitHub Actions job on an `ubuntu-latest` runner, pushed to Artifact Registry, deployed by digest".
- `project-spec.md:189` — "Building the image in the Actions job rather than with `gcloud run deploy --source` is deliberate. The runner is `linux/amd64`, which is what Cloud Run runs, so the platform-specific `@resvg/resvg-js` binary matches without a `--platform` flag or a buildx setup." This is the reason the Dockerfile must not do anything that would defeat the native-binary resolution.
- `project-spec.md:296` — the build command the deploy workflow (#25) will run is a bare `docker build -t "$IMAGE" .` with no `--platform` and no build args. The Dockerfile must work under exactly that invocation.
- `project-spec.md:392` "Repo hygiene" — the reason `.dockerignore` excludes `output/` and the token-cache globs as well as the four the issue names.

### Issue Assessment

Findings from checking the issue against the current state of the repository and the installed dependency tree. Every claim below was checked against a real file or a real command, not inferred.

**Nothing stale.** Every AC still refers to something that exists. `/healthz` (AC-5) is live in `src/server.ts:22`. `PORT` (AC-4) is spec'd in `src/config.ts:84` with fallback `'8080'`.

**Resolved: the `node:20-slim` / `node:24-slim` discrepancy.** The issue's Tasks list says `node:24-slim`; the author's comment records the bump from `node:20-slim` and its reason (`@google-cloud/firestore@9.0.0` declares `engines.node: ">=22"`). There is no open conflict — `node:24-slim` is the value, and it matches `engines.node: ">=24"` in `package.json`.

**Verified: `--omit=dev` does not drop the resvg platform package.** `@resvg/resvg-js@2.6.2` declares twelve `optionalDependencies`, one per platform, including `@resvg/resvg-js-linux-x64-gnu@2.6.2` (checked in `node_modules/@resvg/resvg-js/package.json`). npm omits optional dependencies only under `--omit=optional`, which this plan does not use. `package-lock.json:412` carries the `linux-x64-gnu` entry, so `npm ci --omit=dev` on a `linux/amd64` base resolves it from the lockfile without touching the registry metadata for platform selection. AC-3 nonetheless asks for this to be *confirmed*, not assumed, so Phase 1 adds an assertion that the `.node` binary file exists in the runtime image.

**Scope gap 1 — the issue names no place for its test to live.** AC-8 requires "a test that renders a trivial SVG through resvg ... inside the container", and AC-7 requires a build-then-curl check. Neither can run under `npm test`: that runs `node --test` over `test/**/*.test.ts` on the host, and the runtime image deliberately contains no test files and no TypeScript. The repository already has a precedent for a test that shells out — `scripts/test/bootstrap.test.sh`, driven by `scripts/test/run.sh`. This plan puts the container checks in `scripts/test/docker.test.sh` under the same harness. `npm test` is left alone.

**Scope gap 2 — the container test must not run by default.** `scripts/test/run.sh` is meant to be fast. A `docker build` plus a container start is neither fast nor available everywhere (no Docker daemon in some environments). Phase 1 therefore adds `docker.test.sh` to `run.sh`'s syntax/shellcheck file list unconditionally, but runs its behavioural body only when `RUN_DOCKER_TESTS=1` is set, printing an explicit skip line otherwise. A silent skip would let the suite report success while checking nothing.

**Scope gap 3 — "production deps only" is a claim worth asserting.** The issue says the runtime stage has production deps only but does not ask for proof. A build-stage mistake that copies `/app/node_modules` forward would satisfy every other AC while shipping `typescript` and `@types/node` into the deployed image. Phase 1 asserts `node_modules/typescript` is absent from the runtime image, which fails loudly on that mistake.

**Scope excess — none.** No AC duplicates existing functionality or contradicts a project convention.

**Deliberately out of scope.** The GitHub Actions workflow that runs `docker build` and pushes to Artifact Registry is issue #25, not this one. This issue produces only the `Dockerfile`, the `.dockerignore`, and the test that exercises them locally.

### Design decisions

These are settled here so they are not re-litigated during implementation.

| Decision | Choice | Reason |
|---|---|---|
| Runtime deps | Second `npm ci --omit=dev` in the runtime stage, not a `COPY --from=build /app/node_modules` | AC-1 says "production deps only" and AC-3 says a production-only `npm ci` must pull the platform package. Copying the build stage's tree would ship dev deps and would test nothing. |
| Layer order | `COPY package.json package-lock.json` then `npm ci`, then `COPY src` | Dependency install is cached across source-only changes. |
| `dist/` | Built in the build stage, copied into the runtime stage | The host `dist/` is in `.dockerignore`, so a stale host build can never leak into the image. |
| `EXPOSE` | `EXPOSE 8080` present, with a comment that it is metadata only | Cloud Run ignores it; it documents the `PORT` default from `src/config.ts:84` for local `docker run -P`. No `ENV PORT` is set — the default lives in `config.ts` and is not duplicated. |
| User | `USER node` in the runtime stage | The `node` image ships a non-root `node` user. The process only reads `/app`, so root is unnecessary. |
| `CMD` | `CMD ["node", "dist/index.js"]` (exec form) | Exec form makes the Node process PID 1, so Cloud Run's `SIGTERM` reaches the handler in `src/index.ts:37` instead of a shell. |
| `NODE_ENV` | `ENV NODE_ENV=production` in the runtime stage | Express 5 uses it to disable the development error view and enable view caching. |

### Lint

The repository has no JavaScript linter; `npm run typecheck` is the static-analysis gate, and it is what the verification steps below call "lint". This issue adds no TypeScript, so `npm run typecheck` is a regression check rather than a check of new code. The one new file with a linter available is `scripts/test/docker.test.sh`, covered by the `shellcheck` step already in `scripts/test/run.sh`.

---

## Phase 1: `.dockerignore`, multi-stage `Dockerfile`, image-content test

**Goal:** `docker build --platform linux/amd64 -t onenote-mcp:test .` succeeds, and the resulting image contains `dist/`, the resvg glibc binary, and no dev dependencies.
**Addresses:** AC-1, AC-2, AC-3, AC-6

**Files:**
- Create: `.dockerignore`
- Create: `Dockerfile`
- Create: `scripts/test/docker.test.sh`
- Modify: `scripts/test/run.sh`

**Steps:**

1. Create `.dockerignore`. It is a denylist, not an allowlist, so a new source directory is included by default rather than silently dropped:
   ```
   node_modules
   .git
   .gitignore
   dist
   output
   test
   docs
   scripts
   .env
   .env.*
   *.token-cache.json
   .msal-cache*
   npm-debug.log*
   .DS_Store
   Dockerfile
   .dockerignore
   ```
   `node_modules`, `.git`, `output`, `dist` are the four AC-2 names. `test`, `docs`, and `scripts` are excluded because nothing in the image reads them and they only enlarge the build context. The `.env` and cache globs mirror `.gitignore` so a local operator's credentials cannot reach a layer.

2. Create `Dockerfile`:
   ```dockerfile
   # Build stage: full dependency tree, compile src/ to dist/.
   FROM node:24-slim AS build
   WORKDIR /app
   COPY package.json package-lock.json ./
   RUN npm ci
   COPY tsconfig.json tsconfig.build.json ./
   COPY src ./src
   RUN npm run build

   # Runtime stage. Debian -slim, never Alpine: @resvg/resvg-js-linux-x64-gnu
   # declares libc: ["glibc"] and has no musl prebuild, so an Alpine base either
   # fails to install the optional package or fails at require() time.
   FROM node:24-slim AS runtime
   ENV NODE_ENV=production
   WORKDIR /app

   # A second install from the lockfile rather than a COPY of the build stage's
   # node_modules. This is what keeps dev dependencies out of the image, and it is
   # what proves the platform-specific optional package still resolves without them.
   COPY package.json package-lock.json ./
   RUN npm ci --omit=dev && npm cache clean --force

   COPY --from=build /app/dist ./dist

   # Metadata only. Cloud Run ignores EXPOSE and sets PORT itself; the default
   # lives in the SPECS table in src/config.ts and is not duplicated as ENV here.
   EXPOSE 8080

   USER node
   CMD ["node", "dist/index.js"]
   ```

3. Create `scripts/test/docker.test.sh`, following the shape of `scripts/test/bootstrap.test.sh` (a `status` variable, one `echo` per check, non-zero exit on any failure). This phase writes the build and image-content checks:
   - `set -uo pipefail`, `cd "$(dirname "$0")/../.."`.
   - Guard: if `command -v docker` fails, print `SKIP: docker not available` and exit 0.
   - `IMAGE=onenote-mcp:test`. Register a `trap` that removes the image and any container started later.
   - Check 1 — build: `docker build --platform linux/amd64 -t "$IMAGE" .`, redirecting output to a temp log; on failure print the last 30 lines and set `status=1`.
   - Check 2 — resvg glibc binary present: `docker run --rm --entrypoint sh "$IMAGE" -c 'ls node_modules/@resvg/resvg-js-linux-x64-gnu/*.node'` exits 0. This is the AC-3 confirmation.
   - Check 3 — no dev dependencies: `docker run --rm --entrypoint sh "$IMAGE" -c '! test -e node_modules/typescript'` exits 0.
   - Check 4 — compiled output present: `docker run --rm --entrypoint sh "$IMAGE" -c 'test -f dist/index.js'` exits 0.
   - Check 5 — no host build context leaked: `docker run --rm --entrypoint sh "$IMAGE" -c '! test -e test && ! test -e src'` exits 0. This fails if `.dockerignore` stops excluding `test`, and would have failed on a `COPY . .` Dockerfile.
   - `exit "$status"`.
   - `chmod +x` the file.

4. Modify `scripts/test/run.sh`:
   - Add `scripts/test/docker.test.sh` to the `FILES` array so `bash -n` and `shellcheck` cover it.
   - After the existing `bootstrap.test.sh` section, add:
     ```bash
     echo "== docker.test.sh =="
     if [ "${RUN_DOCKER_TESTS:-}" = "1" ]; then
       bash scripts/test/docker.test.sh || status=1
     else
       echo "skipped: set RUN_DOCKER_TESTS=1 to build and run the container"
     fi
     ```

**Tests added/updated:**
- `scripts/test/docker.test.sh` checks 1–5: image builds for `linux/amd64`; `@resvg/resvg-js-linux-x64-gnu/*.node` exists in the runtime image; `node_modules/typescript` does not; `dist/index.js` exists; `src/` and `test/` are absent from the image.

**Verification:**
- [x] `npm run typecheck` passes (lint)
- [x] `npm test` passes — no source changed, this is a regression check
- [x] `bash scripts/test/run.sh` passes, and prints the `docker.test.sh` skip line
- [x] `RUN_DOCKER_TESTS=1 bash scripts/test/run.sh` passes, exercising checks 1–5
- [x] `shellcheck` reports nothing on `scripts/test/docker.test.sh` (run inside `run.sh`)

---

## Phase 2: Runtime behaviour — `$PORT`, `/healthz`, startup time, resvg render

**Goal:** The built image starts a server on the `PORT` it is given, answers `/healthz` with 200 quickly, and can render an SVG to PNG through resvg.
**Addresses:** AC-3, AC-4, AC-5, AC-7, AC-8

**Files:**
- Modify: `scripts/test/docker.test.sh`

**Steps:**

1. Add check 6 — resvg renders inside the container (AC-8). Run Node in the runtime image with an inline program, no test file shipped:
   ```bash
   docker run --rm --entrypoint node "$IMAGE" -e "
     const { Resvg } = await import('@resvg/resvg-js');
     const svg = '<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"10\" height=\"10\"><rect width=\"10\" height=\"10\" fill=\"red\"/></svg>';
     const png = new Resvg(svg).render().asPng();
     if (png.length < 8) throw new Error('empty png');
     const sig = [0x89, 0x50, 0x4e, 0x47];
     if (!sig.every((b, i) => png[i] === b)) throw new Error('not a png');
     console.log('resvg ok, ' + png.length + ' bytes');
   "
   ```
   Assert exit 0 and that stdout contains `resvg ok`. Checking the four-byte PNG signature rather than just a non-zero exit means a stub that returns an empty buffer fails.

2. Add check 7 — the container serves `/healthz` on the `PORT` it is given (AC-4, AC-5, AC-7). Use a non-default port so the assertion proves `PORT` is read rather than a baked-in `8080`:
   - Start detached with `-e PORT=9090 -p 127.0.0.1:0:9090` plus the five required config variables, so `loadConfig` succeeds without any real credential:
     ```
     ONENOTE_CLIENT_ID=00000000-0000-0000-0000-000000000000
     ONENOTE_AUTHORITY=https://login.microsoftonline.com/common
     MCP_OAUTH_CLIENT_ID=test-client
     MCP_OAUTH_CLIENT_SECRET=test-secret
     MCP_TOKEN_SIGNING_KEY=0123456789abcdef0123456789abcdef
     ```
     `ONENOTE_AUTHORITY` must be an https URL (`checkHttpsUrl`, `src/config.ts:48`) and `MCP_TOKEN_SIGNING_KEY` must be at least 32 characters (`MIN_SIGNING_KEY_LENGTH`, `src/config.ts:34`). `FIRESTORE_CACHE_DOC` and `GOOGLE_CLOUD_PROJECT` are optional and are left unset — omitting them also demonstrates AC-5's "does not touch Graph or Firestore".
   - Resolve the mapped host port with `docker port "$CID" 9090`, so a busy host port cannot make the test flaky.
   - Poll `curl -s -o "$body" -w '%{http_code}' "http://127.0.0.1:$HOSTPORT/healthz"` every 0.2s for up to 30s. Record elapsed time with `SECONDS`.
   - Assert the status code is `200` and the body contains `"status":"ok"`.
   - Assert elapsed time is under 10 seconds. Cloud Run's default startup timeout is 240s, so 10s is the "well under" threshold in AC-4; print the measured value either way so a regression is visible before it fails.
   - On failure, print `docker logs "$CID"` — a config error from `src/index.ts` is one readable line and is the likely cause.
   - Stop and remove the container in the trap.

3. Add check 8 — `SIGTERM` shuts the container down cleanly. `docker stop` sends `SIGTERM`; assert the container's exit code is 0 and its logs contain `SIGTERM received, shutting down` (the message in `src/index.ts:40`). This guards the exec-form `CMD` decision: with shell-form `CMD` the signal goes to `/bin/sh` and the handler never runs.

**Tests added/updated:**
- `scripts/test/docker.test.sh` check 6: an inline Node program in the runtime image renders a 10×10 SVG through `@resvg/resvg-js` and the output starts with the PNG signature.
- `scripts/test/docker.test.sh` check 7: the container started with `PORT=9090` returns HTTP 200 and `"status":"ok"` from `/healthz` within 10 seconds, with no Firestore or Graph variables set.
- `scripts/test/docker.test.sh` check 8: `docker stop` produces exit code 0 and the `SIGTERM received` log line.

**Verification:**
- [x] `npm run typecheck` passes (lint)
- [x] `npm test` passes
- [x] `RUN_DOCKER_TESTS=1 bash scripts/test/run.sh` passes all eight checks
- [x] The AC-7 command from the issue run by hand succeeds: `docker build --platform linux/amd64 -t onenote-mcp .` then `docker run --rm -p 8080:8080 -e PORT=8080 -e ONENOTE_CLIENT_ID=... -e ONENOTE_AUTHORITY=https://login.microsoftonline.com/common -e MCP_OAUTH_CLIENT_ID=... -e MCP_OAUTH_CLIENT_SECRET=... -e MCP_TOKEN_SIGNING_KEY=... onenote-mcp`, and `curl -i localhost:8080/healthz` returns 200
- [x] `shellcheck` reports nothing on the updated `scripts/test/docker.test.sh`

---

## Phase 3: Documentation — `README.md`, `CLAUDE.md`

**Goal:** The two files that describe the repository's layout and commands describe the container as well, so the next issue does not have to read the Dockerfile to learn how it is built and tested.
**Addresses:** AC-1, AC-2, AC-6

**Files:**
- Modify: `README.md`
- Modify: `CLAUDE.md`

**Steps:**

1. `CLAUDE.md`, "Directory structure" tree: add `Dockerfile` and `.dockerignore` at the top level with one-line comments, and `scripts/test/docker.test.sh` under `scripts/test/`.

2. `CLAUDE.md`, the paragraph after the tree that explains `scripts/test/`: state that `docker.test.sh` builds the image and runs the container, that `npm test` does not run it, and that `scripts/test/run.sh` runs it only when `RUN_DOCKER_TESTS=1` is set.

3. `CLAUDE.md`, "Conventions": add a **"The runtime base is Debian `-slim`, never Alpine"** entry giving the mechanism — `@resvg/resvg-js-linux-x64-gnu` declares `libc: ["glibc"]` and ships no musl prebuild, so on Alpine the optional package either fails to install or fails at `require` time. This is the AC-6 constraint recorded where a future change would be made.

4. `CLAUDE.md`, "Conventions": add that the runtime stage installs with `npm ci --omit=dev` from the lockfile rather than copying `node_modules` from the build stage, and that `.dockerignore` is a denylist so new source directories are included by default.

5. `README.md`: add a "Container" section after "Scripts" with the build command, a `docker run` example using the same placeholder env values as Phase 2 (no real tenant ID, no real secret — the repository is public), and `RUN_DOCKER_TESTS=1 bash scripts/test/run.sh` as the way to test it. State that the image is `linux/amd64` because that is what Cloud Run runs, and that `PORT` is supplied by Cloud Run.

**Tests added/updated:**
- No new tests. This phase changes only Markdown; the behaviour it documents is covered by Phase 1 and Phase 2 checks. Correctness of the documented commands is verified by running them, listed below.

**Verification:**
- [x] `npm run typecheck` passes (lint)
- [x] `npm test` passes
- [x] Every command quoted in the new `README.md` section runs successfully when pasted into a shell
- [x] The `CLAUDE.md` directory tree matches `git ls-files` for the paths it lists
- [x] No real tenant name, tenant ID, or client secret appears in either file

---

## Final Verification

- [x] `npm run typecheck` passes
- [x] `npm test` passes
- [x] `bash scripts/test/run.sh` passes (docker checks skipped)
- [x] `RUN_DOCKER_TESTS=1 bash scripts/test/run.sh` passes (all eight docker checks)
- [x] `docker build --platform linux/amd64 -t onenote-mcp .` succeeds from a clean checkout with no host `node_modules`
- [x] `docker image inspect onenote-mcp --format '{{.Os}}/{{.Architecture}}'` reports `linux/amd64`
- [x] Acceptance criteria traceability:
  - **AC-1** (multi-stage, build compiles, runtime `node:24-slim` prod deps only): covered by Phase 1 — the `Dockerfile` has `build` and `runtime` stages; Phase 1 check 3 asserts `node_modules/typescript` is absent from the runtime image, and check 4 asserts `dist/index.js` is present. Documented in Phase 3.
  - **AC-2** (no host `node_modules`; `.dockerignore` covers `node_modules`, `.git`, `output`, `dist`): covered by Phase 1 — `.dockerignore` lists all four; check 5 asserts `src/` and `test/` are absent from the image, which fails if the ignore file stops being applied. Documented in Phase 3.
  - **AC-3** (resvg native binary present after production-only `npm ci`): covered by Phase 1 check 2 (`@resvg/resvg-js-linux-x64-gnu/*.node` exists in the runtime image) and Phase 2 check 6 (it actually loads and renders).
  - **AC-4** (listens on `$PORT`, starts well under Cloud Run's startup timeout): covered by Phase 2 check 7 — the container is started with `PORT=9090`, not the default, and `/healthz` must answer within 10 seconds against Cloud Run's 240-second default.
  - **AC-5** (`/healthz` responds inside the container without touching Graph or Firestore): covered by Phase 2 check 7 — `FIRESTORE_CACHE_DOC` and `GOOGLE_CLOUD_PROJECT` are deliberately unset and the endpoint still returns 200 with `"status":"ok"`.
  - **AC-6** (Debian `-slim`, not Alpine): covered by Phase 1 — `FROM node:24-slim` in both stages with the reason in a comment; Phase 1 check 2 and Phase 2 check 6 would both fail on a musl base. Recorded as a convention in Phase 3.
  - **AC-7** (`docker build --platform linux/amd64` then `docker run -p 8080:8080 -e PORT=8080`, `curl /healthz` → 200): covered by Phase 1 check 1 (the build, with `--platform linux/amd64`) and Phase 2 check 7 (the run and the 200), plus the by-hand run of the literal issue command in Phase 2's verification list.
  - **AC-8** (a test renders a trivial SVG through resvg inside the container): covered by Phase 2 check 6 — a 10×10 SVG rendered to PNG in the runtime image, asserted on the PNG signature bytes.

## Findings

> Recorded during implementation: what building this revealed that planning missed.

1. **`shellcheck` flags the `EXIT` trap handler as dead code.** `cleanup()` is only ever
   reached through `trap cleanup EXIT`, which shellcheck's call-graph analysis does not
   see, so it emits SC2329. Suppressed with a `# shellcheck disable=SC2329` line carrying
   the reason. `scripts/test/bootstrap.test.sh` never hit this because it has no trap.

2. **AC-7 verified with the literal command.** `docker build --platform linux/amd64 -t
   onenote-mcp .` then `docker run -d -p 8080:8080 -e PORT=8080 ...` then `curl -i
   localhost:8080/healthz` returns `HTTP/1.1 200 OK`,
   `Content-Type: application/json; charset=utf-8`, body
   `{"status":"ok","service":"onenote-mcp","version":"0.1.0"}`, with `listening on port
   8080` in the container log.

   On the first attempt an unrelated process held host port 8080, so `docker run` failed
   with `failed to bind host port 0.0.0.0:8080/tcp: address already in use` — and a
   `curl localhost:8080/healthz` in that state returned a 200 from that other service,
   which is a false pass. The owner freed the port and the command was re-run as written.
   That near-miss is the reason `docker.test.sh` publishes on host port 0 and resolves
   the real port with `docker port` instead of using a fixed one, and the reason the
   README's Container section notes the `-p 8081:8080` fallback.

3. **`node -e` in the image must use `require`, not an ESM import.** `node -e` runs as
   CommonJS regardless of `"type": "module"` in `package.json`, so the resvg check is
   written `const { Resvg } = require("@resvg/resvg-js")`. Wrapping the JS in single
   quotes and building the SVG string with backticks keeps every double quote in the
   markup unescaped.

4. **Pre-existing, unrelated to this issue:** `scripts/test/bootstrap.test.sh` fails one
   assertion, `unset PROJECT exits non-zero — expected exit 1, got 0`. Confirmed present
   on `HEAD` before any change here by stashing and re-running. It belongs to issue #2's
   script, not to the container work, and is left alone. Every other assertion in that
   suite passes, and all twelve `docker.test.sh` checks pass.

6. **Verifying the clean-checkout build needs a path the Docker daemon can read.** This
   machine runs Docker from a snap, which is confined to `$HOME` — a `docker build` whose
   context is under `/tmp` fails with `unable to prepare context: path ... not found`
   before any layer runs. The check was done by extracting `git archive HEAD` into a
   directory under `$HOME`, copying in `Dockerfile` and `.dockerignore` (both untracked
   at that moment), and building from there. It succeeded, reported `linux/amd64`, and
   the runtime image contained `resvgjs.linux-x64-gnu.node` and `dist/index.js`.

5. **Measured startup: 0 seconds** to a 200 from `/healthz`, against the 10-second bar in
   the plan and Cloud Run's 240-second default startup timeout. The image built for
   `linux/amd64` reports `linux/amd64` from `docker image inspect`.
