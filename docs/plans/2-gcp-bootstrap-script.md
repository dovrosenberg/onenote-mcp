# Author scripts/gcp-bootstrap.sh

**Issue:** #2
**Goal:** Produce an idempotent, reviewable Bash script that provisions every GCP resource the deploy pipeline needs, and prints the exact `gh` commands for issue #4.
**Date:** 2026-08-18

## Status

| Phase | Status | Commit |
|-------|--------|--------|
| Phase 1: Script scaffold, parameters, preflight, test harness | ☑ | `65c07ba` |
| Phase 2: APIs, Artifact Registry, Firestore | ☑ | `03d41a7` |
| Phase 3: Service accounts and IAM role grants | ☑ | `d6b9c8b` |
| Phase 4: Workload Identity Pool, OIDC provider, repo binding | ☑ | `8bb7eac` |
| Phase 5: Output block, no-key guarantee, full re-run test | ☑ | `0c54a2e` |

## Acceptance Criteria

> Reproduced from issue #2 so the plan is self-contained. The unchecked boxes under "What it must do" are treated as acceptance criteria alongside the stated "Acceptance" line.

- **AC-1:** Enable APIs: `run`, `artifactregistry`, `firestore`, `iamcredentials`, `sts`, `cloudresourcemanager`.
- **AC-2:** Create Artifact Registry Docker repo `onenote-mcp` in `$GAR_REGION`.
- **AC-3:** Create a Firestore database in Native mode: `gcloud firestore databases create --location=$REGION --type=firestore-native`.
- **AC-4:** Create runtime service account `onenote-mcp-run@` and grant `roles/datastore.user`.
- **AC-5:** Create deploy service account `onenote-mcp-deploy@` and grant `roles/run.admin`, `roles/artifactregistry.writer`, and `roles/iam.serviceAccountUser` on the runtime SA.
- **AC-6:** Create a Workload Identity Pool and an OIDC provider for `https://token.actions.githubusercontent.com`.
- **AC-7:** Bind the provider to this repository only, via an attribute condition on `assertion.repository == "dovrosenberg/onenote-mcp"`.
- **AC-8:** Grant `roles/iam.workloadIdentityUser` to the principalSet for the repo on the deploy SA.
- **AC-9:** Print, at the end, the exact `gh variable set` / `gh secret set` commands for #4, with the WIF provider resource name and SA emails filled in.
- **AC-10:** No service-account JSON key is created. If the script emits a key file, it is wrong.
- **AC-11:** The script is re-runnable without error on an already-configured project.
- **AC-12:** The script is parameterised by env vars at the top: `PROJECT`, `REGION`, `GAR_REGION`, `SERVICE`, `GITHUB_REPO`.

## Context

### What exists today

The repository contains `LICENSE` and `project-spec.md` and nothing else. There is no `scripts/`, no `docs/`, no `package.json`, no lint or test configuration. Issue #5 (Phase 1 of the project) creates the TypeScript skeleton; this issue lands before it. That means this plan cannot reuse an existing test runner or lint config and must bring its own, scoped to shell.

`git remote -v` confirms the repository is `https://github.com/dovrosenberg/onenote-mcp.git`, so the `assertion.repository` value in AC-7 (`dovrosenberg/onenote-mcp`) matches the real remote.

### Design specs consulted

There is no `docs/design/` directory. The authoritative design document is `project-spec.md`. The relevant sections:

- `project-spec.md:166` "Deployment target: Google Cloud Run via GitHub Actions" — establishes Workload Identity Federation with no JSON key, image push to Artifact Registry, Firestore Native for the MSAL cache, Secret Manager explicitly not used.
- `project-spec.md:197` "Where each piece of configuration lives" — the table that fixes which values become GitHub repo *variables* versus *secrets*. `WIF_PROVIDER` and `DEPLOY_SA` are variables, not secrets, because they are identifiers rather than credentials. This determines the shape of the AC-9 output block.
- `project-spec.md:238` — gives the two Firestore commands verbatim, matching AC-3 and AC-4.
- `project-spec.md:284` "Workflow shape" — shows the consuming workflow reading `vars.WIF_PROVIDER`, `vars.DEPLOY_SA`, `vars.GAR_REGION`, `vars.GCP_PROJECT`, `vars.GCP_REGION`. The names printed by AC-9 must match these exactly or the workflow in #24 breaks.

Issue #4 lists the full set of variables the output block must cover: `GCP_PROJECT`, `GCP_REGION`, `GAR_REGION`, `WIF_PROVIDER`, `DEPLOY_SA`, `RUNTIME_SA`, `ONENOTE_CLIENT_ID`, `ONENOTE_AUTHORITY`, `MCP_OAUTH_CLIENT_ID`, plus secrets `MCP_OAUTH_CLIENT_SECRET` and `MCP_TOKEN_SIGNING_KEY`.

### Issue Assessment

Findings from checking the issue against the current codebase and against how the `gcloud` commands actually behave.

**Scope gaps — added to the plan beyond what the issue lists:**

1. **Attribute mapping is required, not optional.** AC-7 specifies the attribute *condition* on `assertion.repository` but never mentions the attribute *mapping*. `gcloud iam workload-identity-pools providers create-oidc` rejects a condition referencing `assertion.repository` unless `attribute.repository=assertion.repository` is in `--attribute-mapping`, and the AC-8 principalSet path is `.../attribute.repository/<repo>`, which only resolves if that mapping exists. The script must set `--attribute-mapping="google.subject=assertion.sub,attribute.repository=assertion.repository"`. Phase 4 covers this.

2. **The project number, not the project ID, appears in WIF resource names.** Both the provider resource name for AC-9 and the principalSet for AC-8 are built from the numeric project number. The script must resolve it with `gcloud projects describe`. Phase 4 covers this.

3. **Billing preflight.** Issue #3 asks the operator to confirm billing manually after running the script. Artifact Registry and Cloud Run both fail without it, and the failure message from `gcloud services enable` is not obvious. The script checks `gcloud beta billing projects describe` and warns without aborting — a warning rather than a hard failure because the `beta billing` surface requires an extra permission the operator may not have granted themselves. Phase 1 covers this.

4. **Idempotency is not free for any of these commands.** `gcloud services enable` is genuinely idempotent. Every other command in this script is not: `artifacts repositories create`, `firestore databases create`, `iam service-accounts create`, `workload-identity-pools create`, and `providers create-oidc` all exit non-zero with `ALREADY_EXISTS` on a second run. AC-11 therefore requires an explicit describe-then-create guard on each. `add-iam-policy-binding` is idempotent and needs no guard. Phase 1 builds the guard helper; Phases 2–4 use it.

5. **Existing WIF providers need their condition re-applied.** If the provider already exists but was created with a wrong or missing attribute condition, a describe-then-skip guard leaves the security hole in place. On the already-exists path the script runs `providers update-oidc` to re-assert the mapping and condition. This is the one place where "skip if exists" is the wrong idempotency strategy. Phase 4 covers this.

**Scope excess — none.** Every AC maps to a resource the deploy workflow in #24 actually consumes.

**Nothing stale.** No AC references a file, module, or resource that has since been deleted or renamed, because no implementation code exists yet.

### Testing approach

There is no test infrastructure in the repo and this is a shell script that mutates a cloud account, so it cannot be tested by running it. The plan uses a stub-binary harness:

- `scripts/test/stubs/gcloud` — an executable shell stub, placed first on `PATH`, that appends its full argument list to `$GCLOUD_LOG` and exits 0. A `$STUB_MISSING` variable makes selected `describe` calls exit 1, which simulates a fresh project; leaving it unset makes every `describe` succeed, which simulates an already-configured project.
- `scripts/test/bootstrap.test.sh` — a plain Bash test runner with `assert_logged` / `assert_not_logged` helpers that grep `$GCLOUD_LOG`. It runs the bootstrap script twice, once in each stub mode, and asserts on the recorded command lines.

This tests exactly what the issue asks about: which `gcloud` commands are issued, with which flags, and whether the second run is clean. It does not test that GCP accepts them — that is issue #3, which is a manual issue by design.

Lint is `shellcheck` plus `bash -n`. `shellcheck` is invoked via `shellcheck` if on `PATH`, otherwise skipped with a printed notice, so the suite runs on a machine without it.

### Naming decisions

Fixed in Phase 1 and used throughout:

| Thing | Value |
|---|---|
| Runtime SA id | `onenote-mcp-run` |
| Deploy SA id | `onenote-mcp-deploy` |
| Artifact Registry repo | `onenote-mcp` |
| WIF pool id | `github-pool` |
| WIF provider id | `github-provider` |
| WIF pool location | `global` |

---

## Phase 1: Script scaffold, parameters, preflight, test harness

**Goal:** A runnable script that validates its parameters and environment and does nothing else, plus the test harness every later phase asserts through.
**Addresses:** AC-12

**Files:**
- Create: `scripts/gcp-bootstrap.sh`
- Create: `scripts/test/stubs/gcloud`
- Create: `scripts/test/bootstrap.test.sh`
- Create: `scripts/test/run.sh`

**Steps:**

1. Create `scripts/gcp-bootstrap.sh` with `#!/usr/bin/env bash` and `set -euo pipefail`.

2. Add the parameter block at the top of the file, immediately after the shebang and `set` line, using `${VAR:-default}` so each is overridable from the environment without editing the file:
   ```bash
   PROJECT="${PROJECT:-}"
   REGION="${REGION:-us-central1}"
   GAR_REGION="${GAR_REGION:-us-central1}"
   SERVICE="${SERVICE:-onenote-mcp}"
   GITHUB_REPO="${GITHUB_REPO:-dovrosenberg/onenote-mcp}"
   ```
   Add a comment block above it explaining that `REGION` must be a valid Firestore location and that `PROJECT` has no default on purpose.

3. Add derived read-only names below the parameter block:
   ```bash
   RUNTIME_SA_ID="${SERVICE}-run"
   DEPLOY_SA_ID="${SERVICE}-deploy"
   GAR_REPO="${SERVICE}"
   POOL_ID="github-pool"
   PROVIDER_ID="github-provider"
   ```
   Derive `RUNTIME_SA` and `DEPLOY_SA` email addresses as `${..._SA_ID}@${PROJECT}.iam.gserviceaccount.com` after the `PROJECT` check in step 5.

4. Add three helper functions:
   - `log()` — prints `==> $*` to stdout.
   - `warn()` — prints `WARNING: $*` to stderr.
   - `die()` — prints `ERROR: $*` to stderr and exits 1.

5. Add the preflight block:
   - `command -v gcloud >/dev/null || die "gcloud not found on PATH"`
   - `[[ -n "$PROJECT" ]] || die "Set PROJECT (env var or edit the top of this script)"`
   - Verify the project is reachable: `gcloud projects describe "$PROJECT" >/dev/null 2>&1 || die "Cannot describe project $PROJECT. Run: gcloud auth login"`
   - Billing check, non-fatal: run `gcloud beta billing projects describe "$PROJECT" --format='value(billingEnabled)'`; if the command fails, `warn` that billing could not be checked; if it returns anything other than `True`, `warn` that Artifact Registry and Cloud Run will fail without billing.
   - Echo the resolved parameters so the operator sees what is about to happen.

6. Add the idempotency helper that every later phase uses:
   ```bash
   # ensure_resource <human name> <describe command...> -- <create command...>
   ```
   It runs the describe command with output suppressed. On success it logs `<name> already exists, skipping create` and returns 0. On failure it logs `creating <name>` and runs the create command. Implement the `--` split by iterating `"$@"` into two arrays.

7. Create `scripts/test/stubs/gcloud`, executable (`chmod +x`):
   - Append `"$*"` as one line to `"$GCLOUD_LOG"`.
   - If the first two arguments name a describe-style lookup (`describe`, or `list`) and the joined arguments match any pattern in the space-separated `$STUB_MISSING` list, exit 1.
   - Special-case `projects describe "$PROJECT" --format=value(projectNumber)`: echo a fixed number `123456789012` so the WIF resource names in later phases are deterministic.
   - Special-case `beta billing projects describe`: echo `True`.
   - Otherwise exit 0.

8. Create `scripts/test/bootstrap.test.sh` with the runner scaffolding:
   - `set -uo pipefail` (not `-e`; the runner must survive a failing assertion to report all of them).
   - `FAILURES=0`; `assert_logged <pattern> <description>` greps `-F` for the pattern in `$GCLOUD_LOG` and increments `FAILURES` with a printed `FAIL: <description>` on no match; `assert_not_logged` is the inverse.
   - `run_bootstrap()` sets `GCLOUD_LOG` to a fresh temp file under `$TMPDIR`, prepends `scripts/test/stubs` to `PATH`, exports `PROJECT=test-project`, and runs `scripts/gcp-bootstrap.sh`, capturing exit status into `BOOTSTRAP_STATUS` and stdout into `BOOTSTRAP_OUT`.
   - Exit with `1` if `FAILURES` is non-zero, printing the count.

9. Add the Phase 1 test cases to `bootstrap.test.sh`:
   - Running with `PROJECT` unset exits non-zero and prints an error naming `PROJECT`.
   - Running with `PROJECT=test-project` exits 0.
   - Running with `PROJECT=test-project` and `REGION=europe-west1` echoes `europe-west1` in its parameter summary, proving the env-var override path works.

10. Create `scripts/test/run.sh`: runs `bash -n` on `scripts/gcp-bootstrap.sh` and `scripts/test/bootstrap.test.sh`, then `shellcheck` on both if `command -v shellcheck` succeeds (printing `shellcheck not installed, skipping` otherwise), then executes `scripts/test/bootstrap.test.sh`. Exit non-zero if any step fails.

11. `chmod +x` all four created files.

**Tests added/updated:**
- `PROJECT` unset causes a non-zero exit with an error message naming `PROJECT`.
- `PROJECT` set causes a zero exit.
- `REGION` override is reflected in the printed parameter summary.

**Verification:**
- [ ] `bash scripts/test/run.sh` exits 0
- [ ] `shellcheck scripts/gcp-bootstrap.sh` reports no warnings (or prints the skip notice)
- [ ] `bash -n scripts/gcp-bootstrap.sh` exits 0
- [ ] `git status` shows no `.json` file created anywhere under `scripts/`

---

## Phase 2: APIs, Artifact Registry, Firestore

**Goal:** Provision the project-level services and the two stateful resources.
**Addresses:** AC-1, AC-2, AC-3

**Files:**
- Modify: `scripts/gcp-bootstrap.sh`
- Modify: `scripts/test/bootstrap.test.sh`

**Steps:**

1. Add the API enablement block after preflight. Enable all six in a single call so the operator waits once:
   ```bash
   gcloud services enable \
     run.googleapis.com \
     artifactregistry.googleapis.com \
     firestore.googleapis.com \
     iamcredentials.googleapis.com \
     sts.googleapis.com \
     cloudresourcemanager.googleapis.com \
     --project="$PROJECT"
   ```
   No `ensure_resource` guard — `services enable` is idempotent and succeeds on already-enabled services.

2. Add the Artifact Registry block using `ensure_resource`:
   - describe: `gcloud artifacts repositories describe "$GAR_REPO" --location="$GAR_REGION" --project="$PROJECT"`
   - create: `gcloud artifacts repositories create "$GAR_REPO" --repository-format=docker --location="$GAR_REGION" --project="$PROJECT" --description="Container images for $SERVICE"`

3. Add the Firestore block using `ensure_resource`:
   - describe: `gcloud firestore databases describe --database='(default)' --project="$PROJECT"`
   - create: `gcloud firestore databases create --location="$REGION" --type=firestore-native --project="$PROJECT"`
   Add a comment above it noting that `--location` must be a Firestore location name and that the database cannot be moved after creation.

**Tests added/updated:**
- Fresh-project run (`STUB_MISSING` covering the artifacts and firestore describes) logs `services enable` with all six API names, logs `artifacts repositories create` with `--repository-format=docker` and the `GAR_REGION` value, and logs `firestore databases create` with `--type=firestore-native`.
- Already-configured run (`STUB_MISSING` empty) still logs `services enable`, and logs neither `artifacts repositories create` nor `firestore databases create`.
- `GAR_REGION=europe-west4` override appears in the `artifacts repositories create` line, and `REGION` remains independent of it in the `firestore databases create` line.

**Verification:**
- [ ] `bash scripts/test/run.sh` exits 0
- [ ] Both fresh and already-configured stub runs exit 0
- [ ] `shellcheck scripts/gcp-bootstrap.sh` clean

---

## Phase 3: Service accounts and IAM role grants

**Goal:** Create both service accounts and attach every role, including the SA-scoped `serviceAccountUser` binding.
**Addresses:** AC-4, AC-5

**Files:**
- Modify: `scripts/gcp-bootstrap.sh`
- Modify: `scripts/test/bootstrap.test.sh`

**Steps:**

1. Create the runtime service account with `ensure_resource`:
   - describe: `gcloud iam service-accounts describe "$RUNTIME_SA" --project="$PROJECT"`
   - create: `gcloud iam service-accounts create "$RUNTIME_SA_ID" --display-name="$SERVICE Cloud Run runtime" --project="$PROJECT"`

2. Create the deploy service account with `ensure_resource`:
   - describe: `gcloud iam service-accounts describe "$DEPLOY_SA" --project="$PROJECT"`
   - create: `gcloud iam service-accounts create "$DEPLOY_SA_ID" --display-name="$SERVICE GitHub Actions deployer" --project="$PROJECT"`

3. Grant the runtime SA `roles/datastore.user` at the project level:
   ```bash
   gcloud projects add-iam-policy-binding "$PROJECT" \
     --member="serviceAccount:$RUNTIME_SA" \
     --role="roles/datastore.user" \
     --condition=None --quiet
   ```
   `--condition=None` suppresses the interactive condition prompt that `add-iam-policy-binding` raises when the policy already contains conditional bindings; without it the script hangs on a re-run.

4. Grant the deploy SA `roles/run.admin` and `roles/artifactregistry.writer` at the project level, using the same flag pattern. Write these as a loop over the two role strings rather than two near-identical blocks.

5. Grant the deploy SA `roles/iam.serviceAccountUser` **on the runtime service account resource**, not on the project:
   ```bash
   gcloud iam service-accounts add-iam-policy-binding "$RUNTIME_SA" \
     --member="serviceAccount:$DEPLOY_SA" \
     --role="roles/iam.serviceAccountUser" \
     --project="$PROJECT" --quiet
   ```
   Add a comment stating why: the deploy job needs to act as the runtime SA when it sets `--service-account` on the Cloud Run deploy, and scoping this to the one SA avoids granting impersonation of every SA in the project.

**Tests added/updated:**
- Fresh-project run logs `iam service-accounts create onenote-mcp-run` and `iam service-accounts create onenote-mcp-deploy`.
- Already-configured run logs neither create, but still logs all four `add-iam-policy-binding` calls, confirming role grants are re-asserted on every run.
- The `roles/datastore.user` binding names the runtime SA as its member, and `roles/run.admin` and `roles/artifactregistry.writer` name the deploy SA.
- `roles/iam.serviceAccountUser` is logged on an `iam service-accounts add-iam-policy-binding` line, not on a `projects add-iam-policy-binding` line. Assert both the presence of the former and the absence of `projects add-iam-policy-binding ... roles/iam.serviceAccountUser`.
- `SERVICE=other-name` override produces SA ids `other-name-run` and `other-name-deploy`.

**Verification:**
- [ ] `bash scripts/test/run.sh` exits 0
- [ ] Both stub modes exit 0
- [ ] `shellcheck scripts/gcp-bootstrap.sh` clean
- [ ] `grep -rn "keys create" scripts/` returns nothing

---

## Phase 4: Workload Identity Pool, OIDC provider, repo binding

**Goal:** Create the federation path from this GitHub repository to the deploy SA, restricted to this repository.
**Addresses:** AC-6, AC-7, AC-8

**Files:**
- Modify: `scripts/gcp-bootstrap.sh`
- Modify: `scripts/test/bootstrap.test.sh`

**Steps:**

1. Resolve the numeric project number and build the resource names:
   ```bash
   PROJECT_NUMBER="$(gcloud projects describe "$PROJECT" --format='value(projectNumber)')"
   [[ -n "$PROJECT_NUMBER" ]] || die "Could not resolve project number for $PROJECT"
   POOL_RESOURCE="projects/${PROJECT_NUMBER}/locations/global/workloadIdentityPools/${POOL_ID}"
   WIF_PROVIDER="${POOL_RESOURCE}/providers/${PROVIDER_ID}"
   PRINCIPAL_SET="principalSet://iam.googleapis.com/${POOL_RESOURCE}/attribute.repository/${GITHUB_REPO}"
   ```

2. Create the pool with `ensure_resource`:
   - describe: `gcloud iam workload-identity-pools describe "$POOL_ID" --location=global --project="$PROJECT"`
   - create: `gcloud iam workload-identity-pools create "$POOL_ID" --location=global --display-name="GitHub Actions" --project="$PROJECT"`
   Add a comment noting that deleted pools are soft-deleted for 30 days and a re-create under the same id fails until purged or undeleted.

3. Define the mapping and condition once, as variables, so the create and update paths cannot drift:
   ```bash
   ATTR_MAPPING="google.subject=assertion.sub,attribute.repository=assertion.repository"
   ATTR_CONDITION="assertion.repository == \"${GITHUB_REPO}\""
   ```

4. Create or update the OIDC provider. This is the one resource that does not use the plain `ensure_resource` skip-if-exists path. Run the describe; on success run `update-oidc`, on failure run `create-oidc`:
   ```bash
   gcloud iam workload-identity-pools providers create-oidc "$PROVIDER_ID" \
     --location=global \
     --workload-identity-pool="$POOL_ID" \
     --display-name="GitHub OIDC" \
     --issuer-uri="https://token.actions.githubusercontent.com" \
     --attribute-mapping="$ATTR_MAPPING" \
     --attribute-condition="$ATTR_CONDITION" \
     --project="$PROJECT"
   ```
   The `update-oidc` invocation takes the same `--attribute-mapping` and `--attribute-condition` flags, minus `--issuer-uri` and `--display-name`.
   Add a comment stating why update rather than skip: a pre-existing provider created without the condition would let any GitHub repository impersonate the deploy SA, and skipping would silently leave that in place.

5. Grant `roles/iam.workloadIdentityUser` to the principalSet on the deploy SA:
   ```bash
   gcloud iam service-accounts add-iam-policy-binding "$DEPLOY_SA" \
     --member="$PRINCIPAL_SET" \
     --role="roles/iam.workloadIdentityUser" \
     --project="$PROJECT" --quiet
   ```

**Tests added/updated:**
- Fresh-project run logs `workload-identity-pools create github-pool` with `--location=global`.
- Fresh-project run logs `providers create-oidc github-provider` containing `--issuer-uri=https://token.actions.githubusercontent.com`.
- The create-oidc line contains `attribute.repository=assertion.repository` in its mapping — the mapping the principalSet depends on.
- The create-oidc line contains the attribute condition `assertion.repository == "dovrosenberg/onenote-mcp"`.
- Already-configured run logs `providers update-oidc` and does **not** log `providers create-oidc`, and the update line carries the same attribute condition. This is the regression test for the security hole described in step 4.
- The `roles/iam.workloadIdentityUser` binding member starts with `principalSet://iam.googleapis.com/projects/123456789012/locations/global/workloadIdentityPools/github-pool/attribute.repository/dovrosenberg/onenote-mcp` — asserting the stub's fixed project number, so a regression in the resource-name construction is caught.
- `GITHUB_REPO=someone/else` override changes both the attribute condition and the principalSet, confirming the repo restriction is not hardcoded in one place and missing from the other.

**Verification:**
- [ ] `bash scripts/test/run.sh` exits 0
- [ ] Both stub modes exit 0
- [ ] `shellcheck scripts/gcp-bootstrap.sh` clean
- [ ] Manual read of the create-oidc and update-oidc calls confirms both carry `--attribute-condition`

---

## Phase 5: Output block, no-key guarantee, full re-run test

**Goal:** Print copy-pasteable `gh` commands matching issue #4 and `project-spec.md:197`, and prove the whole script is clean on a second run.
**Addresses:** AC-9, AC-10, AC-11

**Files:**
- Modify: `scripts/gcp-bootstrap.sh`
- Modify: `scripts/test/bootstrap.test.sh`

**Steps:**

1. Add the final output block, emitted via a single quoted heredoc so the operator can select and paste it. Variable names must match `project-spec.md:284` and issue #4 exactly:
   ```
   gh variable set GCP_PROJECT --body "$PROJECT"
   gh variable set GCP_REGION --body "$REGION"
   gh variable set GAR_REGION --body "$GAR_REGION"
   gh variable set WIF_PROVIDER --body "$WIF_PROVIDER"
   gh variable set DEPLOY_SA --body "$DEPLOY_SA"
   gh variable set RUNTIME_SA --body "$RUNTIME_SA"
   ```
   The heredoc uses an unquoted delimiter so these expand. Precede it with a `log` line: `Run these from a clone of ${GITHUB_REPO} to complete issue #4:`.

2. Append a second block covering the values this script cannot know, with placeholders and the generator commands from issue #4 as comments:
   ```
   gh variable set ONENOTE_CLIENT_ID    --body "<Azure app registration client ID from #1>"
   gh variable set ONENOTE_AUTHORITY    --body "https://login.microsoftonline.com/common"
   gh variable set MCP_OAUTH_CLIENT_ID  --body "$(openssl rand -hex 16)"
   gh secret   set MCP_OAUTH_CLIENT_SECRET --body "$(openssl rand -base64 32)"
   gh secret   set MCP_TOKEN_SIGNING_KEY   --body "$(openssl rand -base64 48)"
   ```
   These go in a **quoted** heredoc so `$(openssl ...)` is printed literally rather than executed by the bootstrap script. Add a printed note that `MCP_OAUTH_CLIENT_ID` and `MCP_OAUTH_CLIENT_SECRET` must be saved somewhere retrievable, because they are typed into Claude's connector settings in issue #26 and GitHub will not display the secret again.

3. Add a closing note that no service-account key was created and that none is needed, so a reader of the output is not left looking for one.

4. Add a header comment block at the top of `scripts/gcp-bootstrap.sh`: what the script provisions, that it is idempotent, that it creates no service-account keys, and that issue #3 is the manual runbook for executing it.

**Tests added/updated:**
- The script's stdout contains `gh variable set WIF_PROVIDER` followed by the full resource name `projects/123456789012/locations/global/workloadIdentityPools/github-pool/providers/github-provider`.
- Stdout contains `gh variable set DEPLOY_SA --body "onenote-mcp-deploy@test-project.iam.gserviceaccount.com"` and the matching `RUNTIME_SA` line.
- Stdout contains all six identifier variables plus the three placeholder variables and the two secrets, asserted individually by name so a dropped line fails the test.
- Stdout contains the literal string `$(openssl rand -base64 32)` — proving the quoted heredoc did not execute it.
- Stdout contains no `BEGIN PRIVATE KEY` and no `"private_key"`.
- The script source contains no `iam service-accounts keys create` — asserted by grepping the script file itself from the test.
- No `*.json` file exists in the repository working tree after a run — asserted with `find . -name '*.json' -newer` against a marker file created before the run.
- **Full idempotency test:** run the script twice in sequence against the same stub log with `STUB_MISSING` set for the first run and unset for the second. Assert the second run exits 0, and assert its stdout still contains the complete `gh variable set` output block. A second run that succeeds but prints nothing useful would fail issue #3's step "copy the output block".

**Verification:**
- [ ] `bash scripts/test/run.sh` exits 0
- [ ] Two-run idempotency test passes
- [ ] `shellcheck scripts/gcp-bootstrap.sh` clean
- [ ] `grep -rn "keys create" scripts/` returns nothing
- [ ] Manual read: every variable name in the output block matches the `vars.*` references in `project-spec.md:284`

---

## Final Verification

- [x] `bash scripts/test/run.sh` exits 0 — full suite, 64 assertions passing (`bash -n`, `shellcheck`, `bootstrap.test.sh`)
- [x] `shellcheck` reports no warnings on all four files (run via `npx --yes shellcheck` — no system shellcheck on this machine, so `run.sh` falls back to npx)
- [x] All four created files are executable — `git ls-files -s scripts/` shows mode `100755` for each
- [x] `grep -n "keys create\|--key-file-type\|private_key" scripts/gcp-bootstrap.sh` returns nothing
      (scoped to the script, not all of `scripts/` — `bootstrap.test.sh` contains these strings
      as the assertion patterns that enforce the rule)
- [x] The variable names in the output block are byte-identical to the `vars.*` references in the workflow sketch at `project-spec.md:284` and to the checklist in issue #4
- [x] Acceptance criteria traceability:
  - **AC-1** (enable six APIs): Phase 2 — single `gcloud services enable` call with all six service names, asserted by a test that greps each name.
  - **AC-2** (Artifact Registry repo): Phase 2 — `ensure_resource` guard around `artifacts repositories create --repository-format=docker --location=$GAR_REGION`.
  - **AC-3** (Firestore Native): Phase 2 — `ensure_resource` guard around `firestore databases create --location=$REGION --type=firestore-native`.
  - **AC-4** (runtime SA + datastore.user): Phase 3 — SA create guard plus a project-level `add-iam-policy-binding` for `roles/datastore.user`.
  - **AC-5** (deploy SA + three roles): Phase 3 — SA create guard, project-level bindings for `roles/run.admin` and `roles/artifactregistry.writer`, and an SA-scoped binding for `roles/iam.serviceAccountUser` on the runtime SA, with a test asserting the last is not project-scoped.
  - **AC-6** (pool + OIDC provider): Phase 4 — pool create guard and `providers create-oidc --issuer-uri=https://token.actions.githubusercontent.com`.
  - **AC-7** (repo-only attribute condition): Phase 4 — `--attribute-condition` on both the create and update paths, with the required `attribute.repository` mapping, and a test asserting the already-exists path re-applies the condition rather than skipping.
  - **AC-8** (workloadIdentityUser on principalSet): Phase 4 — SA-scoped binding whose member is the `attribute.repository/$GITHUB_REPO` principalSet, asserted against the stub's fixed project number.
  - **AC-9** (printed gh commands): Phase 5 — two heredocs, one expanding the discovered values and one printing the `openssl` generators literally; every variable name asserted individually.
  - **AC-10** (no JSON key): Phase 5 — asserted three ways: the script source contains no key-create command, stdout contains no key material, and no new `.json` file appears after a run.
  - **AC-11** (re-runnable): Phase 5 — the two-run test, backed by the `ensure_resource` guard introduced in Phase 1 and applied in Phases 2–4, the `--condition=None --quiet` flags on every `add-iam-policy-binding`, and the create-or-update path on the OIDC provider.
  - **AC-12** (env-var parameters): Phase 1 — `${VAR:-default}` block at the top of the file, with override tests for `REGION` (Phase 1), `GAR_REGION` (Phase 2), `SERVICE` (Phase 3), and `GITHUB_REPO` (Phase 4).
