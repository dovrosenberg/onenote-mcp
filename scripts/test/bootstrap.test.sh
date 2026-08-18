#!/usr/bin/env bash
# Behavioural tests for scripts/gcp-bootstrap.sh.
#
# The script is not run against real GCP. A stub `gcloud` is placed first on
# PATH; it records every invocation, and $STUB_MISSING controls which describe
# calls fail, which is how a fresh project versus an already-configured one is
# simulated. Assertions are on the recorded command lines and on stdout.
#
# Deliberately not `set -e`: a failing assertion must not stop the run, so
# that every failure is reported at once.
set -uo pipefail
cd "$(dirname "$0")/../.." || exit 1

FAILURES=0
GCLOUD_LOG=""
BOOTSTRAP_OUT=""
BOOTSTRAP_STATUS=0

pass() { printf '  ok   %s\n' "$1"; }
fail() { printf '  FAIL %s\n' "$1"; FAILURES=$((FAILURES + 1)); }

assert_logged() {
  if grep -qF -- "$1" "$GCLOUD_LOG"; then pass "$2"; else
    fail "$2"
    printf '       expected in gcloud log: %s\n' "$1"
  fi
}

assert_not_logged() {
  if grep -qF -- "$1" "$GCLOUD_LOG"; then
    fail "$2"
    printf '       unexpectedly in gcloud log: %s\n' "$1"
  else pass "$2"; fi
}

assert_out() {
  if printf '%s' "$BOOTSTRAP_OUT" | grep -qF -- "$1"; then pass "$2"; else
    fail "$2"
    printf '       expected in stdout: %s\n' "$1"
  fi
}

assert_not_out() {
  if printf '%s' "$BOOTSTRAP_OUT" | grep -qF -- "$1"; then
    fail "$2"
    printf '       unexpectedly in stdout: %s\n' "$1"
  else pass "$2"; fi
}

assert_status() {
  if [[ "$BOOTSTRAP_STATUS" == "$1" ]]; then pass "$2"; else
    fail "$2"
    printf '       expected exit %s, got %s\n' "$1" "$BOOTSTRAP_STATUS"
  fi
}

# run_bootstrap [VAR=value ...] -- runs the script with a fresh log.
# Set KEEP_LOG=1 to append to the existing log instead (two-run test).
run_bootstrap() {
  if [[ "${KEEP_LOG:-0}" != "1" || -z "$GCLOUD_LOG" ]]; then
    GCLOUD_LOG="$(mktemp)"
  fi
  BOOTSTRAP_OUT="$(
    env "$@" \
      GCLOUD_LOG="$GCLOUD_LOG" \
      PATH="$PWD/scripts/test/stubs:$PATH" \
      bash scripts/gcp-bootstrap.sh 2>&1
  )"
  BOOTSTRAP_STATUS=$?
}

# Every describe the script makes, as underscore-joined patterns for
# STUB_MISSING (the stub converts underscores back to spaces).
ALL_MISSING="artifacts_repositories_describe firestore_databases_describe iam_service-accounts_describe iam_workload-identity-pools_describe iam_workload-identity-pools_providers_describe"

echo "--- Phase 1: parameters and preflight ---"

run_bootstrap PROJECT=
assert_status 1 "unset PROJECT exits non-zero"
assert_out "PROJECT" "error message names PROJECT"

run_bootstrap PROJECT=test-project
assert_status 0 "PROJECT set exits 0"

run_bootstrap PROJECT=test-project REGION=europe-west1
assert_out "europe-west1" "REGION override appears in the parameter summary"

echo "--- Phase 2: APIs, Artifact Registry, Firestore ---"

run_bootstrap PROJECT=test-project STUB_MISSING="$ALL_MISSING"
assert_status 0 "fresh project run exits 0"
for api in run artifactregistry firestore iamcredentials sts cloudresourcemanager; do
  assert_logged "${api}.googleapis.com" "services enable includes ${api}.googleapis.com"
done
assert_logged "artifacts repositories create onenote-mcp" "fresh run creates the Artifact Registry repo"
assert_logged "--repository-format=docker" "Artifact Registry repo is a Docker repo"
assert_logged "firestore databases create" "fresh run creates the Firestore database"
assert_logged "--type=firestore-native" "Firestore database is Native mode"

run_bootstrap PROJECT=test-project STUB_MISSING=
assert_status 0 "already-configured run exits 0"
assert_logged "services enable" "already-configured run still enables APIs"
assert_not_logged "artifacts repositories create" "already-configured run skips repo create"
assert_not_logged "firestore databases create" "already-configured run skips database create"

run_bootstrap PROJECT=test-project REGION=nam5 GAR_REGION=europe-west4 STUB_MISSING="$ALL_MISSING"
assert_logged "artifacts repositories create onenote-mcp --repository-format=docker --location=europe-west4" \
  "GAR_REGION drives the Artifact Registry location"
assert_logged "firestore databases create --location=nam5" \
  "REGION drives the Firestore location, independently of GAR_REGION"

echo "--- Phase 3: service accounts and IAM grants ---"

RUN_SA="onenote-mcp-run@test-project.iam.gserviceaccount.com"
DEP_SA="onenote-mcp-deploy@test-project.iam.gserviceaccount.com"

run_bootstrap PROJECT=test-project STUB_MISSING="$ALL_MISSING"
assert_logged "iam service-accounts create onenote-mcp-run" "fresh run creates the runtime SA"
assert_logged "iam service-accounts create onenote-mcp-deploy" "fresh run creates the deploy SA"

run_bootstrap PROJECT=test-project STUB_MISSING=
assert_not_logged "iam service-accounts create" "already-configured run skips SA creates"
assert_logged "projects add-iam-policy-binding test-project --member=serviceAccount:$RUN_SA --role=roles/datastore.user" \
  "roles/datastore.user is granted to the runtime SA on every run"
assert_logged "projects add-iam-policy-binding test-project --member=serviceAccount:$DEP_SA --role=roles/run.admin" \
  "roles/run.admin is granted to the deploy SA on every run"
assert_logged "projects add-iam-policy-binding test-project --member=serviceAccount:$DEP_SA --role=roles/artifactregistry.writer" \
  "roles/artifactregistry.writer is granted to the deploy SA on every run"
assert_logged "iam service-accounts add-iam-policy-binding $RUN_SA --member=serviceAccount:$DEP_SA --role=roles/iam.serviceAccountUser" \
  "roles/iam.serviceAccountUser is scoped to the runtime SA"
assert_not_logged "projects add-iam-policy-binding test-project --member=serviceAccount:$DEP_SA --role=roles/iam.serviceAccountUser" \
  "roles/iam.serviceAccountUser is NOT granted at the project level"

run_bootstrap PROJECT=test-project SERVICE=other-name STUB_MISSING="$ALL_MISSING"
assert_logged "iam service-accounts create other-name-run" "SERVICE override drives the runtime SA id"
assert_logged "iam service-accounts create other-name-deploy" "SERVICE override drives the deploy SA id"

if [[ "$FAILURES" -ne 0 ]]; then
  printf '\n%s assertion(s) failed\n' "$FAILURES"
  exit 1
fi
printf '\nAll assertions passed\n'
