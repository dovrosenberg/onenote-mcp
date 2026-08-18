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

if [[ "$FAILURES" -ne 0 ]]; then
  printf '\n%s assertion(s) failed\n' "$FAILURES"
  exit 1
fi
printf '\nAll assertions passed\n'
