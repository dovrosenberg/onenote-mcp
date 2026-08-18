#!/usr/bin/env bash
# Full check suite: syntax, lint, then the behavioural tests.
set -uo pipefail
cd "$(dirname "$0")/../.." || exit 1

status=0
FILES=(scripts/gcp-bootstrap.sh scripts/test/bootstrap.test.sh scripts/test/docker.test.sh scripts/test/run.sh scripts/test/stubs/gcloud)

echo "== bash -n =="
for f in "${FILES[@]}"; do
  bash -n "$f" || { echo "FAIL: syntax error in $f"; status=1; }
done

echo "== shellcheck =="
if command -v shellcheck >/dev/null 2>&1; then
  shellcheck "${FILES[@]}" || status=1
elif command -v npx >/dev/null 2>&1; then
  # No system shellcheck. npx fetches the binary; needs network the first time.
  npx --yes shellcheck "${FILES[@]}" || status=1
else
  echo "shellcheck not installed and npx unavailable, skipping"
fi

echo "== bootstrap.test.sh =="
bash scripts/test/bootstrap.test.sh || status=1

# A docker build plus a container start is slow and needs a daemon, so it is
# opt-in. The skip is printed rather than silent, so the suite never reports
# success while checking nothing.
echo "== docker.test.sh =="
if [ "${RUN_DOCKER_TESTS:-}" = "1" ]; then
  bash scripts/test/docker.test.sh || status=1
else
  echo "skipped: set RUN_DOCKER_TESTS=1 to build and run the container"
fi

exit "$status"
