#!/usr/bin/env bash
# Behavioural tests for the Dockerfile and .dockerignore.
#
# Builds the image for linux/amd64 and asserts on what is inside it: the resvg
# glibc prebuild survived a production-only `npm ci`, no dev dependencies came
# along, dist/ is present, and the build context excluded src/ and test/.
#
# Not run by `npm test`, and not run by scripts/test/run.sh unless
# RUN_DOCKER_TESTS=1 is set — a docker build is slow and needs a daemon.
#
# Deliberately not `set -e`: a failing assertion must not stop the run, so that
# every failure is reported at once.
set -uo pipefail
cd "$(dirname "$0")/../.." || exit 1

IMAGE=onenote-mcp:test
FAILURES=0
BUILD_LOG=""

pass() { printf '  ok   %s\n' "$1"; }
fail() { printf '  FAIL %s\n' "$1"; FAILURES=$((FAILURES + 1)); }

# shellcheck disable=SC2329  # invoked indirectly, by the EXIT trap below
cleanup() {
  [[ -n "$BUILD_LOG" ]] && rm -f "$BUILD_LOG"
  docker image rm -f "$IMAGE" >/dev/null 2>&1
  return 0
}
trap cleanup EXIT

if ! command -v docker >/dev/null 2>&1; then
  echo "SKIP: docker not available"
  exit 0
fi

# assert_in_image DESCRIPTION SH_COMMAND -- runs SH_COMMAND in a throwaway
# container and asserts it exits 0.
assert_in_image() {
  local desc="$1" cmd="$2" out
  if out=$(docker run --rm --entrypoint sh "$IMAGE" -c "$cmd" 2>&1); then
    pass "$desc"
  else
    fail "$desc"
    printf '       command: %s\n' "$cmd"
    [[ -n "$out" ]] && printf '       output: %s\n' "$out"
  fi
}

echo "== docker build =="
BUILD_LOG=$(mktemp)
if docker build --platform linux/amd64 -t "$IMAGE" . >"$BUILD_LOG" 2>&1; then
  pass "image builds for linux/amd64"
else
  fail "image builds for linux/amd64"
  tail -n 30 "$BUILD_LOG" | sed 's/^/       /'
  # Nothing downstream can run without an image.
  exit 1
fi

echo "== image contents =="

# AC-3: a production-only `npm ci` must still pull the optional platform package.
assert_in_image "resvg glibc prebuild present in the runtime image" \
  'ls node_modules/@resvg/resvg-js-linux-x64-gnu/*.node'

# The runtime stage installs with --omit=dev. A COPY of the build stage's
# node_modules would satisfy every other check while shipping the compiler.
assert_in_image "dev dependencies absent (no node_modules/typescript)" \
  '! test -e node_modules/typescript'

assert_in_image "compiled entrypoint present at dist/index.js" \
  'test -f dist/index.js'

# .dockerignore excludes both, and neither is COPYed into the runtime stage.
# This fails on a `COPY . .` Dockerfile or a dropped .dockerignore.
assert_in_image "build context excluded: no src/ or test/ in the image" \
  '! test -e src && ! test -e test'

echo
if [[ "$FAILURES" -eq 0 ]]; then
  echo "all docker checks passed"
else
  echo "$FAILURES docker check(s) failed"
fi
exit "$((FAILURES > 0))"
