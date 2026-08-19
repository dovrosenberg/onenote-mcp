#!/usr/bin/env bash
# Behavioural tests for the Dockerfile and .dockerignore.
#
# Builds the image for linux/amd64 and asserts on what is inside it: the resvg
# glibc prebuild survived a production-only `npm ci`, no dev dependencies came
# along, dist/ is present, and the build context excluded src/ and test/.
#
# Then runs it: resvg renders an SVG to PNG, the server binds the port given in
# PORT and answers /healthz with 200, and SIGTERM shuts it down cleanly.
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
CID=""

# Deliberately not the 8080 default, so a passing /healthz check proves the server
# read PORT rather than falling back to the value baked into src/config.ts.
CONTAINER_PORT=9090

# Enough for loadConfig to succeed, and nothing more. ONENOTE_AUTHORITY must be an
# https URL and MCP_TOKEN_SIGNING_KEY at least 32 characters, per the checks in
# src/config.ts. FIRESTORE_CACHE_DOC and GOOGLE_CLOUD_PROJECT are left unset: they
# are optional, and omitting them is part of showing /healthz touches neither
# Graph nor Firestore. None of these are real credentials.
FAKE_ENV=(
  -e ONENOTE_CLIENT_ID=00000000-0000-0000-0000-000000000000
  -e ONENOTE_AUTHORITY=https://login.microsoftonline.com/common
  -e MCP_OAUTH_CLIENT_ID=test-client
  -e MCP_OAUTH_CLIENT_SECRET=test-secret
  -e MCP_TOKEN_SIGNING_KEY=0123456789abcdef0123456789abcdef
  -e MCP_PUBLIC_URL=https://onenote-mcp.example.run.app
)

pass() { printf '  ok   %s\n' "$1"; }
fail() { printf '  FAIL %s\n' "$1"; FAILURES=$((FAILURES + 1)); }

# shellcheck disable=SC2329  # invoked indirectly, by the EXIT trap below
cleanup() {
  [[ -n "$BUILD_LOG" ]] && rm -f "$BUILD_LOG"
  [[ -n "$CID" ]] && docker rm -f "$CID" >/dev/null 2>&1
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

echo "== resvg render =="

# AC-8. An inline program rather than a shipped test file: the runtime image
# contains no test/ and no TypeScript. Asserting the four-byte PNG signature and
# not merely a zero exit means an empty buffer fails.
render_status=0
render_out=$(docker run --rm --entrypoint node "$IMAGE" -e '
const { Resvg } = require("@resvg/resvg-js");
const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="10" height="10"><rect width="10" height="10" fill="red"/></svg>`;
const png = new Resvg(svg).render().asPng();
const sig = [0x89, 0x50, 0x4e, 0x47];
if (png.length < 8 || !sig.every((b, i) => png[i] === b)) throw new Error("not a png");
console.log("resvg ok, " + png.length + " bytes");
' 2>&1) || render_status=$?
if [[ "$render_status" -eq 0 && "$render_out" == *"resvg ok"* ]]; then
  pass "trivial SVG renders to PNG inside the container ($render_out)"
else
  fail "trivial SVG renders to PNG inside the container"
  printf '%s\n' "$render_out" | sed 's/^/       /'
fi

echo "== container startup and the health endpoint =="

# Port 0 on the host: the daemon picks a free one, so a busy host port cannot make
# this flaky. The container side is CONTAINER_PORT, passed in as PORT.
run_status=0
CID=$(docker run -d -e "PORT=$CONTAINER_PORT" "${FAKE_ENV[@]}" \
  -p "127.0.0.1:0:$CONTAINER_PORT" "$IMAGE" 2>&1) || run_status=$?
if [[ "$run_status" -ne 0 || -z "$CID" ]]; then
  fail "container starts"
  printf '%s\n' "$CID" | sed 's/^/       /'
  CID=""
else
  pass "container starts"

  hostport=$(docker port "$CID" "$CONTAINER_PORT" 2>/dev/null | head -n 1 | sed 's/.*://')
  if [[ -z "$hostport" ]]; then
    fail "container port $CONTAINER_PORT is published"
  else
    pass "container port $CONTAINER_PORT is published on host port $hostport"

    body=$(mktemp)
    code=""
    SECONDS=0
    # Cloud Run's default startup timeout is 240s. 10s is the "well under" bar;
    # 30s is only the give-up point, so a failure reports a real number.
    while [[ "$SECONDS" -lt 30 ]]; do
      code=$(curl -s -o "$body" -w '%{http_code}' \
        "http://127.0.0.1:$hostport/healthz" 2>/dev/null)
      [[ "$code" == "200" ]] && break
      sleep 0.2
    done
    elapsed="$SECONDS"

    if [[ "$code" == "200" ]]; then
      pass "/healthz returns 200 on the port given in PORT"
    else
      fail "/healthz returns 200 on the port given in PORT"
      printf '       last status code: %s\n' "${code:-none}"
      docker logs "$CID" 2>&1 | tail -n 20 | sed 's/^/       /'
    fi

    if grep -q '"status":"ok"' "$body"; then
      pass "/healthz body reports status ok"
    else
      fail "/healthz body reports status ok"
      sed 's/^/       /' "$body"
    fi

    # The alias, checked in the image rather than only in the unit test. It is the
    # path an external check has to use: Google's frontend swallows /healthz on a
    # run.app URL, so a build that served only /healthz would look healthy here and
    # be unmonitorable once deployed.
    alias_body=$(mktemp)
    alias_code=$(curl -s -o "$alias_body" -w '%{http_code}' \
      "http://127.0.0.1:$hostport/health" 2>/dev/null)
    if [[ "$alias_code" == "200" ]] && grep -q '"status":"ok"' "$alias_body"; then
      pass "/health answers the same 200"
    else
      fail "/health answers the same 200"
      printf '       status code: %s\n' "${alias_code:-none}"
      sed 's/^/       /' "$alias_body"
    fi
    rm -f "$alias_body"

    # No Firestore or Graph variable was set, so a 200 here also shows the
    # endpoint reaches neither.
    if [[ "$code" == "200" && "$elapsed" -lt 10 ]]; then
      pass "ready in ${elapsed}s, under the 10s bar"
    else
      fail "ready in ${elapsed}s, expected under 10s"
    fi
    rm -f "$body"
  fi

  echo "== SIGTERM shutdown =="

  # docker stop sends SIGTERM. This passes only with the exec-form CMD: under
  # shell form the signal goes to /bin/sh and the handler in src/index.ts never
  # runs, so the container is SIGKILLed after the grace period instead.
  docker stop "$CID" >/dev/null 2>&1
  exit_code=$(docker inspect -f '{{.State.ExitCode}}' "$CID" 2>/dev/null)
  if [[ "$exit_code" == "0" ]]; then
    pass "SIGTERM produces a clean exit (code 0)"
  else
    fail "SIGTERM produces a clean exit, got exit code ${exit_code:-unknown}"
  fi

  if docker logs "$CID" 2>&1 | grep -q 'SIGTERM received, shutting down'; then
    pass "shutdown handler ran (node is PID 1)"
  else
    fail "shutdown handler ran (node is PID 1)"
    docker logs "$CID" 2>&1 | tail -n 20 | sed 's/^/       /'
  fi
fi

echo
if [[ "$FAILURES" -eq 0 ]]; then
  echo "all docker checks passed"
else
  echo "$FAILURES docker check(s) failed"
fi
exit "$((FAILURES > 0))"
