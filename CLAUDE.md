# CLAUDE.md

Working notes for agents in this repository. `project-spec.md` is the authoritative
design document — read the relevant section of it before changing behaviour.
`api-overview.md` is what Microsoft Graph's OneNote endpoints actually do: the resource
paths, the query options and their limits, and the places the service contradicts its own
documentation. Read it before adding or changing a Graph call, and add anything new that
a live run reveals.

## Directory structure

```
onenote-mcp/
├── src/                       # TypeScript source, the only thing that ships│
├── test/                      # node --test suites; test/<name>.test.ts covers src/<name>.ts
│   ├── fixtures/              # hand-authored InkML and HTML; never a captured page dump
├── scripts/                   # operational shell scripts, not part of the built artifact
│   └── test/                  # bash tests, run by scripts/test/run.sh
├── .github/workflows/         # deploy.yml, the only CI; builds, tests, pushes, deploys
└── dist/                      # build output; gitignored, never edited, never committed
```

A test file is named for the source file it covers: `test/config.test.ts` covers
`src/config.ts`. Not every source file has one. `src/index.ts` is a top-level entrypoint
whose behaviour is process exit codes and stderr, so it has no unit test; check it by
running it. Unsetting a required variable and running `npm start` must print the list of
missing names and exit 1 with no line matching `at …(`, and binding a port already in
use must print one readable line rather than an `EADDRINUSE` stack.

`test/bootstrap.test.ts` spawns `src/bootstrap.ts` as a child process, because the module
signs in at import time and exports nothing. It covers only what happens before the CLI
needs a human: the exit code and the message on an incomplete environment. The
device-code sign-in, the cache write, and the Graph probe have no automated test and
cannot get one — they need a browser, the real Entra tenant, and Firestore, and no
credential that could stand in for one may be committed. Verify them by running the CLI.

`test/graph-auth.test.ts` drives `acquireGraphToken` through a hand-written
`SilentTokenSource` and never constructs a `PublicClientApplication`, so
`createGraphAuth` has no automated test. Testing it needs a cache seeded by a real
device-code sign-in, and no credential that could seed one may be committed. Nothing
verifies it until an operator runs `npm run bootstrap` and then the server against the
same document.

`test/graph-structure.test.ts` drives every call through a fake `fetch` whose routes
are keyed by the exact URL, so an unrouted URL fails the test and each behavioural
assertion is also an assertion about the URL that was built. What it cannot check is
whether Graph accepts those URLs: the `$select`, `$orderby`, and `$top` strings are
copied from the validated recon script in Appendix A, and nothing verifies them against
the service until an operator runs the server against the real tenant. The paging and
nesting fixtures are hand-written, because a page of Graph results large enough to carry
an `@odata.nextLink` cannot be captured without real notebook names in it.

`test/ink.test.ts` renders the committed fixtures all the way to PNG bytes and checks the
dimensions in the PNG header, so the resvg dependency is exercised rather than mocked.
The fixtures are hand-authored and small. A real page dump may not be committed: rendered
ink is legible personal notes. What no test covers is whether Graph's InkML matches the
fixtures' shape — the channel order, the units, and the multipart framing come from the
validated recon script in Appendix A, and nothing confirms them against the service until
an operator runs the server against the real tenant.

`test/page-html.test.ts` runs the trimmer over `test/fixtures/styled-page.html`, which is
hand-authored in the shape Graph emits — absolutely positioned outline divs, a styling
span around every text run, `data-tag` list items, a table, and the InkNode comments. A
real page dump may not be committed. What no test covers is whether Graph's markup
matches that shape; the trimmer is tolerant rather than strict for that reason, and
nothing confirms the shape until an operator runs the server against the real tenant.

`test/mcp-server.test.ts` starts a real HTTP server and speaks JSON-RPC to it, so the
transport is exercised rather than mocked: the assertions about statelessness (no
`Mcp-Session-Id` header, `tools/list` answered by a server that saw no `initialize`) are
only meaningful over the wire. It covers the real tool list through `createApp`, and the
failure paths through `mcpRouter` with hand-written tools injected. What no test covers
is whether a real MCP client accepts the responses; nothing confirms that until an
operator points one at the deployed URL.

`test/structure-tools.test.ts` and `test/page-search.test.ts` drive plain fake objects
rather than a fake `fetch`: `StructureClient` and `SearchableStructure` are the narrow
slices of `GraphStructure` those modules call, and the URLs are already asserted in
`test/graph-structure.test.ts`. The fake tree is a notebook per year with a section group
per month, nested twice, because a one-level tree would pass a walk that does not recurse.
What no test covers is whether the search bounds are the right size — 200 sections and 25
seconds are chosen against the account described in `project-spec.md`, not measured, and
nothing says whether an unscoped search finishes inside an MCP client's timeout until an
operator runs one against the real tenant.

`test/page-tools.test.ts` drives `get_page_content` through a fake `PageContentClient`,
but the ink in it is the committed InkML fixture rendered for real: the image block's
base64 is decoded and its PNG header checked against the width the JSON block reports,
so a result that claimed an image and carried something else would fail. It forces the
downscale path by injecting a one-byte budget rather than by committing a large fixture.
What no test covers is whether an MCP client actually shows the image block to its model,
or whether `MAX_INK_PNG_BYTES` is the right number — 750 KB of PNG, about 1 MB once
base64-encoded, is a budget chosen rather than measured against any client's cap.

`test/page-write.test.ts` drives the write client through a fake `fetch` whose routes are
keyed by method and exact URL, so an unexpected verb is an unrouted request and every
assertion about a change array is also an assertion about where it was sent. What it
cannot check is whether Graph accepts any of it. That was settled separately: the change
arrays come from the validated spike in `api-overview.md`, and all three tools were run
against the live account on 2026-08-18 in a scratch notebook — create, append, rename,
and a read-back of each. The fake fetch still cannot notice the service changing its
mind.

`test/write-tools.test.ts` drives the five writing tools through a fake write client, and
every refusal test also asserts the client was not called. That is the point of the file:
a rejected argument must cost no Graph request, and on `create_page` a request that went
through would leave a page behind. The two `_by_name` tools also get a fake
`WriteLookupStructure` that counts its calls, because two of their properties are things
they do not do: the common path is one `getExpandedTree` and no fallback, and a refused
fragment costs not even that. The container matching rules themselves are asserted in
`test/name-lookup.test.ts` rather than again here; what this file adds for
`append_to_page_by_name` is that a title matching no page or two pages reaches neither the
write client nor the page reader, and that an append by name pads for ink identically to
one by id — the padding assertions are the only thing saying the two tools share the
`append` helper.

`test/ink-preservation.integration.test.ts` is the only test that talks to the real
account, and it is skipped unless the environment names a page: `ONENOTE_INK_TEST_PAGE_ID`,
or `ONENOTE_INK_TEST_NOTEBOOK` + `ONENOTE_INK_TEST_SECTION` + `ONENOTE_INK_TEST_PAGE_TITLE`
(plus `ONENOTE_INK_TEST_SECTION_GROUP` when the section sits in one). It renders the page's
ink before and after `append_to_page` and `update_page_title` and asserts the stroke count
and the PNG bytes are identical, because a PATCH that strips handwriting reports no error
and the page cannot be recovered. Two things make it worth the credential: the page must
carry real strokes, which only a person with a tablet can put there, and the write is
confirmed independently — the appended marker has to appear in the page HTML and the new
title has to become visible — so a write that silently did nothing cannot pass as ink
preserved. It leaves the page changed: the marker paragraph stays, and the title is put
back only when the original could be read. It was run against the account on 2026-08-18
against a page carrying 5 strokes: both writes left the stroke count and the PNG bytes
identical. `npm test` runs it as a skip.

`test/page-layout.test.ts` covers the ink-clearance arithmetic, which is pure: it takes
page HTML and an ink bounding box and returns a plan, so the whole decision runs without a
Graph call. Its HTML is written in the shape Graph emits — absolutely positioned top-level
divs with `left`, `top` and `width` in px. What it cannot cover is `LINE_HEIGHT_PX`; no
endpoint reports a line height, so whether 19px per break lands the text just below the
strokes is only visible in OneNote. `test/fixtures/ink-below-text.inkml` is hand-authored
so its strokes fall inside a 48px/624px column and end at 466px, which is the case that
needs padding.

`test/oauth-router.test.ts` drives both metadata documents over a real HTTP server
through `createApp`, the way `test/mcp-server.test.ts` does. The paths are built from the
issuer URL rather than from a mount point, so a wrong mount produces a document that
reads correctly and a 404 at every URL it names — only a request can tell them apart. It
asserts the endpoints the metadata advertises answer something other than 404 or 405, and
that neither document contains the client secret or the signing key. What it cannot check
is whether Claude accepts the documents; nothing confirms that until a real connect
against the deployed URL.

`test/oauth-provider.test.ts` drives the whole Layer-1 flow — `GET /authorize`, the
consent POST, the token exchange, the refresh — over a real HTTP server through
`createApp`. Nothing in it is a unit test of a handler: what it checks is the status, the
OAuth error code in the body, and whether a `Location` header is present, and a direct
call to the provider bypasses the SDK middleware that produces all three. It
re-implements the token format at the bottom of the file rather than importing the
signer, so a change to how a payload is signed breaks the test; a test that signs with
the implementation's own function proves only that it agrees with itself. The filler
forms in the store-bound test are signed the same way rather than fetched, because
`/authorize` allows 100 requests per 15 minutes across every caller and one test would
spend the file's whole budget. What it cannot check is whether Claude accepts any of it;
that waits for a real connect against the deployed URL.

`test/server.test.ts` covers what `createApp` mounts, and the bearer gate is most of it.
Every assertion is a request over the wire, because what is being tested is the status and
the `WWW-Authenticate` header an unauthenticated caller gets, neither of which a direct
call to the middleware produces. The route-enumeration test reaches into Express's
internals: a mounted router's prefix cannot be read back off the stack in Express 5 — the
path string is compiled into a matcher and dropped — so it wraps `Router.prototype.use`
for the duration of one `createApp` call to record the prefixes, and reads leaf paths off
`layer.route`. That is the cost of the property it asserts, which is that a route added in
a later issue shows up without anyone remembering to add it to a list. It signs its own
tokens rather than importing the signer, for the reason `test/oauth-provider.test.ts`
gives.

`test/name-lookup.test.ts` drives the resolver through a fake `LookupStructure` that
counts calls, because what this module is for is what it does not do: the common path is
one `getExpandedTree` and no container walk. Its fixture nests a section group inside a
section group, which the expanded response cannot reach — that is the only thing
exercising the fallback walk, because the real account has no nesting at that level.

`test/keepalive.test.ts` drives the route through its own router with a fake target,
because the success path calls `refresh()` and the real one reaches Firestore and Entra.
`createApp` is used for the two facts a fake router cannot show: that the route is absent
when no secret is configured, and that a request carrying no Authorization header gets the
route's own 401 rather than the bearer gate's challenge. What no test covers is whether a
forced refresh actually slides Microsoft's window — that is a property of Entra, and
nothing confirms it until an operator watches the scheduler job run for longer than the
window.

`test/mirror-reader.test.ts` and `test/mirror-tools.test.ts` cover the read path, and
they are separate from `test/structure-tools.test.ts` and `test/page-tools.test.ts` on
purpose: those two build their tools with no mirror argument, and every assertion in them
still holds unchanged. That is the property worth protecting — an absent mirror means
"always Graph", which is what makes `MIRROR_READ_ENABLED=false` a rollback rather than a
code path with its own bugs. Almost every assertion in the pair is about a **miss**,
because a miss is what sends a read to Graph and every way of getting it wrong produces a
confident wrong answer rather than an error. The worst of them, and the one a refactor is
most likely to reintroduce: a page whose stored ink object has gone must be a miss for
the whole page, because answering `ink: null` there says "this page has no handwriting",
which a model cannot detect and which silently drops the only copy of what the page says.

`test/mirror-sync.test.ts` drives the sync through fakes that count their calls, with no
`fetch` anywhere in it — every URL the sync builds is already asserted in
`test/graph-structure.test.ts` against an exact-URL fake. What it asserts is the
algorithm, and most of the properties worth having are about what does *not* happen. In
order of how much damage the absence would do: a failed sweep enumeration deletes
**nothing**, because an auth failure or a 500 would otherwise empty the mirror one
section at a time; a section whose page listing failed keeps its old watermark, so the
next run retries it rather than skipping every page it never reached; a budget-exhausted
run keeps the advances it earned, because the backfill is five hours of slices; and an
unchanged content hash writes nothing and renders no ink, which is what makes the hour of
watermark overlap nearly free. What no test there covers is whether Graph's timestamps
behave as the algorithm assumes — that a page create, edit and delete each move the
section's `lastModifiedDateTime` is measured in `api-overview.md`, not checked here.

`test/sync-route.test.ts` is `test/keepalive.test.ts`'s shape: its own router with a fake
target for everything the route decides, plus `createApp` for the two facts a fake router
cannot show. It asserts that each of the three paths reaches its own mode and no other,
which is the point of there being three paths — see the convention below.

`test/mirror-schema.test.ts` covers `src/mirror-schema.ts`, which touches no backend.
That split is forced rather than stylistic: `src/mirror-store.ts` and
`src/mirror-blobs.ts` have **no automated test at all** and cannot get one on this
machine. They need a Firestore backend and a GCS bucket; the emulator is not installed
here, for the reason the `test/token-cache.test.ts` paragraph below gives; and an
in-memory fake is ruled out there for reasons that apply doubly to the mirror — the 1 MiB
document limit, the 1500-byte document-id limit, whether a query runs at all without its
composite index, transaction behaviour under a contended lease, and
`FieldValue.serverTimestamp()` are every one of them properties of Firestore, and a fake
would assert the fake. So the rule is that those two files hold only calls, and anything
a person could get wrong lives in `src/mirror-schema.ts` where a plain unit test reaches
it. If you find yourself adding a branch to `mirror-store.ts`, that branch belongs in
`mirror-schema.ts`. What no test covers even there is whether Firestore accepts what the
schema produces: the id rules and both limits are read off Google's documentation rather
than measured, and only a live write settles them.

`test/graph-decode.test.ts` covers `src/graph-decode.ts`, which has no network in it —
every function takes an already-parsed value and the URL it came from. Before the split
these were reachable only through a fake `fetch` keyed by exact URL, so each decode
assertion paid for a routing table. The last test in the file is the one that is easiest
to lose and worth the most: eight decoders are handed a body carrying a plausible page
name in the position they reject, and every thrown message is checked for not containing
it. What no test covers is whether Graph's bodies still have the shape these decoders
assume; that is asserted in `test/graph-structure.test.ts` against hand-written fixtures
and confirmed only by a live run.

`test/firestore.test.ts` covers the memoisation in `src/firestore.ts` and nothing else.
It constructs the real `@google-cloud/firestore` client, which needs no credential and no
backend because the client connects lazily — the same property `test/tools.test.ts` leans
on, so if it stops being true both files fail together. Everything the client then does
is untested, for the reason the `test/token-cache.test.ts` paragraph gives.

`test/tools.test.ts` covers the registry, and it constructs the real MSAL and Firestore
clients while doing so. Neither opens a connection until a token is asked for, so it
needs no credential and no backend; if that ever stops being true, this test is where it
will show up first.

`test/token-cache.test.ts` covers `readCache`, `isEmptyCache` and
`overwriteWouldEmptyCache`, all pure functions that run without a backend. The write
guard's fixtures name `Account` and `RefreshToken` because that is what MSAL writes today,
but every assertion in them holds with those keys renamed — that is the property keeping
the guard from inverting when MSAL changes its format, so do not rewrite them to assert on
the names. `beforeCacheAccess`,
`afterCacheAccess`, the transaction inside `afterCacheAccess`, and
`createFirestoreTokenCachePlugin` have no automated test at all. They need a Firestore
backend, and the emulator is not installed here. Installing it takes `sudo apt-get
install google-cloud-cli-firestore-emulator`, because this machine's `gcloud` is the
Debian package and its component manager is disabled, so `gcloud components install
cloud-firestore-emulator` is refused. Nothing stands in for those tests; the gap is
open, and closing it means an operator driving the plugin against an emulator by hand.
Do not close it by adding an in-memory Firestore fake: the behaviour at stake is transaction retry under contention
and `FieldValue.serverTimestamp()`, and a fake would assert the fake's behaviour rather
than Firestore's.

`scripts/test/` is deliberately separate from `test/`: those are bash tests driven by
`scripts/test/run.sh`. `npm test` does not run them. `bootstrap.test.sh` exercises
`gcp-bootstrap.sh` against a stub `gcloud` on `PATH`. `docker.test.sh` builds the image
and starts the container, so it needs a Docker daemon and takes about a minute;
`run.sh` runs it only when `RUN_DOCKER_TESTS=1` is set and prints an explicit skip line
otherwise. Both suites still get `bash -n` and `shellcheck` on every run.

## Commands

| Command | What it does |
|---|---|
| `npm ci` | Install from the lockfile |
| `npm run build` | `tsc -p tsconfig.build.json` → `dist/` |
| `npm run typecheck` | `tsc -p tsconfig.json` — covers `src/` **and** `test/`, emits nothing |
| `npm test` | `node --test "test/**/*.test.ts"` — runs TypeScript source directly, no build needed |
| `npm run dev` | `node --watch src/index.ts` |
| `npm start` | `node dist/index.js` — run `npm run build` first |
| `npm run bootstrap` | `node src/bootstrap.ts` |

There is no ESLint. `npm run typecheck` is the static-analysis gate; run it before every
commit. It catches the class of error a linter would here.

## Conventions

**Import specifiers carry the `.ts` extension.** Write `import { loadConfig } from
'./config.ts'`, not `'./config.js'`. `rewriteRelativeImportExtensions` turns it into
`.js` on emit, so the same source both runs under `node --test` without a build and
compiles to working output in `dist/`. Test files reach into source as
`'../src/config.ts'`.

**The source must be erasable.** `npm test` relies on Node's native type stripping,
which removes types but cannot transform syntax. So: no `enum`, no `namespace`, no
constructor parameter properties, and type-only imports written `import type`.
`erasableSyntaxOnly` and `verbatimModuleSyntax` enforce this at typecheck time — you
will get `TS1294` rather than a confusing runtime failure.

**Two tsconfigs, on purpose.** `tsconfig.json` is the broad one: it includes `test/`, so
the editor and `npm run typecheck` check test files, and it sets `noEmit` so it cannot
produce output by accident. `tsconfig.build.json` narrows to `src/` and emits. Keeping
them split is what stops compiled tests landing in `dist/` and then in the container
image built by issue #6.

**Configuration is grouped, and read in exactly one place.** `loadConfig(groups, env)`
in `src/config.ts` validates only the groups the caller asks for: `graph`, `firestore`,
`firestore-explicit`, `oauth`, `server`, `mirror`. The server asks for everything except
`firestore-explicit`; `bootstrap.ts` asks for `graph` and `firestore-explicit` only, so
the operator running it locally is never made to hold `MCP_OAUTH_CLIENT_SECRET`. Add a
new variable to the `SPECS` table, not to a new `process.env` read somewhere else.

**The `mirror` group is entirely optional, and that is what makes it a rollback switch.**
Every `MIRROR_*` variable defaults or may be absent, so a service deployed with none of
them set behaves exactly as one built before the group existed. `MIRROR_READ_ENABLED` is
set back to `false` and redeployed to turn the whole page mirror off — no code change, no
data migration, and the tool modules are identical under both settings. The server asks
for the group unconditionally even while nothing reads it, so a malformed `MIRROR_*` value
fails at startup beside every other configuration problem rather than at first use.

**`MIRROR_BUCKET` is the one cross-field rule in `loadConfig`, and it is deliberate.** A
`VarSpec` says required-or-not per variable and cannot say "required when another is
present". A sync has nowhere to put a rendered ink PNG without a bucket and a mirror read
has nowhere to fetch one from, so the `mirror` block checks the pair itself and throws
`ConfigError(['MIRROR_BUCKET'], [], purposes)`. Without it the failure surfaces hours into
a backfill at the first object write. If a second rule of this kind is ever needed, the
right move is probably a `requires` field on `VarSpec` rather than a second hand-written
block.

**`checkBoolean` accepts only `true` and `false`.** Not `1`, not `yes`, not `on`. A
lenient parser would also have to decide what `0` and `off` mean, and at that point a typo
like `ture` reads as `false` and switches the mirror off with nothing to say so. The
comparison in the `mirror` block lowercases to match, because a case-sensitive read of a
case-insensitive validator would let `True` pass validation and then evaluate false —
which is the same silent inversion from the other direction.

**The bootstrap CLI requires the two Firestore variables the server defaults.** That is
the whole difference between the `firestore` and `firestore-explicit` groups:
`FIRESTORE_CACHE_DOC` falls back to `tokencache/msal` and `GOOGLE_CLOUD_PROJECT` may be
absent on Cloud Run, where the metadata server supplies it. The CLI runs against whatever
project the operator's Application Default Credentials point at, so letting either name
default would seed a real Firestore document somewhere the deployed service does not read
and still print a success line. Ask for one group or the other, never both.

**`ConfigError` is given the descriptions to print, not asked to look them up.** A
variable name appears in more than one group with a different reason for being required,
so `SPECS.find(byName)` would print the server's description in an error raised by the
bootstrap CLI. `loadConfig` collects the purposes of the specs it actually consulted and
passes that map to the constructor.

**Startup failures print one readable message and exit 1 — never a stack trace.** That
covers missing variables, malformed values, and bind failures (`EADDRINUSE`, `EACCES`).
`loadConfig` accumulates every problem before throwing, so one run reports the full list.
`exitOnConfigError` is the only place in `src/config.ts` allowed to call `process.exit`,
and only entrypoints call it.

**Never hardcode a port.** Cloud Run sets `PORT`. The only literal `8080` in `src/` is
the documented default in the `SPECS` table in `config.ts`.

**`createApp()` does not listen.** Construction is separate from binding so tests can
drive routes on an ephemeral port. Mount new routes there, not in `index.ts`.

**The health endpoint answers on `/healthz` and on `/health`, and the second one is not
redundant.** Measured against the deployed service on 2026-08-19: Google's frontend
answers `https://<service>.run.app/healthz` with its own 404 page and the request never
reaches Cloud Run — it appears in no request log — while `/health`, `/healthz2` and
`/Healthz` all arrive. So `/healthz` works for Cloud Run's own probes, which reach the
container directly, and `/health` is the only one an external check can call.
`HEALTH_PATHS` in `src/server.ts` is exported so the fail-closed route test enumerates
both rather than someone remembering to add the second.

**The health response reports no configuration.** The service deploys
`--allow-unauthenticated`, so its body is public. A test asserts no key or value in that
body matches the stub config's secrets; do not add config echo to it.

**The container's runtime base is Debian `-slim`, never Alpine.**
`@resvg/resvg-js-linux-x64-gnu` declares `libc: ["glibc"]` and ships no musl prebuild.
On an Alpine base the optional package either fails to install or fails at `require`
time, and handwriting rendering is the point of this service.

**The runtime stage installs its own dependencies; it does not copy `node_modules`
forward.** `npm ci --omit=dev` runs again from the lockfile in the runtime stage. Copying
`/app/node_modules` from the build stage would ship `typescript` and `@types/node` into
the deployed image, and it would test nothing about whether the platform-specific resvg
package still resolves without dev dependencies. `scripts/test/docker.test.sh` asserts
both — that the `.node` binary is present and that `node_modules/typescript` is not.

**`.dockerignore` is a denylist.** A new source directory is included in the build
context by default rather than silently dropped. The entries that mirror `.gitignore`
(`.env*`, `*.token-cache.json`, `.msal-cache*`) are there so a local operator's
credentials cannot reach an image layer.

**The bootstrap CLI shares the server's client id, authority, scopes, and cache
plugin.** It imports `GRAPH_SCOPES` from `src/graph-auth.ts` and
`createFirestoreTokenCachePlugin` from `src/token-cache.ts` rather than restating either.
MSAL keys cached tokens by client id and by scope string, so a short scope name in one
place and the fully-qualified form in the other produces a cache that looks valid and
that the server's silent acquisition misses. The CLI serializes nothing itself; the write
happens inside `acquireTokenByDeviceCode`, through the plugin's `afterCacheAccess`.

**The bootstrap CLI prints counts, not names, and one line about the tenant.** Its Graph
probe is `GET /me/onenote/notebooks?$select=id`, so notebook display names never enter
the response body. The confirmation names the Firestore project, the document path, and
the account's home tenant — the tenant because a device-code sign-in into the wrong
directory succeeds with no error, and that line is the only thing that catches it. The
CLI's own output therefore carries a tenant id and must not be pasted into an issue, a
pull request, or a workflow log; it says so itself.

**The ink pipeline is a port, not a design.** `src/ink.ts` reproduces the recon script in
Appendix A of `project-spec.md`, which is validated against a real 440-stroke page. Four
things in it look arbitrary and are not: namespaces are stripped because Graph emits
`inkml:ink`; channel order is read from `<traceFormat>` because this account's points are
X, Y, F and F is pressure; coordinates are himetric and become px at 96 dpi; and traces
are collected from every `<ink>` root and every nesting level of `<traceGroup>`. Changing
any of the four produces a picture that is wrong rather than an error.

**A page with no ink is `null`, not an error.** Most pages are typed. `renderInk` returns
null whenever no `<ink>` root yields a stroke, including when a root parses to nothing:
`fast-xml-parser` accepts unclosed and mismatched tags without throwing, so "damaged" and
"holds no traces" cannot be told apart, and guessing would turn ordinary pages into failed
requests. The two errors that do exist are `InkParseError` for trace groups nested past 50
levels and `InkRenderError` for a document resvg rejects.

**No ink error message quotes the document.** Stroke coordinates are the user's
handwriting and this repository's output can reach a public log. The parse failure path
discards the thrown parser error for that reason: `fast-xml-parser` puts the offending
markup in its message.

**One fetch serves both halves of a page.** `GraphPageContent.fetchRaw` is the only
`includeInkML=true` call, and it returns the split parts. `fetchContent` turns that one
response into `{ html, ink }` — the trimmed HTML and the rendered PNG — and is what the
`get_page_content` tool wraps. A page with no ink comes back as plain
`text/html` rather than as `multipart/mixed`, so `pageHtml` falls back to the whole body
when there are no parts; dropping it because there are no parts to search would lose the
text of every typed page.

**The content URL asks for `includeIDs=true` as well.** Graph emits the generated `id`
attributes only when asked, so without it a page authored in the OneNote client comes
back with no id and no `data-id` anywhere and nothing on it can be targeted by a PATCH.
The two parameters compose — measured 2026-08-18: `includeIDs=true&includeInkML=true` is
200, `multipart/mixed`, ids present in the HTML part — so `src/multipart.ts` is
unaffected. Generated ids change whenever the page is updated, so anything that targets
one has to read it in the same operation that uses it; they cannot be cached between
calls.

**The HTML trimmer rewrites tags and never touches text.** `trimPageHtml` in
`src/page-html.ts` parses to a small tree, filters attributes, drops comments, and
removes or unwraps elements. Text nodes are copied to the output verbatim, so no entity
and no character of the user's writing can be lost to a parser's escaping rules. The one
exception is `collapseBlankLines`, which closes the gaps removed elements leave behind.
`test/page-html.test.ts` asserts the word sequence of the fixture is identical before and
after.

**Three things survive the trim that look like noise.** `id` and `data-id`, because the
PATCH write model in issue #18 targets an element by id. The `position`, `left`, `top`,
`width`, and `height` style declarations, because page content is laid out absolutely in
px at 96 dpi and that is the coordinate space `src/ink.ts` renders strokes into — nothing
needs it today and it cannot be recovered later. Empty `<td>` elements, because dropping
one shifts every cell after it into the wrong column. An element carrying an id is kept
whatever it holds; what gets dropped is the wrapper whose only attribute was `style`.

**The trimmer's parser is tolerant, not strict.** An unmatched close tag is ignored, an
unclosed element is closed when its parent ends, and an unterminated `<` is text. A page
whose text is readable must not become a failed request over a stray tag. It also does
not decode entities, imply end tags, or treat `<pre>`, `<script>`, and `<style>` as raw
text; Graph emits none of those, and adding an HTML library for this would be a
dependency for one endpoint.

**Stateless Streamable HTTP, one server per request.** Every `POST /mcp` builds its own
SDK `Server` and its own `StreamableHTTPServerTransport`, answers, and closes both.
Nothing carries to the next request, and no session id is issued. `--max-instances=1`
does not make the instance permanent: a revision change replaces it, and a client holding
a session id against a dead instance has no way to recover. Omitting
`sessionIdGenerator` is what selects stateless mode — it is omitted rather than set to
`undefined` because `exactOptionalPropertyTypes` rejects the explicit `undefined` the
SDK's own example passes.

**SSE is refused in two places.** `enableJsonResponse: true` makes a POST answer with a
JSON body instead of opening a stream, and `GET /mcp` is answered 405 by the router
before the transport sees it. The second is not redundant: the SDK's stateless mode still
opens a standalone SSE stream on GET, and that open stream is what holds a Cloud Run
instance alive and bills for idle time. `DELETE` is 405 for the same reason it is
meaningless — there is no session to close.

**The MCP surface is built on the SDK's low-level `Server`, not on `McpServer`.**
`McpServer` installs its `tools/list` and `tools/call` handlers as a side effect of the
first `registerTool` call, so a server with no tools answers `tools/list` with "method not
found" rather than an empty list, and issue #14 has to ship the empty list working. The
low-level class also puts the `tools/call` error mapping in one place. The cost is that
`ToolDefinition` carries hand-written JSON Schema and receives unvalidated arguments,
which is what the `requiredString` / `optionalString` / `optionalInteger` helpers in
`src/mcp-tools.ts` exist for. Use them rather than reaching into the arguments object.

**The tool registry is `src/tools.ts`, not `src/mcp-tools.ts`.** A tool module imports
`ToolDefinition` and the argument helpers from `src/mcp-tools.ts`, so listing the modules
there as well would make the two files import each other. `src/mcp-tools.ts` holds the
contract, the error mapping, and the helpers; `src/tools.ts` builds the shared Graph
client and concatenates the lists each tool module returns. The order it concatenates in
is the order `tools/list` shows: browse, read, then write.

**A tool answers with one JSON text block, and every count it reports is real.**
`list_pages` says `moreAvailable` when Graph returned exactly `top` items, and
`search_pages` says how many sections it searched out of how many it found. A model
cannot tell a truncated search that matched nothing from a complete one, so a bounded
result that omitted the bound would be read as "no such page". `get_page_content` is the
one tool that answers with a second block, and it is an image block for the reason below.

**`get_page_content` returns the ink as an MCP image block, and that is the point.** The
calling model reads the handwriting with its own vision; there is no OCR service and no
handwriting API anywhere in this repository. Base64 inside the JSON text block would
reach the model as characters it cannot see, and a file path would name a container
filesystem no client can read — `test/page-tools.test.ts` asserts the PNG's base64 prefix
does not appear in the text block. A page with no ink yields one block and `inkImage:
null`; that is the normal answer for a typed page and not an error.

**All of a page's ink is one cropped image.** Ink and typed content are independent on
these pages, so the bounding box in `strokesToSvg` loses nothing worth keeping, and
nothing here splices ink fragments back into the HTML at the positions they came from.

**The ink PNG is shrunk by measurement, not by arithmetic.** `fitInkToByteBudget` in
`src/ink.ts` re-rasterises the same SVG at 0.75× the width until the PNG fits
`MAX_INK_PNG_BYTES`, because PNG size depends on how dense the strokes are and cannot be
predicted from the pixel count. It stops at `MIN_RENDER_WIDTH` even when that still does
not fit: an image too small to read is a better answer than a failed request, and the
result says which width it got so the model can tell "illegible handwriting" from
"rendered too small".

**Writing is a PATCH of change objects, and the shapes are measured, not guessed.**
`src/page-write.ts` sends `[{target, action, content}]` to
`PATCH /me/onenote/pages/{id}/content` and reads a 204 with no body. The three shapes and
every error code they can produce were measured against the live service on 2026-08-18 by
the spike in issue #17 and are recorded in `api-overview.md` under **Writing page
content**. The array is applied as a unit: one change naming a missing target fails the
whole request with 400 code 20120 and applies none of the others.

**`target: "body"` is the first outline, not the page, and `append_to_page` says so.** A
page authored in the OneNote client has sibling top-level divs, and an append lands at the
end of the first of them. Reaching another one means reading `?includeIDs=true` and
targeting that div's generated id, which is issue #27. The tool description states where
the content went, because a 204 does not.

**An append is padded off the handwriting, because nothing else can move it.** OneNote
fixes ink in a page-level layer and an outline grows downwards, so appending to a page
whose strokes hang below its text renders the text over the handwriting. Putting the new
content in its own outline below the ink is not possible: measured 2026-08-19, an
absolutely positioned div sent to `target: "body"` is flattened into the first outline
with its position dropped, `insert` beside an outline is 400 code 20135, and `replace` on
one is 20134 or 20141. Margins are normalised away and an empty paragraph is deleted, so
the only lever left is `<br>`, which survives verbatim. `src/page-layout.ts` reads the
outline's `top` and the ink's bounding box and decides how many breaks go in front of the
caller's fragment; `append_to_page` pays one extra Graph read per call for it, and the
result JSON reports the padding rather than hiding it.

**The padding is measured from an estimate of where the text ends, and marked so it
happens once.** No endpoint reports an outline's rendered height, so `estimateContentHeight`
counts block elements at deliberately low heights — `p`, `li` and `br` at 16px, `h1` at
24px, `tr` at 18px — plus the declared `height` of any `img`, which is the case a block
count gets worst. Every constant is chosen low on purpose: an estimate that is too small
pads a little too much, and one that is too large puts the new text back on the
handwriting. Wrapping is not modelled, which errs the same safe way. Measured against the
live page on 2026-08-19: 136px estimated where the real content was about 174px, so the
padding ran two blank lines long. The deeper of the estimate and the marker wins, so a
page whose content wraps is not padded twice. The fragment is wrapped in `<div data-id="ink-clearance-{px}">`, recording the page
position the content was brought down to — not the ink bottom, which reads as new ink on
the next call over rounding, and not on the padding itself, because the service discards a
`data-id` from an element holding only line breaks. `LINE_HEIGHT_PX = 19` and
`INK_CLEARANCE_MARGIN_PX = 12` are chosen numbers: nothing in the API reports a line
height, and only looking at the page in OneNote can confirm them.

**A page created here has one outline, deliberately.** `createPageHtml` omits
`data-absolute-enabled` from `<body>`, so Graph wraps the submission in a single
`<div data-id="_default">` and `body` then covers the whole page. That is what makes
`append_to_page` reach the bottom of a page this server created. Setting the attribute
would produce sibling outlines like a client-authored page, and appends would land in the
first one.

**The title is escaped on create and verbatim on rename.** `createPage` puts it in a
`<title>` element inside a document Graph parses, so it is escaped; `updatePageTitle`
sends it as PATCH `content`, which is stored character for character — the spike produced
a page actually titled `<p>x</p>`. Escaping the second would put `&amp;` in a title, and
not escaping the first would let a `<` open a tag.

**The write tools refuse before they spend a request.** `fragmentArgument` rejects a
fragment carrying `<html>`, `<head>`, `<body>`, `<title>`, `<meta>`, `<base>`, `<link>`,
`<script>` or `<style>`, and `titleArgument` rejects a title holding a complete tag. Both
throw `ToolInputError` and neither reaches Graph — on `create_page` a wrong request that
went through would leave a page behind for someone to find and delete. What is *not*
checked is well-formedness: `<p>unclosed` returns 204 and the service closes the tag, so a
strict parser here would refuse content that works. A bare `<` that opens no tag is left
alone in a title, because `if x <y then` is a legal title and this tool is the only way to
set one.

**Container names are matched by a ladder, and the result says which rung answered.**
`matchNodes` in `src/name-lookup.ts` tries exact and case-insensitive, then the same
comparison against the candidate with a leading ordering prefix removed, then a
case-insensitive substring. The middle rung is the one this account needs: its section
groups are named `062 - February` and a caller knows the month, not the number. A rung is
only tried when the one above it matched nothing, so an exact match can never lose to a
looser one — the test for that is two groups where only one strips to `February`.
`matchedBy` in every `_by_name` result names the rung, so a caller that asked for
`February` can see it got `062 - February` and why. Page titles do not use the ladder;
they are matched in full, ignoring case, and `search_pages` is the substring tool.

**A `_by_name` tool refuses rather than guesses.** A name matching nothing is a
`NameLookupError` listing the sibling names, never an empty result: a caller cannot tell
an empty section from a section that does not exist. A name matching more than once on
the same rung is a `NameLookupError` carrying the candidates. `sectionGroupName` omitted
means the section is a direct child of the notebook, not "search everywhere".

**`list_pages_by_name` is the answer to "which page is it?".** It returns every title in
a named section with its page id, so a model reads the titles, picks one, and passes that
id straight to `get_page_content`. Nothing else is needed in between, and the tool
description says so — a caller that re-resolved the name would pay for the lookup twice.

**The two `_by_name` writing tools resolve names after they check the content.**
`create_page_by_name` and `append_to_page_by_name` are `create_page` and `append_to_page`
with the section named rather than identified, over the same `resolveSection` the
browsing `_by_name` tools use, so writing to a known place costs one call instead of a
`list_notebooks` → `list_sections` walk followed by a write. The title and the fragment
are validated before the lookup runs: a refused argument costs neither the Graph request
that resolves the names nor the one that would write. A container name that matches
nothing or matches twice is a `NameLookupError` and nothing is written — a write tool
that guessed a section would put content in the wrong notebook and nothing would say so.
`resolvedPayload` lives in `src/name-lookup.ts` rather than in either tool module,
because the browsing tools and these two answer with it and a second copy would drift.

**A page title is not matched by the container ladder, and `append_to_page_by_name` says
so in its error.** Graph compares it in full and case-insensitively across the section,
which is the same comparison `find_page_by_name` uses. Zero matches and more than one are
both `NameLookupError`, because this is a write: `find_page_by_name` can answer with an
empty list and let the caller decide, and an append cannot. `NameLookupError` takes a
`matching` argument for that — the container message tells the caller a leading number is
stripped, which is true of a section group and false of a page title, so telling an
`append_to_page_by_name` caller to drop one would send it to do something that changes
nothing. The page-title message names `list_pages_by_name` and `search_pages` instead,
and lists no candidates: getting them costs a second request, and the tool that lists them
is one call away.

**Both appends go through one `append` helper, and that is deliberate.** It reads the
page, plans the ink clearance, writes, and builds the result. A second copy of that in
`append_to_page_by_name` that skipped the read would write across someone's handwriting,
return 204, and report success — the failure is only visible in OneNote.

**The fallback below a named section group is one filtered request, not a walk.**
`getExpandedTree` reaches a notebook's sections and one level of section group, because
Graph caps `$expand` nesting at two levels — a third answers 400, and so does
`$levels=max`. A section nested deeper is therefore absent from that response rather than
known to be absent, and `findSectionsByName` settles it in one request at any depth by
asking for sections account-wide with their parents expanded. It runs only when the
expanded tree came back empty-handed, and `deepSearchUsed` in the result says when it did.

**Let Graph do the comparisons it will do, and know which ones it refuses.** Page titles
are matched by `tolower(title) eq '…'` on the section's pages, so nothing bounds how many
pages a section may hold before a match could be missed. Sections account-wide cannot be
matched that way: `tolower(displayName) eq '…'` there answers 500 with code 19999, while
`contains(tolower(displayName), '…')` answers normally, so `findSectionsByName` asks for
the substring and applies the full-name rule itself. The table of what each endpoint
accepts is in `api-overview.md`.

**`list_sections` returns sections and section groups in one tagged list.** Graph exposes
them as two relationships and the caller has to recurse through both — this account is a
notebook per year with a section group per month — so a result that carried only sections
would make a caller guess that the second relationship exists.

**An unscoped `search_pages` walks sections, and the walk is bounded twice.** At most
`MAX_SECTIONS_SEARCHED` sections and `SEARCH_TIME_BUDGET_MS` of wall clock, both in
`src/page-search.ts`. The account-wide page list is not an option — it is the endpoint
that fails with 20266 — so there is nothing cheaper to fall back on. The time budget is
checked before each section is fetched rather than during, so an overrun costs one round
trip. A failure listing one section aborts the whole search: an expired refresh token
fails every section identically, and returning "no matches" for it would be an answer.

**Layer-1 OAuth is the SDK's router, mounted at the application root.**
`src/oauth-router.ts` mounts `mcpAuthRouter`, which serves both metadata documents,
`/authorize` and `/token`. It builds every path from the issuer URL rather than from a
mount point, so it cannot go behind a prefix. `MCP_PUBLIC_URL` is the issuer, and the
`resource` identifier is that plus `/mcp` — nothing on Cloud Run tells the process what
URL it is reached at, and a value read from the `Host` header is whatever the caller
sent. Protected-resource metadata is served only at
`/.well-known/oauth-protected-resource/mcp`; the bare path is a 404, measured against SDK
1.30.0.

**Both consent responses carry `securityHeaders`, and the CSP has no `form-action`.**
`no-store` because one holds the signed authorization request and the other is a redirect
carrying a code; `no-referrer` because the form posts from the `/authorize` URL, which has
`state` and the PKCE challenge in its query string. `form-action` is the directive that
looks obviously right and would break the flow: browsers have disagreed about whether it
is checked against a redirect target, and the consent POST answers with a redirect to
claude.ai.

**`POST /consent` has a rate limit of its own, set above `MAX_PENDING_CODES`.** It is
mounted ahead of the SDK's `/authorize` limiter on purpose, and a rendered form stays
postable for `CONSENT_TTL_MS`, so one trip through `/authorize` yields a field that can be
replayed for ten minutes. The limit is 200 rather than something near 100 so that the code
store's own eviction — the bound whose behaviour is specified and tested — is what a burst
runs into first, and the limiter is only a backstop.

**The consent screen, the token format and the code store are `src/oauth-provider.ts`,
and `src/server.ts` wires the two together.** `oauthRouter` takes the provider as an
argument and does not construct it, so the dependency runs one way: the provider imports
the clients store and the resource URL from `src/oauth-router.ts`, and nothing imports
back. The provider carries a `consentRouter`, which the mount registers ahead of
`mcpAuthRouter` — the SDK's `/authorize` router renders nothing and owns no route that
resumes the flow after the click.

**A token is an HMAC over a compact payload, and no store is consulted to verify one.**
`base64url(JSON)` plus `base64url(HMAC-SHA256)` under `MCP_TOKEN_SIGNING_KEY`, one-letter
field names, `node:crypto` and no new dependency — nothing outside this server ever reads
one. Access tokens last an hour, refresh tokens 30 days, and the payload carries the kind,
the client id, the audience, the scopes and the expiry. Stateless is what makes a Cloud
Run revision replacement invisible to a connected client: an in-memory token store would
force a reconnect on every deploy. The audience is this server's own resource identifier
taken from configuration, never the `resource` parameter the request asked for.

**The refresh token slides, so the server runs unattended.** Every refresh mints a new
refresh token with a fresh 30-day expiry, which means the 30 days bound how long the
connector may sit idle, not how long the connection may live: Claude refreshes hourly on
its own, so a connector in regular use never returns to the consent screen. Issue #22
specified a non-sliding token and this deviates from it deliberately — unattended
operation is the point of the service, and a click every 30 days regardless of activity
defeats it.

**Sliding is not rotation, and the difference is asserted rather than implied.** The
refresh token handed in stays valid until the expiry stamped inside it; nothing here can
invalidate it, because a stateless token has no record to mark as spent. Real rotation
needs a store, and a store is what would make a revision replacement force a reconnect.
Rotation is required for *public* clients in any case, and the configured client secret
makes this a confidential one. `test/oauth-provider.test.ts` checks that both the renewed
token and the one it replaced still work, so the property cannot change unnoticed.

**There is no `revokeToken`, for the same reason.** No record of a stateless token exists
to delete, and the SDK advertises `revocation_endpoint` only when the provider implements
it. The operator's lever is rotating `MCP_TOKEN_SIGNING_KEY`, which invalidates every
outstanding access token, refresh token and consent form at once. Every refresh failure is
`invalid_grant` — not `invalid_request`, not a custom code — because that is what Claude
keys its re-authentication on.

**The authorization parameters cross the consent screen signed, and the codes stay in
memory.** The client id, redirect URI, PKCE challenge, state and scopes ride the form in
one signed hidden field, so an instance replacement between rendering and clicking does
not break the flow and the form cannot be edited. A field that fails to verify is a 400
with no redirect and no code minted — the redirect URI is part of what failed to verify,
so there is nowhere trustworthy to send the caller. Codes are a `Map`, single-use, 60
seconds, capped at 100 pending; the cap bounds what a burst against `/consent` can cost,
and losing a code costs a retry of the consent click, which is the one moment a human is
already present. The `--max-instances=1` assumption is what makes an in-memory code store
work at all.

**The client secret is compared with `!==`, and that is accepted.** The comparison lives
inside the SDK's `authenticateClient` middleware and cannot be replaced without replacing
the whole token router. The channel is a string comparison behind TLS, reachable only
from the public internet, under the token endpoint's 50-request rate limit. The signature
comparisons in `src/oauth-provider.ts` are `timingSafeEqual`, because those it owns.

**Two things about the mount are load-bearing and look like details.**
`scopesSupported` lists `offline_access`, which is the switch Claude reads to decide
whether to ask for a refresh token; without it the operator re-consents whenever an
access token expires. The clients store has no `registerClient`, which is what keeps
`registration_endpoint` out of the metadata and Dynamic Client Registration out of the
picture — the client id and secret are configured instead.

**`/.well-known/oauth-authorization-server` is served ahead of the SDK's own copy.** The
SDK advertises `none` in `token_endpoint_auth_methods_supported` unconditionally, which
is untrue of a server whose one client record carries a secret. The override is the SDK's
`createOAuthMetadata` output with that one key replaced, served by the SDK's
`metadataHandler`, so the two documents cannot drift and the CORS and method handling
stay identical.

**The OAuth rate limiter keys on a constant.** `trust proxy` is true, so
`express-rate-limit`'s default key is a forgeable `X-Forwarded-For` value and the library
logs a stack trace on every `/authorize` and `/token` request. Behind Cloud Run every
caller shares one bucket anyway. The constant key says so, and the lockout it allows —
anyone who finds the URL can spend the token endpoint's 50 requests — is stated in
`project-spec.md` and accepted.

**`/mcp` is closed by the SDK's `requireBearerAuth`, and everything else is open by
necessity.** The middleware is mounted in `createApp` between `MCP_PATH` and the router,
with `resourceMetadataUrl` set to `protectedResourceMetadataUrl` — that option is what
puts `resource_metadata=` in the 401's `WWW-Authenticate` header, and without it the 401
is a dead end rather than a sign-in prompt. The exempt list is both health paths, both
`.well-known` documents, `/authorize`, `/consent` and `/token`: everything `mcpAuthRouter`
mounts is reached by a client that holds no token yet. No `requiredScopes` is passed,
because passing any makes the middleware enforce them and `offline_access` is about
refresh tokens rather than about what a caller may do.

**The audience is compared with `checkResourceAllowed`, not with `===`.** The SDK's own
comparison, by URL origin plus path prefix. The two strings come from different places —
the audience stamped into the token when it was minted, and the configured
`MCP_PUBLIC_URL` the running process holds — so a byte comparison would refuse a valid
token over a trailing slash or a host's case. There is no separate issuer check: the
audience *is* `MCP_PUBLIC_URL` plus the MCP path, so an `iss` field checked against the
same configuration value would be the same comparison twice. Issue #23 asks for both; this
is the deviation.

**A tool failure is an `isError` result; a protocol fault is a JSON-RPC error.** Every
error a tool throws goes through `toolErrorResult`, because an expired refresh token, a
page that is gone, and a document resvg rejects are normal outcomes the calling model is
meant to read and act on. Calling a tool that `tools/list` never offered is the one thing
that comes back as a JSON-RPC error. Nothing reaches the transport as an unhandled
rejection.

**No tool error quotes a body, and an unrecognised error quotes nothing.** A Graph
failure is reduced to its status and the `error.code` / `error.message` of the OData
body, which is where "20266, maximum sections exceeded" lives; a body that is not that
shape is dropped rather than pasted through. An error type this repository does not model
yields "the server hit an unexpected error" and nothing else — an arbitrary message may
carry a request body, and a tool result reaches a client that may log it.

**What a request log line may contain is fixed in `src/logging.ts`.** The HTTP verb, the
path, the status, the duration, the JSON-RPC method, and the tool name on a `tools/call`.
Not headers, not bodies, not tool arguments, not results, and not the query string — the
MCP auth spec forbids a token in a query string and issue #23 rejects one that arrives
anyway, so logging the query would put the rejected token in the log. The path is
captured when the request arrives, not on finish: Express rewrites `req.url` when a
request enters a mounted router.

**Never call the account-wide page list.** `GET /me/onenote/pages` fails with error
20266, "maximum sections exceeded", once the account has enough sections across all
notebooks — a notebook-per-year with a section-group-per-month reaches that. Page
listing is always scoped to `/me/onenote/sections/{id}/pages`. A test in
`test/graph-structure.test.ts` scans every file under `src/` for the path and fails on
it; it strips comments first, so prose mentions are fine, and it bans the path only when
no further segment follows it, because `/me/onenote/pages/{id}/content` is a different
endpoint that the page-content work needs.

**List calls follow `@odata.nextLink`.** Graph chooses its own page size and ignores a
`$top` larger than it, so one response is never proof that a collection is complete.
`collectValues` in `src/graph-structure.ts` follows the link verbatim — it carries
Graph's own paging cursor and cannot be rebuilt from parts. `listPagesInSection` stops
following once `top` items are in hand, so `top` is a result count and not a page size.
Both walks have a bound: 50 followed links, and 20 levels of section-group nesting.
Neither is a real structure Graph produces; each exists so a self-referencing response
ends as one named error rather than as an unbounded loop.

**Prefer `getExpandedTree` to `getFullTree`.** They answer nearly the same question and
cost 1 request against 195 — see the `Graph request budget` section. `getFullTree` is
still the only thing that reaches section groups nested inside section groups, so it is
not dead, but nothing new should call it without needing that.

**Structure traversal recurses.** A notebook holds sections and section groups, and a
section group holds further sections and section groups — the UI's "tab groups". Both
container kinds expose the same two child relationship names, which is why
`listSections` and `listSectionGroups` take a `containerKind` of `notebooks` or
`sectionGroups` rather than existing twice. A walk that stopped at a notebook's direct
children would miss most of this account's sections.

**Graph read failures split into two error types.** `GraphRequestError` is a non-2xx
response and carries the status and the body; error 20266 is only distinguishable from
any other 400 by that body text. `GraphResponseError` is a 2xx whose body is not the
shape the caller needs, or a listing that would not terminate. The split is about what a
caller can do: a status can be retried or mapped, and a malformed body cannot. Neither
message prints a notebook, section, or page name — those are user content, and this
repository's output can reach a public log.

**`forceRefresh` exists for the keepalive route and nothing else.** `acquireTokenSilent`
without it answers from MSAL's in-memory access token, so no request reaches Entra and
Microsoft's refresh token does not slide. That is fine for a tool call and useless for a
keepalive: the whole point of `GraphAuth.refresh()` is the side effect, a replacement
refresh token with a fresh inactivity window written back to Firestore. Do not add it to
the tool path — every forced refresh is a token-endpoint round trip and a Firestore write.

**The read branch lives in the tool modules, not behind an adapter.** An adapter
implementing `StructureClient` would return `PageSummary[]` with nowhere to say who
answered; reporting the source would need either state on an object `createTools` shares
across every request — wrong the moment two calls overlap — or a widening of every narrow
interface's return type, which would change every existing fake in three test files. And
"which source answered" is part of the tool's contract with the model, exactly like
`moreAvailable`, `stoppedEarly` and `deepSearchUsed`. `readSourced` in
`src/mirror-reader.ts` is the one place the branch is written; every covered tool calls it
and spreads `source` and `mirroredAt` into its JSON.

**Every mirror read answers `null` on a miss, and every miss means "ask Graph".** A page
the mirror does not hold, a write-marked-stale page, a Firestore outage, and a caller that
passed `useLiveData` all end in the same place. Refusing a tool call because a cache is
down would be strictly worse than the behaviour before the mirror existed, which is the
bar. Three misses look like hits if you write them carelessly: a page whose stored ink
object is gone (answering `ink: null` claims the page has no handwriting), a section group
whose `childGroupsKnown` is false (answering a short list that looks complete), and an
empty structure collection (answering "no notebooks" rather than "never synced").

**`list_notebooks` and unscoped `search_pages` cannot miss, which is why structure is
mirrored for the whole account.** Neither takes an argument the mirror could fail to find,
so both would answer confidently and partially from a mirror holding three notebooks out
of fifty-five. The tree read returns every notebook and section for the same one request,
so storing all of it is free; `mirrored` on each document records which have their *pages*
held. Unscoped `search_pages` still reports `notebooksSearched` against
`notebooksInAccount`, because its pages really are a subset — and it drops
`stoppedEarly` and the section counts, which describe a walk that did not happen.

**A `NameLookupError` from the mirror is a miss, not an answer.** The mirror's structure
equals Graph's expanded tree by construction but is only as fresh as the last sync, so a
section created ten minutes ago is absent and `resolveSection` would turn that into
"sectionName matched nothing" listing the wrong siblings — which reads to a model as "no
such section". Both `_by_name` reading tools retry the whole resolve-and-list against
Graph before reporting. The retry costs one request and only on a failure, and a name that
exists nowhere still raises.

**Every write resyncs its page immediately, and that costs one Graph request.** All five
writing tools call `resyncPage` after a successful write — including `create_page`, whose
page the mirror has never seen and which would otherwise be a miss until the next
scheduled run. The alternative, marking the page stale and letting the next sync repair
it, leaves every read falling through to Graph for up to a whole poll interval, which in
the middle of a conversation is the window that matters most.

**A resync re-reads content and nothing else, and this is the measured reason.**
`api-overview.md` records that a PATCH *is* visible to the next content read — 3.7 seconds
including both round trips — while page *metadata* is weaker: `GET /pages/{id}?$select=title`
returned `""` for pages created seconds earlier. So the title travels in a hint from the
caller, which either just set it (`update_page_title`, `create_page`) or knows an append
cannot change it, and `lastModifiedDateTime` is stamped locally. Reading either back from
Graph here would trade a correct value for an unreliable one. Do not "improve" this by
fetching page metadata.

**`writePageFromRaw`'s short-circuit compares the title and the section, not just the
content hash, and both are load-bearing.** `update_page_title` changes a title and nothing
else, so a content-hash-only comparison short-circuited every rename: the mirror kept
serving the old title, which `find_page_by_name` and `search_pages` then matched against.
A page moved between sections is the same shape of miss — page ids are stable across a
move, so only the placement changed. `lastModifiedDateTime` is deliberately **not**
compared: it moves on every write, so including it would rewrite every page the watermark
overlap re-read and defeat the short-circuit entirely.

**An append or create that resyncs to `unchanged` is treated as a lost race, and the page
is marked stale.** Those two always change a page's content, so a resync that found
nothing to write did not read what was just written. Measured 2026-08-19, a PATCH is
visible to the next content read at 3.7 seconds — but that is one observation, and if the
read ever loses the race the stored copy is pre-write content marked `present`, which the
read path serves as current with nothing saying so. A stale marker sends the next read to
Graph, which cannot be wrong. A rename does not fall through this branch, because it
changes no content by design.

**There is one page writer, `writePageFromRaw`, shared by the sync and the resync.** A
second copy that skipped the ink render, or spilled to GCS at a different threshold, would
make a page's stored form depend on which path last touched it — and the difference would
only surface as a wrong answer to a model days later. `test/mirror-sync.test.ts` asserts
both paths build the same document from the same response.

**Two failure levels below a write, and neither fails the write.** A resync that throws
falls back to `markPageStale`, which makes the next read a miss — correct, just slower. If
that fails too, the event is logged and nothing else happens. The write has already
happened by then, so turning either into a reported error would send the caller to retry a
change that is already made. It is self-healing regardless: the write moved the page's
`lastModifiedDateTime`, so the next incremental run repairs whatever this could not.

**The write-sync is bound whenever a mirror exists, not only when reads are enabled.** A
mirror being filled by the sync while `MIRROR_READ_ENABLED` is false still holds copies a
write supersedes, and keeping them current then is what makes turning reads on later safe
rather than a race with whatever was written in between. It shares the Graph content
client with the read tools, so a resync passes through the same process-wide request gate
as everything else — a burst of writes cannot outrun the per-user rate limit through this
path.

**Cloud Run freezes an overrunning request rather than killing it, and the lease release
is written for that.** The service runs with CPU throttling — the default, and unset in
the deploy — so a request cut at the 300-second timeout does not finish: the process is
suspended and resumes whenever the next request arrives, possibly many minutes later. By
then the sync lease has expired on age and another run has taken it. An unconditional
`releaseLease` would then clear the *live* run's lease and let a third start alongside it,
both spending the same hourly Graph budget — the exact failure the lease exists to
prevent, reached by way of the lease itself. `releaseLease(heldSince)` therefore runs in a
transaction and clears nothing unless `runningSince` still matches the value this run
wrote. A superseded run logs `mirror-lease-superseded` and leaves the document alone.

**`POST /sync` is the page mirror's way in, on the same terms as `/keepalive` and for the
same reason.** Its own secret rather than the keepalive one, because the two reach
different things and a credential should reach one of them. Unmounted when
`MIRROR_SYNC_SECRET` is unset, so the path 404s and an operator learns the service is
unconfigured rather than that they mistyped a secret.

**The sync's mode is the path — `/sync`, `/sync/sweep`, `/sync/sweep/full` — not a body
field or a query parameter.** `src/logging.ts` records the method, the path and the
status, and deliberately records no query string and no body. A mode carried in either
would appear in no log line, and "which job ran, and did it answer?" is the first
question when the mirror looks wrong. A body would also need a JSON parser on a route
outside the bearer gate. A time-based rule would be worse than both: it makes behaviour
depend on the container's clock and removes the ability to force a sweep on demand, which
is the move the keepalive runbook documents as the way to prove a job works.

**A budget-exhausted sync run answers 200, not 503.** It is a normal outcome with
committed work behind it and a report attached. A 503 makes the scheduler retry
immediately and spend the next hour's Graph budget inside this one, which is the failure
the budget exists to prevent. 503 is reserved for a run that could not start; a held
lease is 409.

**`KEEPALIVE_PATH` and `SYNC_PATH` are a third category, neither exempt nor gated.** Both
are called by Cloud Scheduler, which has no browser and nowhere to keep a refresh token,
so neither can ever satisfy a bearer challenge. `test/server.test.ts` asserts they answer
401 with **no** `WWW-Authenticate` header — the discriminator between "this route refused
you" and "the bearer gate refused you" — and separately asserts both are registered
routes, so a `STUB_CONFIG` that stopped configuring their secrets could not make that
proof vacuous. It was vacuous until 2026-08-19: the stub set neither secret, so neither
route was mounted and the enumeration never saw them.

**`POST /keepalive` is authenticated by a shared secret, not by a bearer token.** A
scheduler cannot run the OAuth flow: no browser, nowhere to keep a refresh token. So
`MCP_KEEPALIVE_SECRET` is compared with `timingSafeEqual` before any work happens, and the
route sits outside the bearer gate in `createApp`. It is its own variable rather than the
Layer-1 client secret because that one can reach the whole MCP surface. The route is not
mounted at all when the variable is unset — a 404 tells an operator the service is
unconfigured, where a 401 reads as a mistyped secret. Only POST is answered, so a link
preview cannot spend a token exchange.

**One `GraphAuth` per process, built by `createGraphAuthFor` and passed to both
consumers.** `createTools` takes it as an argument rather than building it. Two MSAL
clients would each hold their own in-memory access token and their own view of the
Firestore cache, so a forced refresh through one would leave the other on a superseded
blob until its next read, and both would be writing the same document.

**The mirror's blobs are in GCS, and the InkML is stored beside the PNG.** Firestore
caps a document at 1 MiB and `MAX_INK_PNG_BYTES` is already 750 KB, so a rendered PNG
would fill most of a document alone. The InkML going to the bucket too is not redundant:
`MAX_INK_PNG_BYTES` is documented above as a chosen number rather than a measured one,
and `fitInkToByteBudget` shrinks a render by re-rasterising and measuring — so with only
the PNG kept, changing that budget would mean re-fetching every inked page from Graph,
hours of the request budget to correct a guess. It is the same protection the raw HTML
gets by being stored untrimmed. `src/mirror-blobs.ts` treats a 404 on read as a mirror
miss and a 404 on delete as success, and lets everything else propagate: a permission
failure that read as "not mirrored" would send every request to Graph and exhaust the
hourly budget with nothing saying why.

**`@google-cloud/storage` pins `uuid` through a `package.json` override.** The version
it depends on transitively carries GHSA-w5hq-g745-h8pq, a missing bounds check in uuid's
v3/v5/v6 generators when the caller supplies its own buffer. Nothing here or in `gaxios`
does that, so it is not reachable — but `npm audit` reporting clean is worth more than
the argument, because there is no audit gate in the deploy workflow to notice a future
finding that *is* reachable. Drop the override when the upstream dependency moves.

**One Firestore client per process too, and `src/firestore.ts` is the only place one is
constructed.** `firestoreFor(config)` memoises by project id; `createFirestoreTokenCachePlugin`
calls it rather than building its own, and anything else needing Firestore does the same.
Two clients against one database would open two gRPC channels and run two credential
refreshers for no benefit. It is memoised at module level rather than threaded through
constructors for the reason `PRODUCTION_GATE` is — a process-wide shared resource whose
wiring should not appear in every signature between the entrypoint and the leaf. The key
is the project id rather than a single instance so that the inferred client, the one
built when `GOOGLE_CLOUD_PROJECT` is absent and Cloud Run's metadata server supplies it,
can never be handed to a caller that named a project explicitly. Construction opens no
connection.

**Two events exist for alerting, and their field vocabulary is fixed.**
`graph-auth-failure` and `token-cache-write-refused`, written through `logEvent` in
`src/logging.ts` to stderr. They carry a reason from a fixed set and a document path from
configuration — never an account identifier, never a cause message, never user content.
They are separate from the request line because a log-based metric keyed on a status code
cannot tell a dead refresh token from a missing page. `setEventSink` exists for tests only.

**`@odata.nextLink` is checked against the Graph origin before it is followed.** It is the
one URL in this repository that comes out of a response body rather than being built, and
every request carries the Graph access token in a header. Graph is what writes these
links, so the check is not expected to fire; it is there so the token cannot leave
`https://graph.microsoft.com` whether or not that stays true. A link elsewhere is a
`GraphResponseError` that does not quote the link.

**The server never signs in interactively.** `acquireTokenByDeviceCode` belongs to
`src/bootstrap.ts` alone. The deployed service acquires silently from the Firestore
cache through `src/graph-auth.ts`. A device-code call in a request path would block on a
human who is not there, and Cloud Run would time the request out. A test in
`test/graph-auth.test.ts` scans every file under `src/` for that call.

**Graph auth failures are `GraphAuthError`, never a raw MSAL error.**
`acquireGraphToken` wraps both the `getTokenCache().getAllAccounts()` call and the
`acquireTokenSilent` call, so an undecodable cache, an empty cache, and a dead refresh
token all reach the caller as one error type whose message ends in `npm run bootstrap`.
Letting the MSAL error through instead surfaces to the operator as a bare 401 from
Graph, which does not say that a human has to re-run the CLI. The messages name the
document path and the underlying error, never `username` or `homeAccountId` — the first
is the user's UPN and the second embeds the tenant id.

**The token cache is one opaque string in one document.** `FirestoreTokenCachePlugin`
stores the output of MSAL's `serialize()` verbatim in the document's `cache` field,
alongside an `updatedAt` server timestamp and a `previousCache` field holding the blob
that write replaced. Do not parse that blob and do not split it across documents. MSAL
owns its structure and changes it between library versions.

**`isEmptyCache` is the one exception to not parsing it, and it reads no key name.** A
cache counts as empty when it parses to a JSON object whose every value is an empty
container — which is true of `{"Account":{},"RefreshToken":{},…}` whatever those keys are
called next version. Anything it does not recognise answers "not empty", so the guard
fails open: the only thing it decides is whether to refuse a write, and refusing every
write would strand the refresh token in memory. Do not make it look for `Account` or
`RefreshToken` by name; that is the change that would silently invert it.

**A write that would empty the document is refused, not attempted.**
`overwriteWouldEmptyCache` in `afterCacheAccess` is the guard. MSAL removes credentials
from its in-memory cache on some failures and `afterCacheAccess` runs inside MSAL's
`finally`, so an emptied serialization can reach the write while the stored blob is still
good — and that blob is the only copy of the refresh token, so replacing it costs a
device-code sign-in. The refusal logs `token-cache-write-refused` and throws nothing:
whatever emptied the cache is already on its way out, and a second error would bury it.

**A Firestore failure is `TokenCacheUnavailableError`, never a bare rejection.** Firestore
is read and written inside `acquireTokenSilent`, so without the distinct type a backend
outage reaches the operator as `silent-failed`, whose message says to re-run the bootstrap
CLI — sending a human to a browser to replace a credential that works. `acquireGraphToken`
walks the cause chain for it and reports `cache-unavailable`, the one reason with
`retryable: true` and the one message that does not name the CLI. Writes are retried
three times before it is raised.

**`afterCacheAccess` re-reads inside the transaction.** `--max-instances=1` does not
prevent two instances existing during a revision transition. The transaction re-reads the
document, and a stored blob differing from the one this instance last saw is fed back
through `context.tokenCache.deserialize` before `serialize()` is called, because MSAL's
`deserialize` merges into the in-memory cache rather than replacing it. Removing the
re-read turns an overlap into a lost refresh token, and the cache stays unusable until
`npm run bootstrap` is run again.

**The deploy workflow sets the service's whole environment, every time.**
`.github/workflows/deploy.yml` passes `env_vars_update_strategy: overwrite` to
`google-github-actions/deploy-cloudrun@v2`, whose default is `merge`. Under `merge`, a
variable deleted from the workflow keeps whatever the previous revision had, so the
running service and the file disagree and nothing says so. Under `overwrite` the list in
the workflow is the environment. Two names are deliberately absent from it: `PORT`, which
Cloud Run supplies and rejects as an input, and `GOOGLE_CLOUD_PROJECT`, which the metadata
server supplies. `MCP_KEEPALIVE_SECRET` is always listed, because an unset repository
secret interpolates to an empty string and `loadConfig` treats an empty value as unset —
which is the same thing as not listing it, without the conditional.

**`MCP_PUBLIC_URL` has three sources and the first deploy has none of them.** It is the
OAuth issuer and the audience of every access token, and no Cloud Run service exists to
have a URL before the first deploy. So the workflow takes the `MCP_PUBLIC_URL` repository
variable, falls back to `gcloud run services describe`, and on a first deploy uses
`https://placeholder.invalid` and then replaces it with `--update-env-vars` once the
service exists. The placeholder is well-formed enough to pass `checkPublicUrl`, which is
what lets the revision start and get a URL assigned. Setting the repository variable is
still the right end state — it is the only one of the three that survives putting a custom
domain in front of the service.

**Nothing in the workflow echoes a value.** The preflight step prints the names of
unconfigured variables and never their contents, and no step logs an environment. Two of
the eleven are secrets, and GitHub's log masking is a backstop rather than the control.

## Graph request budget

The OneNote endpoints throttle harder than the rest of Graph, and this repository has
already been throttled by its own structure walk. The limits below are Microsoft's
documented ones for a delegated app — which is what this service is — and the numbers
beside them were measured against the real account on 2026-08-18.

| Limit | Delegated app per user | What this repository does |
|---|---|---|
| Concurrent requests | 5 | `getFullTree()` opens 108 |
| Requests per minute | 120 | a paced walk at ~80/min saw no 429; the burst above failed immediately |
| Requests per hour | 400 | one `getFullTree()` costs 195 |

A throttled response is HTTP 429 with OData code `10007`, "the server is too busy". The
penalty outlasts a short backoff: after the burst above, five retries spanning three
minutes were all refused.

The design principles below come from Microsoft's own guidance
(<https://devblogs.microsoft.com/microsoft365dev/onenote-api-throttling-and-how-to-avoid-it/>)
and from what the measurements showed.

**Collapse a hierarchy walk into one request with `$expand`.** This is the article's
first best practice and it is not a micro-optimisation here.
`GET /me/onenote/notebooks?$expand=sections,sectionGroups($expand=sections)` returns
every notebook, its sections, its section groups, and those groups' sections in a single
call: measured at 2.9 seconds and 441 KB for 54 notebooks, 290 direct sections, 43
section groups and their 270 sections. The per-container walk that `getFullTree` does
instead costs `1 + 2 x containers`, which is 195 requests on the same account — half the
hourly budget for the same data. Note that the nested `$expand` reaches one level of
section group; a group nested inside a group still needs a follow-up request, so the
walk cannot be deleted, only avoided in the common case.

**Never fan out with an unbounded `Promise.all` over ids.** The concurrency limit is 5,
so any code that maps a list of containers, sections, or pages onto concurrent requests
has to cap what is in flight — 4 is the safe cap — and space request starts to stay
under 120 per minute. `getFullTree` and `#expandGroups` in `src/graph-structure.ts` both
violate this today; that is an open gap, not a decision.

**Retry only what is retryable, and inspect the OData code first.** The article's second
best practice: a 4xx other than 429 means the request itself is wrong and retrying it
burns quota to fail again. 429 and 503 are worth retrying after a wait; 400 with code
`20266`, 400 with code `20112` "invalid entity id", and 404 are not. `retryWait` in
`src/graph-throttle.ts` is the one place that decides, and every client runs through
`PRODUCTION_GATE`, so nothing here retries on its own.

**A retry longer than `MAX_RETRY_WAIT_MS` is declined, not shortened.** Graph decides how
long a 429 lasts and OneNote's answer can be minutes — the `Graph request budget` section
below records five retries spanning three minutes all refused after one burst. Shortening
that would hammer a service which has just asked for room. Honouring it verbatim is worse
in a different way: Cloud Run cuts a request at 300 seconds and the mirror sync budgets
240, and **both are checked before an operation starts rather than during**, so one
request sleeping for three minutes inside the gate blows through both and the run is
killed mid-flight. So the wait is declined: the caller sees the 429, the sync leaves that
section's watermark where it is, and the next scheduled run picks it up. Waiting is what
is given up on, not the work.

**One 500 is retried, and the rule is deliberately narrow.** Measured 2026-08-19 and
recorded in `api-overview.md`: every `$expand` on `/me/onenote/notebooks` answered 500
with OData code `19999` for seven minutes across 18 attempts and then recovered with no
change to the request, while un-expanded calls on the same collection answered 200
throughout. Without a retry that takes down `search_pages`, `find_page_by_name`,
`list_pages_by_name`, `create_page_by_name` and `append_to_page_by_name`, all of which go
through `getExpandedTree()`. So a 500 is retried only when the body carries code `19999`
**and** only on a GET. Both halves matter: `19999` is also what an account-wide
`/sections` request with no `$filter` answers every time, permanently, which is why the
retry rides the existing three-attempt cap rather than becoming an open-ended wait; and
`PATCH /pages/{id}/content` is not safe to repeat blindly, which is why `src/page-write.ts`
passing an explicit `method` excludes its writes by construction while
`src/graph-structure.ts` and `src/page-content.ts` default to GET and are covered. A body
that is not JSON, or is JSON of another shape, is left alone rather than guessed at.

**Treat the hourly 400 as the binding constraint on tool design, not the per-minute
120.** A tool that walks the account can run at most twice an hour before every later
call in that hour fails. That is what makes an unscoped `search_pages` expensive: it walks
the tree and then lists pages per section, and the account held 568 sections when last
counted on 2026-08-19. `MAX_SECTIONS_SEARCHED` is 60 and `SEARCH_TIME_BUDGET_MS` is 25
seconds, so one unscoped search costs up to 61 requests — a seventh of the hourly budget
for a sample of roughly a tenth of the account, and the result says so through
`stoppedEarly` and the counts. Removing that cost is most of the point of the Firestore
mirror in issue #30.

**Prefer a cached structure over re-reading it.** Notebooks, section groups and sections
change rarely; pages change often. Anything that needs the tree more than once in an
hour should hold it rather than fetch it, and no tool should walk the account as a side
effect of answering an unrelated question.

**Ask for the fields, not the objects, inside the expand clauses too.** `$select` is
already used on every listing in `src/graph-structure.ts`. It also works inside `$expand`
here, and it is worth 5.7x on the tree above — same 54 notebooks, 290 sections, 43
section groups and 270 group sections in every case:

| Request | Bytes |
|---|---|
| `?$expand=sections,sectionGroups($expand=sections)` | 441,012 |
| `?$expand=sections($select=id,displayName),sectionGroups($expand=sections($select=id,displayName))` | 158,070 |
| the same plus `$select=id,displayName` at the top level and on `sectionGroups` | 77,982 |

The third form is the one to use. Note the separator inside a clause that carries both:
`sectionGroups($select=id,displayName;$expand=sections(...))` — a semicolon, not a comma.

## Repository hygiene

This repository is public, and so are its issues, pull requests, and Actions logs.

Never commit, and never put in an issue or a workflow log: real page content, rendered
ink images, the Entra tenant name or ID, Firestore document contents, or the Layer-1
OAuth client secret. `.gitignore` covers `output/`, `.env*`, `*.token-cache.json`, and
`.msal-cache*`. Example output in `README.md` must come from a throwaway page with fake content.
The Azure client ID is *not* secret in this design (public client, device-code flow) and
is safe in config examples.

## Workflow

Issues are labelled by phase (`phase-1-skeleton` … `phase-5-deploy`) and by who does them
(`claude` or `manual`). Work a `claude` issue directly; no separate plan document is
written. The issue's own task list and acceptance section are the specification, and the
commit message is where the reasoning goes — what the
implementation revealed that the issue did not anticipate belongs there, not in a
separate document.
