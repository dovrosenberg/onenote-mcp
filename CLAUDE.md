# CLAUDE.md

Working notes for agents in this repository. `project-spec.md` is the authoritative
design document — read the relevant section of it before changing behaviour.

## Directory structure

```
onenote-mcp/
├── src/                       # TypeScript source, the only thing that ships│
├── test/                      # node --test suites; test/<name>.test.ts covers src/<name>.ts
│   ├── fixtures/              # hand-authored InkML and HTML; never a captured page dump
├── scripts/                   # operational shell scripts, not part of the built artifact
│   └── test/                  # bash tests, run by scripts/test/run.sh
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

`test/tools.test.ts` covers the registry, and it constructs the real MSAL and Firestore
clients while doing so. Neither opens a connection until a token is asked for, so it
needs no credential and no backend; if that ever stops being true, this test is where it
will show up first.

`test/token-cache.test.ts` covers `readCache` and nothing else. `readCache` is a pure
function over a document snapshot, so it runs without a backend. `beforeCacheAccess`,
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
`firestore-explicit`, `oauth`, `server`. The server asks for `graph`, `firestore`,
`oauth`, and `server`; `bootstrap.ts` asks for `graph` and `firestore-explicit` only, so
the operator running it locally is never made to hold `MCP_OAUTH_CLIENT_SECRET`. Add a
new variable to the `SPECS` table, not to a new `process.env` read somewhere else.

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

**`/healthz` reports no configuration.** The service deploys `--allow-unauthenticated`,
so its response body is public. A test asserts no key or value in that body matches the
stub config's secrets; do not add config echo to it.

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
client and concatenates the lists each tool module returns. `#18` appends there.

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
alongside an `updatedAt` server timestamp. Do not parse that blob and do not split it
across documents. MSAL owns its structure and changes it between library versions.

**`afterCacheAccess` re-reads inside the transaction.** `--max-instances=1` does not
prevent two instances existing during a revision transition. The transaction re-reads the
document, and a stored blob differing from the one this instance last saw is fed back
through `context.tokenCache.deserialize` before `serialize()` is called, because MSAL's
`deserialize` merges into the in-memory cache rather than replacing it. Removing the
re-read turns an overlap into a lost refresh token, and the cache stays unusable until
`npm run bootstrap` is run again.

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
