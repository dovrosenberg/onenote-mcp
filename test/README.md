# Test suites

What each suite covers, what it deliberately does not, and why. Read the entry for a
file before changing it — several of these tests are shaped the way they are to assert
something a more obvious shape would silently stop asserting.

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

`test/mirror-reader.test.ts` also covers the listing hold, and its three cases fail in
different directions: a held section is a miss for `listPagesInSection` *and* for a search
scoped to it *and* for an unscoped search; a hold in another notebook does not spoil a
search scoped away from it, which is the precision that keeps an unscoped `search_pages`
from paying 61 Graph requests over an unrelated write; and a hold older than
`LISTING_HOLD_EXPIRY_MS` answers from the mirror again, because a hold a dead process left
behind must not wedge a section permanently.

`test/read-sync.test.ts` covers the inline refresh the read tools run, and it drives the
policy rather than a sync: the run function is a fake and the clock is injected, the way
`SyncDeps.now` is in `test/mirror-sync.test.ts`. Two things are asserted and nothing else
touches either. Which sync reports license `source: "onenote"` — four conditions, and a
test that negates each one separately, because a single `outcome === 'complete'` check
passes three of the four ways a run returns with work left behind. And how often a refresh
may run: every assertion about a refresh that did **not** run is an assertion about
OneNote's 400 requests an hour, which is the budget this feature spends.

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
watermark overlap nearly free; and a sweep whose two stamps disagree re-fetches the page
rather than marking it stale, because a mark deletes the content document and nothing
re-fetches a stale page. What no test there covers is whether Graph's timestamps
behave as the algorithm assumes — that a page create, edit and delete each move the
section's `lastModifiedDateTime` is measured in `api-overview.md`, not checked here.

The activity tests in `test/mirror-sync.test.ts` are all about something that does **not**
happen, because a filter that silently stopped filtering costs Graph requests against a
400-per-hour budget and nothing in a run report would look wrong. The one easiest to lose is
the `sectionRollUpTrusted: false` case: `pickCandidates` returns early there, so a filter
folded into it would have to be applied on both sides of that branch, and missing the second
side means visiting every archived section every run. `test/mirror-reader.test.ts` asserts
all nine `sourceFor` combinations rather than a sample, because each one is a claim a model
acts on, and the `all` row reports `best-available` even with a failed refresh — which reads
as a bug until you notice the refresh was never going to check that notebook.

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
