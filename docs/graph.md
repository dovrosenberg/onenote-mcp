# Graph conventions

Read before adding or changing a Microsoft Graph call, and before anything that spends
Graph requests in a loop. `api-overview.md` records what the endpoints actually do; this
file records what this repository does about it.

## Calling Graph

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

## Auth and the token cache

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

**Every real Graph call carries a 60-second timeout, and the reason is the gate rather
than politeness.** Node's `fetch` is undici, whose `headersTimeout` and `bodyTimeout` both
default to 300 seconds — the same as Cloud Run's request timeout and longer than the
mirror sync's whole 240-second budget. A call that connects and never answers therefore
holds one of the gate's four slots for longer than the run that started it is allowed to
live, and four of them wedge the gate completely: every later request queues and every
caller is cut at 300 seconds having done nothing. Neither budget prevents it, because both
are checked *before* an operation starts and the operation is the thing hanging.
`withRequestTimeout` is applied by the three `create*` factories to the real `fetch` only,
so a test injecting a fake is unaffected — what it bounds is a socket, and a fake has
none. A timed-out request rejects with a `TimeoutError`, which carries no `status`, so
`retryWait` declines it: a service that has stopped answering is not helped by being asked
again inside the same run.

**The gate's concurrency cap is sound, and the microtask ordering is why.** `release()`
decrements `inFlight` and then wakes a waiter whose continuation is a microtask, which
looks like it should let a caller arriving in between take the slot and push concurrency
past the cap. It cannot: microtasks are FIFO, so the woken waiter's increment always runs
before any later caller's `acquire`, and `release()` is itself always inside a microtask
rather than synchronous with a gate-open. This was probed rather than assumed. Do not
"fix" it by incrementing in `release`.

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

**The inline refresh is a third term in the hourly arithmetic, and it is bounded by an
interval rather than by a budget.** With `MIRROR_READ_ENABLED=true` every covered read
runs `runIncremental` at `INLINE_SYNC_REQUEST_BUDGET` (12) and
`INLINE_SYNC_TIME_BUDGET_MS` (15s) — small on purpose, because this runs inside a tool call
a person is waiting on and a refresh needing more than a dozen requests is a backfill
rather than a catch-up. What actually bounds the spend is how often it may run at all: at
most once per 30 seconds when refreshes are finishing, which on a quiet account is one
request each and about 120/hour from a continuously busy conversation, and once per 5
minutes when they are not, which is at most 144/hour. Both numbers sit beside the
scheduler's own runs under the same 400, so raising either constant means redoing the
schedule arithmetic in `README.md`.

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
