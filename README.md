# onenote-mcp

An MCP server that exposes Microsoft OneNote through Microsoft Graph — notebook and
section structure, page content, and handwriting rendered to an image the calling model
can read.

**[`project-spec.md`](./project-spec.md) is the authoritative design document.** It covers
the ink reconstruction pipeline, the two independent OAuth layers, the Cloud Run
deployment model, and the Firestore-backed token cache. Read it before changing anything
here.

## Requirements

Node >= 24. `@google-cloud/firestore` requires Node >= 22, and Node 24 is the current
Active LTS.

## Quick start

```bash
npm ci
npm run build
npm test
```

## Scripts

| Script | What it does |
|---|---|
| `npm run build` | Compile `src/` to `dist/` |
| `npm run typecheck` | Type-check without emitting |
| `npm run dev` | Run the server from source with `--watch` |
| `npm start` | Run the compiled server from `dist/` (run `build` first) |
| `npm run bootstrap` | Local device-code sign-in that seeds the Firestore token cache |
| `npm test` | `node --test` over `test/**/*.test.ts` |

Tests live in `test/` and mirror `src/`. They run directly against TypeScript source
using Node's native type stripping, so `npm test` does not need a build. That constrains
the source: no `enum`, no `namespace`, no constructor parameter properties, and type-only
imports must be written `import type`. The `erasableSyntaxOnly` and `verbatimModuleSyntax`
compiler options enforce this.

See [`CLAUDE.md`](./CLAUDE.md) for the directory layout and the conventions that go with
it.

## Token cache

`src/token-cache.ts` implements MSAL's `ICachePlugin` against a single Firestore
document, whose path comes from `FIRESTORE_CACHE_DOC`. `beforeCacheAccess` reads the
document's `cache` field and hands the string to MSAL. `afterCacheAccess` writes the
serialized cache back inside a Firestore transaction, and only when MSAL reports that the
cache changed. A document that does not exist is read as an empty cache, which is the
state before `npm run bootstrap` has been run. Both entrypoints use this same plugin:
the bootstrap CLI writes the cache through it and the server reads through it, so there
is one serializer and no second format to keep in step.

That blob is the only copy of the refresh token, so two things protect it.

**A write that would empty the document is refused.** MSAL removes credentials from its
in-memory cache on some failures, and `afterCacheAccess` runs inside MSAL's `finally`
block, so a serialization that has lost the account can reach this code while the stored
one is still good. `overwriteWouldEmptyCache` stops it and logs
`{"event":"token-cache-write-refused"}`. The emptiness check reads no MSAL key name — a
cache is empty when it parses to an object whose every value is an empty container — so
it cannot invert when MSAL changes its format, and anything it does not recognise is
allowed through rather than blocked.

**The blob each write replaces is kept in a `previousCache` field.** One generation, not
a history: the cache is rewritten on every refresh, and the useful copy is always the
most recent good one. Recovering from a bad write is copying that field over `cache` in
the Firestore console, which is worth having because the alternative is a device-code
sign-in. Turn on point-in-time recovery for a second layer:

```bash
gcloud firestore databases update --enable-pitr
```

**A backend failure is not a credential failure.** Firestore being unreachable, or a
revoked `roles/datastore.user` binding, raises `TokenCacheUnavailableError` rather than
surfacing as the error a dead refresh token produces. Writes are retried three times
before that. See the `cache-unavailable` row in the table below for why the distinction
is worth the code.

`npm test` covers only `readCache`, the function that decodes a document snapshot. The
two callbacks, the transaction, and `createFirestoreTokenCachePlugin` have no automated
test — they need a Firestore backend. Exercising them means the emulator, which needs
`java` on `PATH` and an install of its own:

```bash
sudo apt-get install google-cloud-cli-firestore-emulator
```

`gcloud components install cloud-firestore-emulator` does not install it on a
Debian-packaged Google Cloud CLI. The component manager is disabled in that build, and
`gcloud` prints the `apt-get` command above in its place.

## Graph auth

`src/graph-auth.ts` turns the seeded token cache into a Microsoft Graph access token.
`createGraphAuth` builds one `PublicClientApplication` from `ONENOTE_CLIENT_ID`,
`ONENOTE_AUTHORITY`, and the Firestore cache plugin, and holds it for the life of the
process. `getAccessToken()` reads the cached account, calls `acquireTokenSilent`, and
returns the token. Requested scopes are `Notes.Read` and `Notes.ReadWrite`, fully
qualified.

The deployed server never signs in interactively. It has no way to prompt anyone, and
Graph's OneNote endpoints do not support app-only auth, so there is no fallback when the
stored refresh token dies — a human re-runs `npm run bootstrap`. Every failure is
therefore a `GraphAuthError` saying so, rather than a raw MSAL error that would reach the
caller as a bare 401 from Graph:

| `reason` | What happened | What to do |
|---|---|---|
| `cache-unreadable` | The Firestore document is absent, or its `cache` field is not something MSAL can deserialize | `npm run bootstrap` |
| `cache-unavailable` | Firestore did not answer, or the runtime service account lost `roles/datastore.user` | Retry. **Not** a sign-in. |
| `no-account` | The cache was read but holds no signed-in account | `npm run bootstrap` |
| `silent-failed` | The stored refresh token is expired or revoked, or the token endpoint returned nothing usable | `npm run bootstrap` |

`cache-unavailable` is the row that earns its keep. Firestore is read and written inside
`acquireTokenSilent`, through the cache plugin, so a backend outage used to arrive as the
same rejection a dead refresh token produces — and that message tells the operator to go
to a browser and replace a credential that is working. `GraphAuthError.retryable` carries
the distinction and only that reason sets it.

Every one of these also writes one line to stderr:

```json
{"event":"graph-auth-failure","reason":"silent-failed","documentPath":"tokencache/msal","retryable":"false"}
```

That line is the point. A tool failure otherwise appears only inside a Claude
conversation, so without it nothing tells the operator the connector has stopped working.
See **Alerting** below.

The messages name the document path and the underlying MSAL error, and deliberately
carry no account identifier: `username` is the user's UPN and `homeAccountId` embeds the
tenant id, neither of which belongs in a log.

`npm test` covers the acquisition logic through a fake client. `createGraphAuth` itself
has no automated test: it needs a cache seeded by a real device-code sign-in, and no
credential that could seed one may be committed. Run `npm run bootstrap` and then the
server against the same document to exercise it. Its consumer is the Graph structure
client below; nothing wires either into `createApp` yet.

## Graph structure

`src/graph-structure.ts` reads the OneNote tree: notebooks, section groups, sections, and
the page list inside one section. `new GraphStructure(auth)` takes anything with a
`getAccessToken()`, so the server passes it the `GraphAuth` above.

| Method | Returns |
|---|---|
| `listNotebooks()` | Every notebook, by display name |
| `listSections(containerKind, containerId)` | Sections directly under a notebook or section group |
| `listSectionGroups(containerKind, containerId)` | Section groups directly under a notebook or section group |
| `listContainerChildren(containerKind, containerId)` | Both of the above, fetched together |
| `listPagesInSection(sectionId, top?)` | Pages in one section, most recently modified first, at most `top` (default 50) |
| `getNotebookTree(notebook)` | One notebook with every nested section group resolved |
| `getFullTree()` | Every notebook, each with its tree resolved |
| `getExpandedTree()` | Every notebook with its sections and one level of section group, in a single request |
| `findSectionsByName(displayName)` | Sections anywhere in the account whose name contains that text, each with its parent notebook and section group |
| `findPagesByTitle(sectionId, title)` | Pages in one section whose title matches, compared case-insensitively by Graph |

`containerKind` is `notebooks` or `sectionGroups` — the two Graph relationship names.
Both container kinds expose the same child relationships, which is why the list methods
take the kind instead of existing twice.

`getExpandedTree()` is the cheap one. It asks Graph to expand the relationships rather
than walking them:

```
GET /me/onenote/notebooks?$select=id,displayName
    &$expand=sections($select=id,displayName),
             sectionGroups($select=id,displayName;$expand=sections($select=id,displayName))
```

Measured against a 54-notebook account: one request and 78 KB, against 195 requests for
`getFullTree()`, which matters because OneNote allows 400 requests an hour and 5
concurrent. The `$select` inside each expand clause is what takes the response from
441 KB to 78 KB, and the separator inside a clause carrying both `$select` and `$expand`
is a semicolon. What it does not reach is a section group nested inside a section group —
Graph caps `$expand` nesting at two levels — so `findSectionsByName` covers that case in
one request instead, by filtering the account-wide section list and expanding each
section's parents.

`api-overview.md` records what these endpoints accept, including the places the service
contradicts its own documentation.

Three things the traversal handles that a single Graph call does not:

- **Nesting.** Section groups are the UI's "tab groups", and they contain further section
  groups. `getNotebookTree` recurses.
- **Paging.** Every list call follows `@odata.nextLink` until it stops appearing. Graph
  chooses its own page size and ignores a larger `$top`, so one response is never proof
  that a collection is complete. `listPagesInSection` stops as soon as `top` items are in
  hand, so `top` is a result count rather than a page size.
- **The account-wide page list is never called.** `GET /me/onenote/pages` fails with
  error 20266, "maximum sections exceeded", on a notebook-per-year structure. Page
  listing is always scoped to `/me/onenote/sections/{id}/pages`, and a test scans `src/`
  for the account-wide path.

Failures are `GraphRequestError` for a non-2xx response — it carries `status`,
`statusText`, and the response `body`, because error 20266 is only distinguishable from
any other 400 by that text — and `GraphResponseError` for a 2xx whose body is not the
expected shape, a listing that will not terminate, or section groups nested past 20
levels. No message contains a notebook, section, or page name.

`npm test` drives all of it through a fake `fetch` keyed by exact URL. What that cannot
check is whether Graph accepts those URLs; the query strings come from the validated
recon script in Appendix A of `project-spec.md` and are confirmed only by running against
the real tenant.

## Ink

Graph's normal page-content endpoint drops handwriting and leaves
`<!-- InkNode is not supported -->` behind, and Graph cannot export a page as an image or
a PDF. Handwriting is therefore rebuilt from raw stroke data:
`GET /me/onenote/pages/{id}/content?includeInkML=true` answers `multipart/mixed`, one
part the same HTML and another the InkML. The strokes become an SVG and then a PNG, which
goes to the calling model as an image for its own vision to read. No OCR service is
involved.

| Module | What it does |
|---|---|
| `src/multipart.ts` | `splitMultipart(body, contentType)` → the parts, or `null` when the response is not multipart |
| `src/ink.ts` | `parseInkStrokes(text)` → strokes; `strokesToSvg`; `rasterizeSvg`; `renderInk(text, width?)` → a PNG or `null` |
| `src/page-content.ts` | `GraphPageContent.fetchRaw(pageId)` → the split response; `.fetchInk(pageId)` → the PNG or `null` |

Four details decide whether this works at all, and all four come from the validated recon
script in Appendix A of `project-spec.md`:

- **Namespaces are stripped.** Graph emits `inkml:ink`, `inkml:trace`, `inkml:traceFormat`.
  `fast-xml-parser` is configured with `removeNSPrefix: true` and every lookup uses the
  bare name.
- **Channel order comes from `<traceFormat>`.** This account's points are X, Y, F, where F
  is pen pressure. Reading the first two numbers of each point draws pressure as a
  coordinate.
- **Coordinates are himetric.** `px = himetric * 96 / 2540`. That is the same coordinate
  space the page HTML positions typed content in, so ink and typed content could later be
  registered against each other by arithmetic.
- **Traces are anywhere in the tree.** A page can carry more than one `<ink>` root, and
  `<traceGroup>` elements nest. All of them are collected.

A page with no ink renders to `null`. That is the normal answer for a typed page, not an
error. The failures that do raise are `InkParseError` for trace groups nested past 50
levels and `InkRenderError` for a document resvg rejects; neither message reproduces any
of the document, because stroke coordinates are the user's handwriting.

`test/fixtures/*.inkml` are hand-authored — a few strokes, X/Y/F channel order, himetric
units, one file with two `<ink>` roots and nested `<traceGroup>` elements. No captured
page dump may be committed: rendered ink is fully legible personal notes.

## MCP endpoint

The server speaks MCP over **stateless Streamable HTTP** at `POST /mcp`. Every request
builds its own MCP server, answers, and tears it down; nothing survives to the next one.
There is no session id, and there is no SSE — `GET /mcp` and `DELETE /mcp` are answered
405, and a POST replies with a JSON body rather than opening a stream. An open stream
would hold a Cloud Run instance alive and bill for idle time.

```bash
curl -s -X POST localhost:8080/mcp \
  -H 'content-type: application/json' \
  -H 'accept: application/json, text/event-stream' \
  -d '{"jsonrpc":"2.0","id":1,"method":"tools/list","params":{}}'
# {"result":{"tools":[]},"jsonrpc":"2.0","id":1}
```

Both `Accept` types are required by the Streamable HTTP spec even though this server
never streams. `createTools` in `src/tools.ts` is the registry — six browsing tools
(`list_notebooks`, `list_sections`, `list_pages`, `search_pages`, `find_page_by_name`,
`list_pages_by_name`), one reading tool (`get_page_content`), and five writing tools
(`append_to_page`, `append_to_page_by_name`, `create_page`, `create_page_by_name`,
`update_page_title`) — and `src/mcp-server.ts` is the
JSON-RPC surface around them.

A tool that throws comes back as a tool result with `isError: true` and a readable
message — an expired refresh token, a page that is gone, and a document resvg rejects are
all normal outcomes, not protocol faults. Only a call to a tool that was never registered
is a JSON-RPC error.

Every request writes one JSON log line: the HTTP verb, the path, the status, the
duration, the JSON-RPC method, and the tool name on a `tools/call`. Never the query
string, the headers, the arguments, or the result — see `src/logging.ts`.

`/mcp` is closed behind a bearer token — see [Bearer tokens on the MCP
endpoint](#bearer-tokens-on-the-mcp-endpoint). The health endpoint stays open.

## OAuth discovery

Claude has to find the authorization server before it can start a flow. `src/oauth-router.ts`
mounts the SDK's `mcpAuthRouter` at the application root — it builds its paths from the
issuer URL rather than from a mount point, so it cannot go behind a prefix — and serves
five routes, all of them unauthenticated by necessity:

| Path | What it is |
|---|---|
| `GET /.well-known/oauth-authorization-server` | RFC 8414 authorization-server metadata |
| `GET /.well-known/oauth-protected-resource/mcp` | RFC 9728 protected-resource metadata |
| `GET,POST /authorize` | Authorization endpoint |
| `POST /consent` | Where the consent form posts back; the SDK's `/authorize` router owns no resume path |
| `POST /token` | Token endpoint |

```bash
curl -s localhost:8080/.well-known/oauth-authorization-server
curl -s localhost:8080/.well-known/oauth-protected-resource/mcp
```

Everything in both documents is derived from `MCP_PUBLIC_URL`: it is the issuer, and the
`resource` identifier is it plus `/mcp`. `MCP_PUBLIC_URL` is rejected at startup if it
carries a trailing slash, so that every URL built by concatenating a path onto it is
well-formed; the `issuer` field then reports the URL-normalised form, which for an
origin-only value is the same string with a trailing slash added back. The protected-resource document is served only
at the path-suffixed URL — the bare `/.well-known/oauth-protected-resource` is a 404, and
so is `/.well-known/openid-configuration`. Claude probes the suffixed path first.

`scopes_supported` lists `offline_access`, which is what makes Claude ask for a refresh
token rather than re-consenting whenever an access token expires. There is no
`registration_endpoint`: the client id and secret are configured, so Dynamic Client
Registration has nothing to do. One client is registered, with three redirect URIs —
`https://claude.ai/api/mcp/auth_callback` for the hosted Claude surfaces, and
`http://localhost/callback` plus `http://127.0.0.1/callback` for Claude Code, whose port
is ignored per RFC 8252.

`GET /authorize` renders a consent page rather than redirecting: one Approve button
naming what is granted and the host the authorization code will be sent to. Approving
posts back to `POST /consent`, which mints a 60-second single-use code and redirects to
the client's callback. The whole authorization request crosses that page in one hidden
field signed with `MCP_TOKEN_SIGNING_KEY`, so a mid-consent instance replacement does not
break the flow and the form cannot be edited; a field that fails to verify is a 400 with
no redirect and no code minted.

Both consent responses carry `Cache-Control: no-store`, `Referrer-Policy: no-referrer`
— the form posts from the `/authorize` URL, which has `state` and the PKCE challenge in
its query string — `X-Frame-Options: DENY`, and a CSP of `default-src 'none'; style-src
'unsafe-inline'; frame-ancestors 'none'; base-uri 'none'`. There is deliberately no
`form-action`: browsers have disagreed about whether it is checked against a redirect
target, and the consent POST answers with a redirect to claude.ai.

`POST /consent` has a rate limit of its own — 200 in 15 minutes — because it is mounted
ahead of the SDK's `/authorize` limiter on purpose and a rendered form stays postable for
ten minutes, so one trip through `/authorize` yields a field that can be replayed. The
limit sits above the 100-entry pending-code cap so that the store's own eviction, whose
behaviour is specified, is what a burst runs into first.

`POST /token` issues an access token good for one hour and a refresh token good for 30
days. Both are an HMAC-SHA256 over a compact payload under `MCP_TOKEN_SIGNING_KEY` and
nothing else — no store is consulted to verify one, which is what keeps a Cloud Run
revision replacement from forcing a reconnect. The payload carries the audience, which is
`MCP_PUBLIC_URL` plus `/mcp`, so a token is good for this MCP endpoint and no other. How
long those tokens live, and how to make a human approve more often, is two sections down.

## Bearer tokens on the MCP endpoint

Every request to `/mcp` needs `Authorization: Bearer <access token>`. The SDK's
`requireBearerAuth` sits in front of the MCP router in `createApp`, and
`verifyAccessToken` in `src/oauth-provider.ts` is what it calls: the HMAC signature under
`MCP_TOKEN_SIGNING_KEY`, the token kind, the expiry, and the audience. A token that is
correctly signed and unexpired but carries another server's resource identifier is
refused — the SDK checks no audience of its own, so without that check a token minted for
a different MCP server by a server sharing this signing key would be accepted.

A request with no token, an expired token, or a token that fails any of those checks is
`401` with a challenge header:

```
WWW-Authenticate: Bearer error="invalid_token", error_description="…",
                  resource_metadata="https://<MCP_PUBLIC_URL>/.well-known/oauth-protected-resource/mcp"
```

The `resource_metadata` parameter is the part that matters: it is how Claude finds the
authorization server and starts the flow, so a 401 without it is a dead end rather than a
sign-in prompt. Claude refreshes reactively on a 401 and proactively a few minutes before
the stored expiry, so a 401 here is an ordinary event.

The token is read from the `Authorization` header and from nowhere else. `?access_token=`
in the query string is not honoured — the MCP authorization spec forbids it, and
`src/logging.ts` leaves the query string out of the log line on the strength of that.

Which routes are open is the exempt list, and it is longer than "everything except
`/mcp`" because the whole authorization flow has to answer callers who hold no token yet:
`/healthz` and `/health`, both `.well-known` documents, `/authorize`, `/consent`, and
`/token`. A test
in `test/server.test.ts` enumerates the routes `createApp` actually registers and asserts
that every one not on that list answers 401 without a token, so a route added later is
closed unless someone opens it deliberately.

No scopes are required. `offline_access`, the one scope this server issues, is about
whether a refresh token is granted rather than about what a caller may do, and requiring
it would answer 403 for tokens that are otherwise good. If a scope check is ever added,
the 403 has to carry `WWW-Authenticate: Bearer error="insufficient_scope"` — which this
middleware does — because Claude treats any other 403 as terminal and prompts for
nothing.

## Token lifetime and forcing revalidation

This server is built to run unattended. The default settings reflect that, and they trade
away some ability to cut off a leaked credential. Read this before deploying it somewhere
that matters, and change the numbers if the trade is wrong for you.

### What the defaults do

| Token | Lifetime | What renews it |
|---|---|---|
| Access token | 1 hour | The refresh token, automatically |
| Refresh token | 30 days | Every refresh mints a new one with a fresh 30 days |
| Consent form | 10 minutes | Nothing; a stale form is refused and the flow restarts |

Claude refreshes on its own — proactively before the hour is up, and reactively on a 401.
So a human clicks Approve when the connector is first added, and then only if the
connector goes unused for 30 days. That is the sliding window: the 30 days bound how long
the connection may sit **idle**, not how long it may live.

### Why sliding, and what it costs

Every token this server issues is stateless. It is a signed payload and nothing more — no
database row, no session record, nothing to look up when it comes back. That is what makes
a Cloud Run revision replacement invisible: the new instance verifies a token the old
instance issued, with no shared state between them. A token store would mean a reconnect on
every deploy.

The cost is that **nothing can be revoked individually**. There is no revocation endpoint
because there is nothing for it to delete. Specifically:

- A refresh token that leaks grants access for up to 30 days, and each use extends its
  holder's access by another 30. There is no server-side record to invalidate, and no way
  to tell a stolen refresh token from the legitimate one — both are the same bytes signed
  by the same key.
- Sliding the window is **not** rotation. When a refresh mints a new refresh token, the
  one it replaces keeps working until the expiry stamped inside it. Real rotation means
  marking the old token spent, which needs the store this design does not have.
- An access token cannot be cut off inside its hour, for the same reason.

What is left is one blunt lever, and it works immediately: **change
`MCP_TOKEN_SIGNING_KEY` and redeploy.** Every access token, every refresh token and every
open consent page is invalidated at once, because all of them are verified against that
key. The next Claude request gets a 401 and the operator clicks Approve once. Rotating the
key on a schedule is a reasonable policy on its own.

The consent screen, for what it is worth, authenticates nobody — it has one button and no
password. What stands between a stranger and your notebooks is `MCP_OAUTH_CLIENT_SECRET`,
which `POST /token` requires, the redirect-URI allowlist that sends every authorization
code to `claude.ai` or to loopback, and PKCE binding the code to the client that started
the flow.

### Making a human approve more often

Each of these is a source change, not a configuration value. That is deliberate: an
operator who shortens the window is changing the security posture of the deployment, and
that belongs in a commit somebody can read rather than in an environment variable somebody
can forget.

**Shorten the idle window.** In `src/oauth-provider.ts`:

```ts
const REFRESH_TOKEN_TTL_S = 30 * 24 * 60 * 60;   // 30 days
const REFRESH_TOKEN_TTL_S = 7 * 24 * 60 * 60;    // a week
```

Unused for that long, the connector needs a click. Used regularly, it still never asks —
the window keeps sliding forward. This bounds how long a leaked refresh token survives
after the leak stops being used, and nothing more.

**Stop the window sliding.** This is what issue #22 originally specified, and it caps the
total life of a connection rather than its idle time: a human approves every 30 days no
matter how busy the connector is. One line in `exchangeRefreshToken`, in
`src/oauth-provider.ts`:

```ts
// Sliding: a new refresh token, 30 days from now.
return issueTokens(client.client_id, requested, mintRefreshToken(client.client_id, granted));

// Fixed: hand back the same token, expiring 30 days after the consent click.
return issueTokens(client.client_id, requested, refreshToken);
```

**Refuse to issue refresh tokens at all.** The strictest setting: a human approves every
hour, because an expired access token has nothing to renew it. Two edits, both needed —
the metadata switch alone does not stop the token being issued.

1. In `src/oauth-router.ts`, empty `SCOPES_SUPPORTED`. Claude appends `offline_access` to
   an authorization request only when the metadata advertises it, and that is the switch
   deciding whether it asks for a refresh token.
2. In `src/oauth-provider.ts`, drop the `refresh_token` field from what `issueTokens`
   returns. It is issued today regardless of the scopes requested.

Expect this one to be visible in use: Claude sends the browser back to the consent screen
mid-session when the hour runs out.

**Shorten the access token.** `ACCESS_TOKEN_TTL_S` in `src/oauth-provider.ts` narrows the
window in which a leaked *access* token works. It costs a token request per expiry and no
human involvement at all, so it is cheap — but it does nothing about a leaked refresh
token, which is the credential worth worrying about.

## Keepalive

Microsoft's delegated refresh tokens lapse after roughly 90 days without use. The token
only slides forward when it is actually exchanged, and it is only exchanged when a tool
call arrives after the held access token has expired — so a connector nobody uses for
three months is a connector that needs a person at a browser running `npm run bootstrap`.
Nothing in the server can prevent that on its own, because nothing in the server runs
when nobody is calling it.

`POST /keepalive` is the fix. It calls `acquireTokenSilent` with `forceRefresh: true`,
which skips the held access token and exchanges the refresh token, so Entra issues a
replacement with a fresh window and `src/token-cache.ts` writes it to Firestore.
`forceRefresh` is the load-bearing part: without it MSAL answers from its own cache, no
request reaches Entra, and the window does not move.

Set `MCP_KEEPALIVE_SECRET` to at least 32 random characters and the route is mounted;
leave it unset and the path 404s. A scheduler presents the secret in the
`X-Keepalive-Secret` header, which is compared in constant time before any work is done.
It is a shared secret rather than a bearer token because a scheduler cannot run the OAuth
flow — it has no browser and nowhere to keep a refresh token — and it is its own variable
rather than the Layer-1 client secret so that a credential which can reach the whole MCP
surface is not also sitting in a scheduler job.

### Setting it up on your own account

Four steps: generate a secret, get it onto the service, create the job, and check the
job actually reaches the service. The whole thing is optional — skip it and the server
works, up until the day nobody has called it in 90 days.

```bash
# 1. Generate the secret. umask 077 so the file is not world-readable, and it is never
#    printed: the two commands below read it from disk.
umask 077
openssl rand -hex 32 > keepalive.secret

# 2. Get it onto the service. The GitHub secret alone changes nothing — the value
#    reaches the container as an env var, which happens on a deploy and only then.
gh secret set MCP_KEEPALIVE_SECRET < keepalive.secret
gh workflow run deploy.yml --ref main

# 3. Create the job. The API enable takes a minute or two to propagate; a create run
#    immediately after it fails with SERVICE_DISABLED, which is a retry rather than a
#    misconfiguration.
gcloud services enable cloudscheduler.googleapis.com --project="$GCP_PROJECT"

gcloud scheduler jobs create http onenote-mcp-keepalive \
  --project="$GCP_PROJECT" \
  --location="$GCP_REGION" \
  --schedule="0 4 * * 1" \
  --time-zone=UTC \
  --uri="$MCP_PUBLIC_URL/keepalive" \
  --http-method=POST \
  --headers="X-Keepalive-Secret=$(cat keepalive.secret)" \
  --attempt-deadline=60s \
  --max-retry-attempts=3

rm keepalive.secret
```

`--location` is required and is the scheduler's own region, which has nothing to do with
where the job's target is; using the Cloud Run region keeps one fewer value in your head.
`scripts/gcp-bootstrap.sh` enables `cloudscheduler.googleapis.com` too, so step 3's
enable is only needed on a project bootstrapped before that line existed.

Then prove it, because every failure mode here is silent — a wrong secret, a job pointed
at the old URL, and a service deployed without the variable all look like a scheduler job
sitting there enabled:

```bash
gcloud scheduler jobs run onenote-mcp-keepalive --project="$GCP_PROJECT" --location="$GCP_REGION"

# An empty status code is success. A code is a gRPC status; the request log below says why.
gcloud scheduler jobs describe onenote-mcp-keepalive \
  --project="$GCP_PROJECT" --location="$GCP_REGION" \
  --format='value(status.code,lastAttemptTime,scheduleTime)'

# The request as the service saw it. 200 is the answer; 401 is a secret mismatch and 404
# means the deploy in step 2 did not happen.
gcloud logging read \
  'resource.type="cloud_run_revision" AND httpRequest.requestUrl:"keepalive"' \
  --project="$GCP_PROJECT" --limit=3 --freshness=10m \
  --format='value(timestamp,httpRequest.requestMethod,httpRequest.status,httpRequest.userAgent)'
```

A 200 means the token was exchanged. The thing that proves it was *stored* is the
`updatedAt` field of the Firestore document moving to the time of the run, which is the
only evidence the replacement refresh token survived the call. Measured on this
deployment on 2026-08-19: the forced run answered 200 and `tokencache/msal` advanced from
`14:15:53Z` to `14:30:54Z`.

Weekly is ample against a 90-day window and leaves room for several missed runs. The job
costs one token-endpoint round trip and one Firestore write.

To rotate the secret: set the new GitHub secret, deploy, then
`gcloud scheduler jobs update http onenote-mcp-keepalive --update-headers=…`. Any run
between the deploy and the job update answers 401 and does no work, which on a weekly
schedule is a window nothing lands in.

The header sits in the job definition, readable by anyone with `roles/cloudscheduler.viewer`
on the project, exactly as the env vars sit readable in the Cloud Run revision spec. That
is the same tradeoff taken in **Deploy** for not running Secret Manager.

| Status | Meaning | What the scheduler should do |
|---|---|---|
| 200 | The refresh token was exchanged and the new one stored | Nothing |
| 401 | The secret is absent or wrong | Fix the job; the route did no work |
| 404 | `MCP_KEEPALIVE_SECRET` is not set on the service | Set it and redeploy |
| 503 with `"retryable": true` | Firestore was unreachable | Retry |
| 503 with `"retryable": false` | The grant is dead | `npm run bootstrap` |

What this does **not** protect against: a conditional-access sign-in-frequency policy, a
password change, an MFA reset, or an admin revoking the grant. Any of those kills the
refresh token whatever the schedule says, and no code change avoids it. If the Entra
tenant is yours, exempt this app registration from sign-in-frequency policies; if it is
not, treat 90 days as an upper bound somebody else can shorten without telling you.

The keepalive route is also unrelated to the Layer-1 30-day window in **Token lifetime**
below. That refresh token lives in Claude's connector store, and only Claude can present
it or receive its replacement, so nothing running here can keep it alive. Losing it costs
one click on the Approve button; losing the Microsoft one costs a device-code sign-in.


## Page mirror sync

The mirror keeps Firestore copies of the pages in a hand-picked set of notebooks, so the
read tools can answer without spending OneNote's request budget. `POST /sync` is how it
gets filled. Nothing about it is automatic: the route is not mounted without a secret, and
the read tools do not consult the mirror without `MIRROR_READ_ENABLED`.

### What it costs

OneNote allows a delegated app 400 requests an hour. The sync is sized to stay well under
that and to stop rather than overrun.

| Run | Requests |
|---|---|
| Incremental, nothing changed | 1 — the expanded tree, and nothing else |
| Incremental, a busy day (2 sections, 8 pages) | about 11 |
| Incremental, degraded (tree read failed, or the section roll-up distrusted) | 1 + one per mirrored section + one per changed page |
| Nightly sweep, scoped to sections whose timestamp moved | about 18 on 40 sections |
| Weekly full sweep | about 53 on 40 sections |
| First backfill | one per page, spread across runs |

`MIRROR_SYNC_REQUEST_BUDGET` (default 120) is a hard stop. A run that reaches it commits
what it did, answers 200 with `"done": false`, and resumes on the next schedule. That is
what makes the first backfill possible at all.

**The schedule and the budget multiply, and the product has to stay under 400.** This is
the one arithmetic mistake here that gets the account throttled, and a throttle outlasts a
short backoff — see the `Graph request budget` section of `CLAUDE.md`. Every run during a
backfill spends its whole budget, so:

| Schedule | Budget | Requests/hour | |
|---|---|---|---|
| `*/15` | 120 | 480 | **over the limit** |
| `*/20` | 120 | 360 | leaves 40/hour for tools |
| `*/20` | 100 | 300 | leaves 100/hour for tools — use this while backfilling |
| `*/15` | 120 | ~4–45 | fine *after* the backfill, when runs are near-empty |

Backfill on this account — 180 mirrored sections — is roughly one request per section plus
one per page, so a notebook set holding a few thousand pages is ten hours or more of
slices. Start on `*/20`, and move to `*/15` once `sync/state.backfillComplete` is true and
runs are answering `outcome: complete` in a couple of requests.

### Setting it up

```bash
# 1. The secret. Never printed; both commands below read it from disk.
umask 077
openssl rand -hex 32 > sync.secret

gh secret set MIRROR_SYNC_SECRET < sync.secret
gh workflow run deploy.yml --ref main

# 2. The jobs. --attempt-deadline sits just above Cloud Run's 300s request timeout, so
#    the platform's cut is the one that happens and the scheduler records it.
gcloud scheduler jobs create http onenote-mcp-sync \
  --project="$GCP_PROJECT" --location="$GCP_REGION" \
  --schedule="*/20 * * * *" --time-zone=UTC \
  --uri="$MCP_PUBLIC_URL/sync" --http-method=POST \
  --headers="X-Sync-Secret=$(cat sync.secret)" \
  --attempt-deadline=330s --max-retry-attempts=2

gcloud scheduler jobs create http onenote-mcp-sync-sweep \
  --project="$GCP_PROJECT" --location="$GCP_REGION" \
  --schedule="25 8 * * *" --time-zone=UTC \
  --uri="$MCP_PUBLIC_URL/sync/sweep" --http-method=POST \
  --headers="X-Sync-Secret=$(cat sync.secret)" \
  --attempt-deadline=330s --max-retry-attempts=1

gcloud scheduler jobs create http onenote-mcp-sync-sweep-full \
  --project="$GCP_PROJECT" --location="$GCP_REGION" \
  --schedule="40 9 * * 0" --time-zone=UTC \
  --uri="$MCP_PUBLIC_URL/sync/sweep/full" --http-method=POST \
  --headers="X-Sync-Secret=$(cat sync.secret)" \
  --attempt-deadline=330s --max-retry-attempts=1

rm sync.secret
```

Retries are safe on all three: every mode is idempotent, and a retry after a
budget-exhausted run is actively useful during the backfill. `25 8` and `40 9` rather than
the hour mark keep the sweeps out of an incremental run's minute — and the run lease makes
an overlap a 409 rather than a double spend of the Graph budget.

### Choosing what to mirror

The selection is one hand-edited Firestore document, `MIRROR_ROOT_DOC`, default
`onenoteMirror/default`. It is the only document the service never writes.

| Field | Type | Meaning |
|---|---|---|
| `notebookIds` | array of string | Notebook ids whose **pages** are mirrored |
| `note` | string | Free text for you. Never read. |

Get the ids from `list_notebooks`. Structure — every notebook and section name in the
account — is mirrored regardless, because the tree read returns it all for the same one
request and a partial structure would make `list_notebooks` and an unscoped `search_pages`
answer confidently and partially. Only page *content* follows the selection.

An id matching no notebook is reported as `unknownNotebookIds` in every run report. Check
it after editing: the ids are opaque strings and a typo is otherwise a notebook that
silently never syncs.

### Proving it works

```bash
gcloud scheduler jobs run onenote-mcp-sync --project="$GCP_PROJECT" --location="$GCP_REGION"

# The report the run returned, in the service log.
gcloud logging read \
  'resource.type="cloud_run_revision" AND jsonPayload.event="sync-completed"' \
  --project="$GCP_PROJECT" --limit=3 --freshness=30m --format='value(jsonPayload)'
```

A 200 is not proof on its own. The two things that are:

- **`sections/{id}.pagesSyncedThrough` moved** in Firestore. A run that answered 200 and
  advanced no watermark did nothing.
- **The bucket holds an object per inked page**, and its byte count matches the page
  document's `ink.bytes`.

Watch the backfill by re-running the job and reading `done` in successive reports: `false`
with a rising `pagesUpdated`, then `true`. `sync/state.backfillComplete` says the same
thing.

Watch for a 429 in the request log during a run. One means something is bypassing the
shared request gate, which is the failure that gets the whole account throttled.

| Status | Meaning | What to do |
|---|---|---|
| 200, `outcome: complete` | The run finished its work | Nothing |
| 200, `outcome: budget-exhausted` | It stopped on the request or time budget | Nothing; the next run resumes |
| 401 | The secret is absent or wrong | Fix the job; the route did no work |
| 404 | `MIRROR_SYNC_SECRET` is not set on the service | Set it and redeploy |
| 409 | Another run holds the lease | Nothing; it expires after 15 minutes |
| 503, `reason: silent-failed` | The Microsoft refresh token is gone | Re-run `npm run bootstrap` |
| 503, `reason: cache-unavailable` | Firestore did not answer | Retry; the credential is fine |

### Reading from it

`MIRROR_READ_ENABLED` is the switch. With it `false` — the default — every tool answers
from Microsoft Graph exactly as it did before the mirror existed. With it `true`, seven
tools try the local copy first and fall back to Graph on anything they do not hold:
`list_notebooks`, `list_sections`, `list_pages`, `list_pages_by_name`,
`find_page_by_name`, `get_page_content` and `search_pages`.

Every one of them reports `source` as `mirror` or `graph`, and a mirrored answer carries
`mirroredAt`. Every one also takes `useLiveData: true`, which skips the local copy for
that call — the argument to pass when you need an edit made in the last few minutes.

Turning it off is a complete rollback: set the variable to `false`, redeploy, and the
tools are byte-identical to their pre-mirror behaviour. No data migration, and the mirror
keeps filling in the background.

**Writes resync their page immediately.** All five writing tools re-read the page from
OneNote after a successful write and store it, so a `get_page_content` straight after an
`append_to_page` answers from the local copy with the text that was just added. That
costs one extra OneNote request per write. If the resync fails the page is marked stale
instead, which sends the next read to OneNote — slower, but never wrong — and the next
scheduled sync repairs it either way. A failed resync never fails the write.

Two things to check before turning it on:

- **All four composite indexes read `READY`.** A query against one still `CREATING` fails
  with `FAILED_PRECONDITION`.
- **`sync/state.backfillComplete` is true.** Turning it on earlier is safe — an unmirrored
  page is a miss and goes to Graph — but most reads would fall through, which is slower
  than not having the mirror at all and spends budget the backfill needs.

The one behaviour change worth knowing about is `search_pages`. Answered from Graph it
walks sections and stops at 60 of them or 25 seconds, so it reports `sectionsSearched`,
`sectionsFound` and `stoppedEarly`. Answered from the mirror there is no walk and no
bound — but page content is held only for the notebooks in your selection, so it reports
`notebooksSearched` against `notebooksInAccount` instead, and says so in its note when
they differ. A model reading "no matches" needs to know which of those it got.

### Recovering from a bad mirror

Delete the `notebooks`, `sectionGroups`, `sections`, `pages`, `pageContent` and
`tombstones` subcollections of `MIRROR_ROOT_DOC`, and clear `structureHash` and
`sectionsScannedThrough` in `sync/state`. The next run backfills from nothing. The
selection document and the token cache are untouched by this, and no Graph write is
involved — the mirror is a copy, and Microsoft Graph stays the source of truth.


## Alerting

Two failures are invisible without a log-based metric, because both show up only as a
message inside a Claude conversation or as a line nobody is reading:

| Event | Means |
|---|---|
| `graph-auth-failure` with `retryable: "false"` | The Microsoft grant is dead. Someone has to run `npm run bootstrap`. |
| `token-cache-write-refused` | MSAL handed over a cache with no credentials in it. The stored copy survived; something is wrong. |

```bash
gcloud logging metrics create onenote_mcp_auth_failure \
  --description="Microsoft Graph credential failures needing an operator" \
  --log-filter='resource.type="cloud_run_revision"
    resource.labels.service_name="onenote-mcp"
    jsonPayload.event=("graph-auth-failure" OR "token-cache-write-refused")
    jsonPayload.retryable!="true"'
```

Then an alert policy on that metric being above zero. A consent approval is worth
watching too: `POST /consent` answering 302 should happen only when you add the
connector, and the request log already carries it.

```
jsonPayload.event="request" jsonPayload.path="/consent" jsonPayload.status=302
```

## Bootstrap

`npm run bootstrap` is the only interactive Microsoft sign-in in the project, and it runs
on your machine, not on Cloud Run. It signs in with the device-code flow and writes the
resulting MSAL cache to the Firestore document the server reads.

```bash
gcloud auth application-default login

ONENOTE_CLIENT_ID=00000000-0000-0000-0000-000000000000 \
ONENOTE_AUTHORITY=https://login.microsoftonline.com/common \
GOOGLE_CLOUD_PROJECT=your-project \
FIRESTORE_CACHE_DOC=tokencache/msal \
npm run bootstrap
```

It prints Microsoft's device-code message, waits for you to approve in a browser, then
lists your notebooks once and prints the count as proof the token works. The closing
lines name the Firestore project and document written and the account's home tenant, so
you can see you signed into the right directory. That output carries the tenant id; keep
it off issues, pull requests, and workflow logs.

`GOOGLE_CLOUD_PROJECT` and `FIRESTORE_CACHE_DOC` are **required here**, unlike on the
server where the first is inferred and the second defaults. The CLI writes with your own
Application Default Credentials, so an unset value would seed a real document in
whichever project your `gcloud` login points at, and still print a success line. The
`MCP_OAUTH_*` values are not read, so running this never puts the Layer-1 client secret
on your machine.

Run it again whenever a `graph-auth-failure` event with `retryable: "false"` appears in
the server's logs. The refresh token is rotated on every use and dies if the service sits
idle past roughly 90 days; there is no automatic recovery. Configuring the keepalive job
above is what stops idleness being one of the ways to get there.

## Container

The service deploys to Cloud Run, which runs `linux/amd64`. The image is built for that
platform explicitly so the `@resvg/resvg-js` native binary matches. The runtime base is
Debian `node:24-slim` and must not become Alpine — the resvg prebuild is glibc-only.

```bash
docker build --platform linux/amd64 -t onenote-mcp .

docker run --rm -p 8080:8080 \
  -e PORT=8080 \
  -e ONENOTE_CLIENT_ID=00000000-0000-0000-0000-000000000000 \
  -e ONENOTE_AUTHORITY=https://login.microsoftonline.com/common \
  -e MCP_OAUTH_CLIENT_ID=test-client \
  -e MCP_OAUTH_CLIENT_SECRET=test-secret \
  -e MCP_TOKEN_SIGNING_KEY=0123456789abcdef0123456789abcdef \
  -e MCP_PUBLIC_URL=https://onenote-mcp.example.run.app \
  onenote-mcp

curl -i localhost:8080/health     # 200, {"status":"ok",...}
```

`/healthz` answers the same thing and is what Cloud Run's own probes use. Do not call it
from outside: Google's frontend answers `https://<service>.run.app/healthz` with its own
404 page and the request never reaches the container, so an external uptime check has to
use `/health`. Measured against the deployed service on 2026-08-19 — `/health`,
`/healthz2` and even `/Healthz` all arrive, and only the exact lowercase `/healthz` is
swallowed.

Those values are placeholders that are only well-formed enough to pass startup
validation; they authenticate against nothing. On Cloud Run, `PORT` is supplied by the
platform and the rest come from the deploy workflow. If host port 8080 is already taken,
map a different one: `-p 8081:8080` with `-e PORT=8080`.

To test the image:

```bash
RUN_DOCKER_TESTS=1 bash scripts/test/run.sh
```

That builds the image, checks that the resvg glibc binary survived the production-only
install and that no dev dependencies came with it, renders an SVG to PNG inside the
container, and asserts `/healthz` answers 200 on the port given in `PORT`. Without
`RUN_DOCKER_TESTS=1` the docker suite is skipped and the rest still runs.

## Deploy

`.github/workflows/deploy.yml` runs on every push to `main` and on `workflow_dispatch`.
It type-checks, runs `npm test`, builds, builds and pushes the container image to
Artifact Registry tagged with the commit sha, and deploys that image to Cloud Run. A
failing type-check or test stops the run before an image is built.

There is no long-lived credential in GitHub. The job authenticates through Workload
Identity Federation: `permissions: id-token: write` lets it request a GitHub OIDC token,
and `google-github-actions/auth@v2` exchanges that for short-lived Google credentials.
No service-account JSON key is created by `scripts/gcp-bootstrap.sh` or needed anywhere.
The provider only accepts tokens whose `repository` claim is this repository.

The image is built on `ubuntu-latest`, which is `linux/amd64` — the platform Cloud Run
runs, and the one the `@resvg/resvg-js` prebuild is compiled for. That is why the
workflow builds the image itself rather than using `gcloud run deploy --source`, which
would also mean enabling Cloud Build and granting the roles that go with it.

The deploy runs at `--max-instances=1` with `--allow-unauthenticated`, as the runtime
service account, which holds `roles/datastore.user` for the Firestore token cache.
`--allow-unauthenticated` is what lets Claude reach the service at all; the MCP endpoint
is closed by the bearer token instead. See **Bearer tokens on the MCP endpoint**.

### What the repository has to hold

`scripts/gcp-bootstrap.sh` provisions the GCP side and prints the `gh variable set`
commands for the first six. The workflow fails on its first step, naming what is
missing, rather than deploying half-configured.

| Name | Kind | Value |
|---|---|---|
| `GCP_PROJECT` | variable | Project ID |
| `GCP_REGION` | variable | Cloud Run region |
| `GAR_REGION` | variable | Artifact Registry region |
| `WIF_PROVIDER` | variable | Full workload identity provider resource name |
| `DEPLOY_SA` | variable | Deploy service account email |
| `RUNTIME_SA` | variable | Runtime service account email |
| `ONENOTE_CLIENT_ID` | variable | Azure app registration client ID |
| `ONENOTE_AUTHORITY` | variable | Entra authority URL |
| `MCP_OAUTH_CLIENT_ID` | variable | Layer-1 OAuth client ID |
| `MCP_PUBLIC_URL` | variable | The service's public URL; see below |
| `FIRESTORE_CACHE_DOC` | variable | Optional; defaults to `tokencache/msal` |
| `MCP_OAUTH_CLIENT_SECRET` | **secret** | Layer-1 OAuth client secret |
| `MCP_TOKEN_SIGNING_KEY` | **secret** | Access-token signing key, at least 32 characters |
| `MCP_KEEPALIVE_SECRET` | **secret** | Optional; unset means `POST /keepalive` is not mounted |
| `MIRROR_BUCKET` | variable | Optional; the page mirror's Cloud Storage bucket. `scripts/gcp-bootstrap.sh` prints the name it created |
| `MIRROR_READ_ENABLED` | variable | Optional; defaults to `false`. The switch that makes read tools use the mirror |
| `MIRROR_ROOT_DOC` | variable | Optional; defaults to `onenoteMirror/default` |
| `MIRROR_SYNC_REQUEST_BUDGET` | variable | Optional; defaults to `120` |
| `MIRROR_SYNC_SECRET` | **secret** | Optional; unset means `POST /sync` is not mounted |

Only four of these are credentials. The WIF provider name and the service account
emails are identifiers, useless to anyone who cannot present this repository's OIDC
identity, so they are variables rather than secrets.

The deploy passes `env_vars_update_strategy: overwrite`, so the list in the workflow is
the service's entire environment on every revision. The action's default is `merge`,
under which a variable removed from the workflow would silently survive from the previous
revision. `PORT` and `GOOGLE_CLOUD_PROJECT` are deliberately not in the list: Cloud Run
supplies both, and it rejects `PORT` as an input.

### The first deploy, and `MCP_PUBLIC_URL`

`MCP_PUBLIC_URL` is the OAuth issuer and the audience of every access token this server
issues, and no service exists to have a URL until the first deploy has happened. The
workflow resolves it in three steps: the `MCP_PUBLIC_URL` repository variable, then the
URL Cloud Run has already assigned to the service, then — only when neither exists —
`https://placeholder.invalid`, which it replaces with the real URL immediately after the
deploy. So a first run works unattended and finishes with the correct value in place. It
leaves a warning naming the URL; set the repository variable to it, because that is the
only one of the three sources that survives putting a custom domain in front of the
service.

Changing `MCP_PUBLIC_URL` invalidates nothing by itself, but every access token already
issued is bound to the old audience and will be refused. Claude re-runs the authorization
flow when that happens.

### Rolling back

The image tag is the commit sha, so an earlier image is still in Artifact Registry:

```bash
gcloud run services update-traffic onenote-mcp --region "$GCP_REGION" --to-revisions <revision>=100
```

Re-running the workflow from an earlier commit with `workflow_dispatch` also works, and
is the one that keeps the deployed environment in step with that commit's workflow file.

## Configuration

Every value comes from an environment variable, validated at startup. A missing or
malformed variable produces a `ConfigError` listing everything that is wrong at once,
and the process exits 1 without a stack trace.

| Variable | Required | Default | Purpose |
|---|---|---|---|
| `ONENOTE_CLIENT_ID` | yes | — | Azure app registration client ID (public client) |
| `ONENOTE_AUTHORITY` | yes | — | Entra ID authority URL for the tenant |
| `MCP_OAUTH_CLIENT_ID` | yes | — | Layer-1 OAuth client ID that Claude presents |
| `MCP_OAUTH_CLIENT_SECRET` | yes | — | Layer-1 OAuth client secret |
| `MCP_TOKEN_SIGNING_KEY` | yes | — | Key used to sign issued access tokens (min 32 chars) |
| `MCP_PUBLIC_URL` | yes | — | The service's own public URL: `https`, no query, no fragment, no trailing slash |
| `FIRESTORE_CACHE_DOC` | server: no · bootstrap: **yes** | `tokencache/msal` | Firestore document path holding the MSAL token cache |
| `GOOGLE_CLOUD_PROJECT` | server: no · bootstrap: **yes** | — | GCP project; inferred automatically on Cloud Run |
| `PORT` | no | `8080` | Bind port. Cloud Run sets this; the server never hardcodes one. |
| `MCP_KEEPALIVE_SECRET` | no | — | At least 32 characters. Set it and `POST /keepalive` is mounted; leave it unset and the path 404s. See **Keepalive**. |
| `MIRROR_ROOT_DOC` | no | `onenoteMirror/default` | Firestore document holding the hand-edited list of notebook ids to mirror. Its subcollections are the mirror. Document path, even segment count, same rule as `FIRESTORE_CACHE_DOC`. |
| `MIRROR_SYNC_SECRET` | no | — | At least 32 characters. Set it and `POST /sync` is mounted; leave it unset and the path 404s. |
| `MIRROR_BUCKET` | conditionally | — | Cloud Storage bucket for rendered ink and oversized page HTML. Required once `MIRROR_SYNC_SECRET` or `MIRROR_READ_ENABLED` is set; see below. |
| `MIRROR_READ_ENABLED` | no | `false` | `true` or `false`, nothing else. When true the read tools answer from the mirror and fall back to Graph on a miss. |
| `MIRROR_SYNC_REQUEST_BUDGET` | no | `120` | Graph requests one sync run may spend before it stops and reports more outstanding. 10–350, against an hourly limit of 400. |

The five `MIRROR_*` variables configure the Firestore page mirror. All are optional, so a
service deployed with none of them set behaves exactly as it did before the mirror
existed — which is what makes `MIRROR_READ_ENABLED` a complete rollback switch rather
than a code change.

`MIRROR_BUCKET` is the one cross-field rule in `loadConfig`. A `VarSpec` says
required-or-not per variable and cannot say "required when another is present", but a
sync has nowhere to put a rendered ink PNG without a bucket and a mirror read has nowhere
to fetch one from. So setting `MIRROR_SYNC_SECRET`, or setting `MIRROR_READ_ENABLED` to
`true`, without a bucket fails at container startup with `MIRROR_BUCKET` in the missing
list — rather than hours into a backfill at the first object write.

`MIRROR_READ_ENABLED` accepts only `true` and `false`, case-insensitively. It is
deliberately not lenient: `1`, `yes` and `on` all read as true to a human, and a parser
that accepted them would also have to decide what `off` and `0` mean, at which point a
typo like `ture` becomes `false` and switches the mirror off with nothing to say so.

`FIRESTORE_CACHE_DOC` names the document the MSAL cache plugin in `src/token-cache.ts`
reads and writes. Its value must be a document path, meaning an even number of
slash-separated segments; `loadConfig` rejects a collection path at startup.

`ONENOTE_CLIENT_ID` and `ONENOTE_AUTHORITY` identify the Azure app registration that
`src/graph-auth.ts` presents to Entra ID. It is a public client, so there is deliberately
no Layer-2 client secret; the `MCP_OAUTH_*` values below it belong to Layer 1, between
Claude and this server, and are unrelated.

`MCP_PUBLIC_URL` is the URL Claude reaches this service at. The OAuth issuer, the
`resource` identifier a token is bound to, and the URL of the protected-resource metadata
document are all built from it. Nothing on Cloud Run tells the process what URL it is
reached at, and a value taken from the `Host` header would be whatever the caller sent,
so it is configured. It can only be filled in after the first deploy has produced the
URL.

`npm run bootstrap` reads only `ONENOTE_CLIENT_ID`, `ONENOTE_AUTHORITY`,
`FIRESTORE_CACHE_DOC`, and `GOOGLE_CLOUD_PROJECT` — not the `MCP_OAUTH_*` values — and it
requires the last two rather than defaulting them. See [Bootstrap](#bootstrap).

## Repository hygiene

This repository is public. Do not commit real page content, rendered ink, Entra tenant
names or IDs, or Firestore document contents. `.gitignore` excludes `output/` and the
token-cache file patterns; see the "Repo hygiene" section of `project-spec.md`.
