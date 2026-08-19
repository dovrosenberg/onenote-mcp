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
| `no-account` | The cache was read but holds no signed-in account | `npm run bootstrap` |
| `silent-failed` | The stored refresh token is expired or revoked, or the token endpoint returned nothing usable | `npm run bootstrap` |

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
`list_pages_by_name`), one reading tool (`get_page_content`), and three writing tools
(`append_to_page`, `create_page`, `update_page_title`) — and `src/mcp-server.ts` is the
JSON-RPC surface around them.

A tool that throws comes back as a tool result with `isError: true` and a readable
message — an expired refresh token, a page that is gone, and a document resvg rejects are
all normal outcomes, not protocol faults. Only a call to a tool that was never registered
is a JSON-RPC error.

Every request writes one JSON log line: the HTTP verb, the path, the status, the
duration, the JSON-RPC method, and the tool name on a `tools/call`. Never the query
string, the headers, the arguments, or the result — see `src/logging.ts`.

Nothing on `/mcp` is authenticated yet. Issue #23's bearer-token middleware goes in front
of the router; `/healthz` stays open.

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

`POST /token` issues an access token good for one hour and a refresh token good for 30
days. Both are an HMAC-SHA256 over a compact payload under `MCP_TOKEN_SIGNING_KEY` and
nothing else — no store is consulted to verify one, which is what keeps a Cloud Run
revision replacement from forcing a reconnect. The payload carries the audience, which is
`MCP_PUBLIC_URL` plus `/mcp`, so a token is good for this MCP endpoint and no other. How
long those tokens live, and how to make a human approve more often, is the next section.

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

Run it again whenever `GraphAuthError` appears in the server's logs. The refresh token is
rotated on every use and dies if the service sits idle past roughly 90 days; there is no
automatic recovery.

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

curl -i localhost:8080/healthz    # 200, {"status":"ok",...}
```

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
