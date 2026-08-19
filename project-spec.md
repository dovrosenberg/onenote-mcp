# OneNote MCP Server — Build Spec

## Scope

A standalone MCP server that exposes Microsoft OneNote via Microsoft Graph.
It reads notebook/section/page structure, returns page content (typed text
**and** rendered handwriting), and writes to pages.

It should work correctly and be independently testable — callable from
Claude Desktop or Claude.ai chat with no assumptions about who's calling or
why.

This spec is self-contained. Appendix A contains a complete, working recon
script that already validates the auth flow, the structure traversal, and
the handwriting reconstruction — port it rather than rebuilding from
scratch.

---

## The one hard problem: handwriting

Most of this server is a straightforward Graph API wrapper. The part that
isn't — and the reason this server needs to exist rather than using an
off-the-shelf OneNote MCP — is ink.

The target notebooks mix typed text with handwriting (written on an Android
tablet with a stylus). Graph's normal page-content endpoint drops
handwriting entirely, replacing it with `<!-- InkNode is not supported -->`
comments. Graph also cannot export a page as an image or PDF. So handwriting
has to be reconstructed from raw stroke data.

**This is solved and validated.** `GET
/me/onenote/pages/{id}/content?includeInkML=true` returns a
`multipart/mixed` response. One part is the normal HTML with the useless
`InkNode` placeholders. The other part is well-formed InkML containing real
stroke data. Confirmed working approach (see Appendix A for the
implementation):

- Strip XML namespaces — Graph returns everything as `inkml:ink`,
  `inkml:trace`, etc.
- Read `<traceFormat><channel name="X".../>` for actual channel order and
  units. Do **not** assume the first two numbers per point are X,Y — this
  account's data is X, Y, F (pressure), in that order.
- Convert `himetric` units to px: `px = himetric * 96 / 2540`.
- A page may contain more than one `<ink>` root, and `<traceGroup>` elements
  can nest — collect all traces from anywhere in the tree.
- Reconstruct strokes as SVG paths, then rasterize to PNG
  (`@resvg/resvg-js`; no headless browser needed for this part).

The resulting PNG is returned to the caller as an image, and the calling
model reads the handwriting with its own vision. **No OCR service, no
handwriting-recognition API, no paid dependency.** Tested against a real
page with 440 strokes; rendered fully legible.

### Coordinate spaces line up

Not needed for anything in the current tool surface, but worth recording:
OneNote's page HTML positions elements with absolute pixel coordinates (e.g.
`style="position:absolute;left:463px;top:1531px;width:624px"`, under a
`<body data-absolute-enabled="true">`). The `himetric → px @96dpi`
conversion above puts ink into **that same coordinate space**. So if a
future tool ever needs typed content and ink registered against each other,
the alignment is arithmetic, not guesswork.

---

## Other validated groundwork (reuse, don't re-derive)

- **Auth**: MSAL Node device-code flow against Microsoft Graph, delegated
  permissions, tested against the **Alar Insights** Entra ID tenant.
  Required scopes: `Notes.Read` and `Notes.ReadWrite`.
- **Never call `/me/onenote/pages`** (the account-wide page list). It fails
  with error 20266, "maximum sections exceeded", once there are enough
  sections across all notebooks — which a per-year-notebook,
  section-per-month structure hits easily. Always scope page listing to a
  specific section (`/sections/{id}/pages`).
- **Structure is nested**: notebooks contain sections *and* section groups
  ("tab groups" in the UI), and section groups can contain further sections
  and section groups. Browsing has to handle that recursion.
- **Packages that work**: `@azure/msal-node`, `fast-xml-parser` (with
  `removeNSPrefix: true`), `@resvg/resvg-js`.

### Target notebook structure (context for tool design)

- One notebook per year.
- Each month is a section group ("tab group") containing sections for:
  monthly calendar/todo, weekly calendar/todo, daily todo, and topic notes.
- Pages mix typed and handwritten content, but **ink and typed content are
  generally independent** — handwriting isn't annotating or interleaved with
  typed paragraphs.
- Page titles are typed, not handwritten, so the OneNote `title` property
  can be relied on for matching. This is user-side discipline, not something
  the server enforces.

---

## Tools to expose

### Structure / browsing

| Tool | Purpose |
|---|---|
| `list_notebooks()` | All notebooks |
| `list_sections(containerType, containerId)` | Sections **and** section groups under a notebook or section group. `containerType` is `notebook` or `sectionGroup`. |
| `list_pages(sectionId, top?)` | Pages in one section, sorted by last-modified |
| `search_pages(query, sectionId?)` | Find pages by title. Scope to a section where possible; if unscoped, iterate sections rather than hitting the account-wide endpoint. |
| `find_page_by_name(notebookName, sectionGroupName?, sectionName, pageTitle)` | The pages in that section whose title matches, for a caller that already knows the names |
| `list_pages_by_name(notebookName, sectionGroupName?, sectionName, top?)` | Every page in a section named the same way |

**Design notes for the two `_by_name` tools:**
- They exist because every other tool takes an id, so a caller holding names
  spends three round trips converting them. The account allows 400 Graph
  requests an hour, so those round trips are not free.
- The container path resolves in one request:
  `GET /me/onenote/notebooks?$select=id,displayName&$expand=sections($select=id,displayName),sectionGroups($select=id,displayName;$expand=sections($select=id,displayName))`.
  Measured at 78 KB for 54 notebooks. It reaches a notebook's sections and one
  level of section group, which is Graph's cap on `$expand` nesting. A section
  nested deeper is found by one filtered request against the account-wide
  section list with its parents expanded, which runs only when the cheap answer
  comes back empty.
- Page titles are matched by Graph, not read and compared locally:
  `/sections/{id}/pages?$filter=tolower(title) eq '…'`. Nothing bounds how many
  pages a section may hold before a match could be missed.
- Container names go through a ladder: exact and case-insensitive, then the
  same against the stored name with a leading ordering prefix removed, then a
  substring. The middle rung is what lets `February` find the section group
  named `062 - February`, which is how this account names its months. A rung is
  tried only when the one above it matched nothing, and `matchedBy` in the
  result says which one answered. Page titles stay exact; `search_pages` is the
  substring tool for those.
- A name that matches nothing is an error listing the names that were there,
  not an empty result. A caller cannot tell an empty section from a
  non-existent one otherwise.
- Ambiguity is reported with the candidate ids, never resolved by picking one.
- `sectionGroupName` omitted means the section sits directly in the notebook.
  It is not a wildcard.

### Reading

| Tool | Purpose |
|---|---|
| `get_page_content(pageId)` | Returns `{ html, ink_image }` — `html` is the typed content, `ink_image` is a PNG of the reconstructed handwriting (or null if the page has no ink). Both come from a single `includeInkML=true` fetch. |

**Design notes for `get_page_content`:**
- Return both parts in one call. A caller shouldn't have to make two round
  trips or know that ink is a separate concern.
- Return `ink_image` as an actual MCP image content block, not a file path
  or a base64 blob buried in text — the point is for the calling model to
  *see* it.
- If the page has no ink, `ink_image` is null. That's normal, not an error.
- **Cropping all of a page's ink into one bounding-box image is correct
  here.** Since ink and typed content are independent on these pages, no
  spatial relationship between them needs preserving. Don't build machinery
  to interleave ink fragments into the HTML.
- Consider trimming the HTML — Graph's page markup carries a lot of inline
  styling noise. Worth stripping to something readable rather than passing
  raw markup through.

### Writing

| Tool | Purpose |
|---|---|
| `append_to_page(pageId, htmlFragment)` | Appends content to the end of a page's body |
| `create_page(sectionId, title, htmlFragment)` | Creates a new page in a section |
| `update_page_title(pageId, newTitle)` | Changes a page's title |

**Design notes for writing:**
- OneNote's Graph write model is a PATCH to `/pages/{id}/content` with an
  array of actions targeting elements on the page — not a simple text
  append. Appending to the body means targeting the `body` element with
  action `append`. This is the least intuitive part of the Graph OneNote
  API; get it right early. The three request bodies below were confirmed
  against the live service on 2026-08-18 (issue #17). The full measured
  behaviour, including every error code, is in `api-overview.md` under
  **Writing page content**.

  ```jsonc
  // append_to_page
  [{ "target": "body", "action": "append", "content": "<p>appended</p>" }]

  // update_page_title
  [{ "target": "title", "action": "replace", "content": "New title" }]

  // insert above or below a known element (issue #27)
  [{ "target": "#beta", "action": "insert", "position": "before", "content": "<p>above</p>" }]
  ```

- `update_page_title` is `target: "title"`, `action: "replace"`, and the
  content string becomes the title verbatim. Nothing in it is parsed:
  `<p>x</p>` produces a title that reads `<p>x</p>`, and `&`, `<` and `"`
  survive unescaped. No other action works on the title — `append` is 400
  with code `20141` — and the target must be written `title`, never
  `#title`.
- **`target: "body"` is the first top-level div, which on a real page is the
  first outline and not the page.** Pages authored in the OneNote client
  have several sibling top-level divs; a `body` append lands at the end of
  the first one, so `append_to_page` appends to the first outline. Writing
  anywhere else means reading `?includeIDs=true` and targeting that div's
  generated id. Generated ids change on every update, so they must be read
  in the same operation that uses them.
- **One bad change rejects the whole array.** A request holding a valid
  change and one naming a missing target returned 400 and applied neither.
  A multi-change PATCH is therefore all-or-nothing, which is what a tool
  wanting title-plus-body in one call should rely on.
- Writes must not clobber existing content, **including ink**. Verify on a
  test page that appending to a page containing handwriting leaves the
  handwriting intact. This is the single most important thing to test.
- `create_page` must set a real page title (the `<title>` element in the
  submitted HTML), not just heading text in the body.
- **A page created here omits `data-absolute-enabled` from its `<body>`.**
  Graph then wraps the submission in one `<div data-id="_default">`, so
  `body` addresses the whole page and `append_to_page` reaches the bottom of
  it. Setting the attribute would give sibling outlines like a
  client-authored page and put every append in the first one. The two page
  shapes cannot both behave well; this picks the one this server creates.

---

## Deployment

Runs as a persistent server over **Streamable HTTP** transport (not
local-only stdio) so it's reachable from contexts that can't touch a
device-local process.

Config via environment variables: Azure client ID, tenant/authority, the
Layer-1 OAuth client ID and secret, the access-token signing key, the
Firestore document path for the token cache, and the bind port (`PORT`).
Cloud Run sets `PORT` itself; the server must bind to it rather than a
hardcoded value.

### Deployment target: Google Cloud Run via GitHub Actions

The core tension: Graph's OneNote endpoints don't support app-only
(client-credentials) auth, so this must use **delegated** auth — which
requires a one-time interactive sign-in and a refresh token that must
survive across restarts. Cloud Run containers are ephemeral and stateless,
so the token cache cannot live on the container filesystem.

That is resolved by putting the token cache in Firestore. The decisions
below are settled; they are recorded here so they aren't re-litigated during
the build.

| Concern | Decision |
|---|---|
| Runtime | Cloud Run, single service, `--max-instances=1` |
| Transport | Stateless Streamable HTTP. Not SSE — SSE holds an instance open, which fights scale-to-zero and bills for idle time. |
| Image build | `docker build` inside the GitHub Actions job on an `ubuntu-latest` runner, pushed to Artifact Registry, deployed by digest |
| Deploy credential | Workload Identity Federation. No service-account JSON key is created or stored anywhere. |
| Static secrets | GitHub repository secrets, injected as Cloud Run env vars at deploy time |
| MSAL token cache | Firestore in Native mode, one document |
| Secret Manager | Not used |

Building the image in the Actions job rather than with `gcloud run deploy
--source` is deliberate. The runner is `linux/amd64`, which is what Cloud Run
runs, so the platform-specific `@resvg/resvg-js` binary matches without a
`--platform` flag or a buildx setup. It also avoids enabling Cloud Build and
granting the roles it needs.

#### Where each piece of configuration lives

| Value | Secret? | Stored in | Reaches the service as |
|---|---|---|---|
| Azure app registration client ID | No | GitHub repo variable | Env var `ONENOTE_CLIENT_ID` |
| Azure authority / tenant | No | GitHub repo variable | Env var `ONENOTE_AUTHORITY` |
| Layer-1 OAuth client ID (Claude → this server) | No | GitHub repo variable | Env var `MCP_OAUTH_CLIENT_ID` |
| Layer-1 OAuth client secret | **Yes** | GitHub repo secret | Env var `MCP_OAUTH_CLIENT_SECRET` |
| Access-token signing key | **Yes** | GitHub repo secret | Env var `MCP_TOKEN_SIGNING_KEY` |
| The service's own public URL (#21) | No | GitHub repo variable | Env var `MCP_PUBLIC_URL` |
| WIF provider resource name | No | GitHub repo variable | Used by the Action only |
| Deploy service account email | No | GitHub repo variable | Used by the Action only |
| MSAL token cache | **Yes** | Firestore document | Read and written at runtime by the service |
| GCP project / region / service name | No | GitHub repo variable | Used by the Action only |

Only two things in GitHub are actually secret. The WIF provider name and the
service account email are identifiers, not credentials — they are useless
without the repository's OIDC identity, so they belong in repo *variables*,
not secrets.

Cloud Run's `--set-env-vars` replaces the entire environment set rather than
merging into it. The workflow must therefore set every variable on every
deploy, or use `--update-env-vars`. The `google-github-actions/deploy-cloudrun`
action exposes an update-strategy input for this — confirm its behavior when
writing the workflow rather than assuming.

Env vars set this way are visible in the Cloud Run revision spec to anyone
with `roles/run.viewer` on the project. For a single-user project where the
owner is the only principal, that is acceptable. It is the tradeoff accepted
in exchange for not running Secret Manager.

#### Token cache: Firestore, not Secret Manager

The cache is mutable state that is rewritten on every token refresh, because
Microsoft issues rotating refresh tokens. That write pattern is the deciding
factor.

| | Secret Manager | Firestore |
|---|---|---|
| Write semantics | `addSecretVersion` creates a new version every time | `doc.set()` overwrites in place |
| Ongoing maintenance | Must list and destroy superseded versions; active versions are billed per version per month and there is a per-secret version quota | None |
| Concurrency control | No compare-and-set against "latest" | Transactions |
| Failure mode if neglected | Versions accumulate silently until billing or quota bites | None |

Firestore setup is two commands:

```bash
gcloud firestore databases create --location=us-central1 --type=firestore-native
gcloud projects add-iam-policy-binding "$PROJECT" \
  --member="serviceAccount:$RUNTIME_SA" --role=roles/datastore.user
```

#### Why the token cache cannot live in GitHub

GitHub Secrets are write-only through the API — a workflow can read them,
but the REST API will not return a secret's value. Writing one requires a
token with repository-admin scope plus libsodium encryption against the
repository public key. Putting that in the path of every token refresh is
not viable.

Reading a cache from GitHub at boot and never writing back also fails. After
the first refresh the stored refresh token is superseded, and whether Entra
ID keeps honoring the old one is undocumented. Do not build on it.

#### Constraints that still apply

1. **Refresh-token rotation must be persisted.** Every refresh returns a new
   refresh token and the service must write the updated cache back to
   Firestore. `--max-instances=1` removes most of the concurrent-write risk;
   use a Firestore transaction for the read-modify-write anyway, since
   revision transitions can briefly overlap instances.
2. **Bootstrap sign-in happens locally, never on Cloud Run.** A local CLI runs
   the device-code flow and writes the resulting cache into Firestore. The
   deployed service only ever refreshes.
3. **Refresh tokens expire if unused.** Microsoft's delegated refresh tokens
   generally last around 90 days and renew on use, so an actively-used
   service stays alive indefinitely. A service idle past that window needs the
   local bootstrap re-run. Conditional access policies can shorten the window.
   Emit a clear, specific error when refresh fails rather than a bare 401.
4. **Do not deploy with unauthenticated `allUsers` access at the IAM layer
   without the Layer-1 OAuth described below.** A public Cloud Run URL
   proxying an entire OneNote account is a serious exposure. Claude custom
   connectors cannot present Google IAM credentials, so the endpoint must be
   IAM-public and protected by the server's own OAuth — which means that OAuth
   has to work before the service is exposed.

#### Workflow shape

```yaml
permissions:
  contents: read
  id-token: write   # required for Workload Identity Federation

steps:
  - uses: actions/checkout@v4
  - uses: google-github-actions/auth@v2
    with:
      workload_identity_provider: ${{ vars.WIF_PROVIDER }}
      service_account: ${{ vars.DEPLOY_SA }}
  - uses: google-github-actions/setup-gcloud@v2
  - run: gcloud auth configure-docker ${{ vars.GAR_REGION }}-docker.pkg.dev
  - run: |
      IMAGE="${{ vars.GAR_REGION }}-docker.pkg.dev/${{ vars.GCP_PROJECT }}/onenote-mcp/server:${{ github.sha }}"
      docker build -t "$IMAGE" .
      docker push "$IMAGE"
      echo "IMAGE=$IMAGE" >> "$GITHUB_ENV"
  - uses: google-github-actions/deploy-cloudrun@v2
    with:
      service: onenote-mcp
      region: ${{ vars.GCP_REGION }}
      image: ${{ env.IMAGE }}
      flags: --max-instances=1 --allow-unauthenticated
      env_vars: |
        ONENOTE_CLIENT_ID=${{ vars.ONENOTE_CLIENT_ID }}
        ONENOTE_AUTHORITY=${{ vars.ONENOTE_AUTHORITY }}
        MCP_OAUTH_CLIENT_ID=${{ vars.MCP_OAUTH_CLIENT_ID }}
        MCP_OAUTH_CLIENT_SECRET=${{ secrets.MCP_OAUTH_CLIENT_SECRET }}
        MCP_TOKEN_SIGNING_KEY=${{ secrets.MCP_TOKEN_SIGNING_KEY }}
        FIRESTORE_CACHE_DOC=tokencache/msal
```

#### Alternative not taken

A home server behind a Cloudflare Tunnel would remove the Firestore work
entirely, since a local file cache just works. It was not chosen because it
ties uptime to home hardware. The Firestore cache plugin plus the bootstrap
CLI is the cost of that independence, and it is bounded — one module and one
script.

### Two independent OAuth layers — don't conflate them

This is the most important thing to get straight before writing auth code.
There are **two separate OAuth relationships**, and they share nothing:

| | Layer 1: Claude → this server | Layer 2: this server → Microsoft Graph |
|---|---|---|
| Who's the client | Claude | This server |
| Who's the authorization server | This server | Microsoft Entra ID |
| Credentials | An OAuth client ID + secret this server issues, entered into Claude's connector config | The Azure app registration's client ID (public client, no secret) |
| Flow | Authorization code, browser-based, at connect time | Device code once at bootstrap, then silent refresh forever |
| What it protects | The MCP endpoint itself, from anyone who finds the URL | The user's OneNote data |

Layer 2 is already validated (Appendix A). Layer 1 is new work.

**What Layer 1 requires the server to implement:**

- **An OAuth 2.1 authorization server**, or a minimal implementation of one.
  Claude custom connectors have an "Advanced settings" field for an OAuth
  Client ID and Client Secret — generate a pair, hardcode/config them into
  the server, enter them in Claude. Claude then runs a standard
  authorization-code flow against this server.
- **OAuth metadata discovery endpoints** so Claude can find the authorization
  server (`/.well-known/oauth-authorization-server`, and protected-resource
  metadata per the MCP auth spec). Claude supports the 2025-03-26,
  2025-06-18 and 2025-11-25 MCP auth specs.
- **An authorize endpoint with a real consent step.** Anthropic's docs are
  explicit that a pure machine-to-machine `client_credentials` grant is
  **not** supported — every connection requires user consent, which means a
  browser-visible approve screen. Even for a single-user server, there has
  to be something to click. Keep it trivial (a single "Approve" button), but
  it must exist.
- **A token endpoint** issuing access tokens, and **bearer token validation
  on every MCP request**. Never accept a token in a query string — the MCP
  auth spec prohibits it, and URLs leak into logs and proxies.

**Practical notes:**

- Claude attempts Dynamic Client Registration by default when no client
  ID/secret is supplied. Supplying them explicitly (the Advanced settings
  path) sidesteps DCR, which is the simpler route here — implementing DCR
  for a single-user server is wasted effort.
- Anthropic's connector documentation now lists `none` (an authless server)
  as a supported authentication type, so the connector UI is no longer the
  reason to build Layer 1. The reason is that the Cloud Run URL is public
  and deployed `--allow-unauthenticated`, and OneNote content sits behind
  it.
- Since this is single-user, the "user database" is one user. Don't build
  user management — a single configured credential pair, a trivial consent
  screen, and signed/opaque access tokens with an expiry is enough.
- The TypeScript MCP SDK covers most of the OAuth machinery. What it covers
  and what it leaves is measured in the next section.

### Layer 1: what the SDK provides — settled by the spike in issue #20

**Decision: use the SDK's authorization-server router and its bearer
middleware. Do not hand-roll an OAuth server.** `mcpAuthRouter` and
`requireBearerAuth` from `@modelcontextprotocol/sdk` 1.30.0 carry the
protocol: metadata documents, PKCE verification, redirect-URI matching,
client authentication, error shapes, and the `WWW-Authenticate` header.
What is left is an `OAuthServerProvider` implementation — the consent
screen, the token format, and the stores. The named gaps are below; each
one is work for #21, #22 or #23, not a reason to write the router.

#### Measured, 2026-08-18

The lines below come from a probe that mounted `mcpAuthRouter` over a
minimal provider and drove it end to end. They are observations of SDK
1.30.0, not readings of its documentation.

| Probe | Result |
|---|---|
| `GET /.well-known/oauth-authorization-server` | 200. `code_challenge_methods_supported: ["S256"]`, `grant_types_supported: ["authorization_code","refresh_token"]`, no `registration_endpoint` |
| `GET /.well-known/oauth-protected-resource/mcp` | 200, `resource` and `authorization_servers` as configured |
| `GET /.well-known/oauth-protected-resource` | 404 |
| `GET /.well-known/openid-configuration` | 404 |
| `GET /authorize` → `https://claude.ai/api/mcp/auth_callback` | 302 carrying `code` and `state` |
| `GET /authorize` → `http://localhost:3118/callback`, registered as `http://localhost/callback` | 302 — RFC 8252 port relaxation already implemented |
| `GET /authorize` → an unregistered redirect URI | 400 `invalid_request`, no redirect |
| `POST /token`, credentials in the form body | 200 with the provider's token response |
| `POST /token`, credentials in an `Authorization: Basic` header | 400 `invalid_request` — HTTP Basic client authentication is not implemented |
| `POST /token` with a JSON body | 400 — form-urlencoded only, which is what Claude sends |
| `POST /mcp` with no token | 401, `WWW-Authenticate: Bearer error="invalid_token", error_description="…", resource_metadata="…"` |
| `POST /mcp?access_token=…` | 401 — the query string is never read |
| `POST /mcp` with a token whose `resource` names another server | **200** |

#### The gaps, and where each one lands

1. **No audience check.** The last probe line is the important one: a token
   minted for a different resource server was accepted. `AuthInfo.resource`
   is carried but nothing compares it. `checkResourceAllowed` exists in
   `shared/auth-utils.js` and is called only by the SDK's client. Binding
   the audience is `verifyAccessToken`'s job, and it is #23's work.
2. **No consent screen.** `provider.authorize` is handed the response object
   and must eventually redirect with a code. Rendering a form and resuming
   after the POST is entirely ours, and the resume path needs a route the
   SDK's `/authorize` router does not own.
3. **No token format, no stores, no refresh.** The SDK issues nothing. The
   demo provider it ships (`examples/server/demoInMemoryOAuthProvider.js`)
   keeps codes and tokens in a `Map` and throws on refresh.
4. **Metadata advertises `none` in `token_endpoint_auth_methods_supported`
   unconditionally.** `createOAuthMetadata` builds that array from nothing
   the caller controls, so a confidential-client-only server still says it
   accepts public clients. The enforcement is correct regardless — a client
   record holding a secret makes the secret mandatory — so this is a
   truthfulness problem in the document, not a hole.

   **Closed by #21.** `src/oauth-router.ts` registers
   `/.well-known/oauth-authorization-server` ahead of `mcpAuthRouter` and
   serves `client_secret_post` alone. The document is still the SDK's own
   `createOAuthMetadata` output with that one key replaced, and it is still
   served by the SDK's `metadataHandler`, so the CORS headers and the
   GET/OPTIONS-only method handling are identical to the route it replaces
   and the two documents cannot drift apart as the SDK changes. The reason
   to bother: the document is what a client reads to decide whether to send
   the secret at all.
5. **Protected-resource metadata is served only at the path-suffixed URL.**
   With the resource at `…/mcp`, `/.well-known/oauth-protected-resource`
   itself is a 404. Claude probes the suffixed path first and only falls
   back to the bare one, and the `401` will point at the suffixed URL
   explicitly, so nothing here breaks. A client that probes only the bare
   path fails to discover anything.
6. **The client secret is compared with `!==`.** Issue #22 asks for a
   constant-time comparison; the comparison lives inside the SDK's
   `authenticateClient` middleware and cannot be replaced without replacing
   the whole token router. Recommendation: amend #22. The channel is a
   string comparison behind TLS, reachable only from the public internet,
   under a 50-request rate limit — measuring it is not practical, and
   replacing the SDK's token endpoint to close it costs more than it buys.
7. **Rate limiting keys on the socket address.** The SDK applies
   `express-rate-limit` at 100 requests per 15 minutes on `/authorize` and
   50 on `/token`. Behind Cloud Run every request arrives from the front
   end, and a probe with `X-Forwarded-For` set did not change the counter,
   so all callers share one bucket. Anyone who finds the URL can exhaust the
   token endpoint's budget and block the operator's own connect flow for 15
   minutes. Neither raising `trust proxy` nor per-IP keying fixes that for a
   single-user server.

   **#21 made the shared bucket explicit and stopped the log noise.** With
   `trust proxy` set to true — which this service needs behind Cloud Run —
   `express-rate-limit` derives its key from `X-Forwarded-For`, which any
   caller can set, and prints a `ValidationError` stack trace on every
   `/authorize` and `/token` request rather than a log line. Observed while
   wiring the mount, not predicted. `src/oauth-router.ts` therefore passes a
   constant `keyGenerator` to both handlers: one global bucket, which is
   what the measurement above already showed, with no forgeable key and no
   stack trace. The lockout the paragraph above describes is unchanged and
   still accepted.
8. **`mcpAuthRouter` must be mounted at the application root.** It builds
   its paths from the issuer URL, not from a mount point.

#### What Claude does, from Anthropic's connector documentation

Read 2026-08-18. These are the constraints the implementation has to meet.

- **Spec revisions**: 2025-03-26, 2025-06-18 and 2025-11-25 are all
  supported.
- **Dynamic Client Registration is skippable**, which settles the third task
  on #20. A custom connector added by URL takes an OAuth Client ID, and the
  Client Secret field beside it is optional — supply it only if the server
  requires confidential-client authentication. Ours does. Pre-registered
  credentials are one of the three client-identity mechanisms Claude
  accepts, alongside DCR and Client ID Metadata Documents.
- **Callback URLs**: `https://claude.ai/api/mcp/auth_callback` for
  Claude.ai, Desktop, mobile and Cowork. Claude Code is a native client on
  an RFC 8252 loopback redirect with an ephemeral port, declaring
  `http://localhost/callback` and `http://127.0.0.1/callback`; both must be
  accepted with the port ignored. The SDK's matcher already does this, so
  registering all three redirect URIs makes both surfaces work.
- **PKCE `S256` on every authorization request**, and the metadata must
  advertise it.
- **The `resource` parameter is sent** on authorization and token requests,
  set to the canonical form of the MCP server URL including its path. The
  token has to be bound to it, and the check should accept the canonical
  value rather than a byte comparison against what the user typed.
- **`offline_access` is appended to the requested scopes** when the
  authorization server metadata lists it in `scopes_supported`. That is the
  switch that decides whether Claude asks for a refresh token.
- **Refresh happens reactively on a 401**, plus proactively up to five
  minutes before the stored expiry. A dead refresh token must come back as
  `invalid_grant`.
- **The `401` must carry `WWW-Authenticate` with a `resource_metadata`
  parameter.** Claude does not honour that header on a 200, and a tool
  result carrying an "isError" sign-in message produces no auth prompt at
  all — it reaches the model as an ordinary tool failure.
- **Endpoint latency**: 10 seconds for discovery and token requests, 30 for
  refresh. A Cloud Run cold start plus an HMAC is well inside that.
- **Anthropic's egress range is `160.79.104.0/21`**, if anything in front of
  the service ever needs an allowlist.

#### The decisions this settles for #21–#23

- **Token format: HMAC-SHA256 over a compact payload, using
  `MCP_TOKEN_SIGNING_KEY` and `node:crypto`.** No new dependency, and
  nothing outside this server ever reads one. The payload carries the
  audience, the expiry, and the client id.
- **Access tokens are stateless and last an hour. Refresh tokens are
  stateless and last 30 days.** Stateless is what makes a revision
  replacement invisible: an in-memory token store would force a reconnect
  every deploy.
- **Refresh tokens are not rotated.** OAuth 2.1 and the MCP spec require
  rotation for *public* clients; supplying the secret in Claude's Advanced
  settings makes this a confidential client, and rotation cannot be
  enforced by a stateless token anyway. Say so in the code.
- **Authorization codes stay in memory**, single-use, 60-second lifetime.
  Losing one to a revision replacement costs a retry of the consent click,
  which is the one moment a human is already present.
- **The authorization parameters cross the consent screen signed**, in a
  hidden form field, with the same key. That keeps the flow working across
  an instance replacement mid-consent and stops the form being edited.
- **The consent screen authenticates nobody.** It has one Approve button.
  What protects the endpoint is that `/token` requires the client secret,
  that the redirect URI allowlist sends every code to `claude.ai` or to
  loopback, and that PKCE binds the code to the client that started the
  flow. An attacker holding the client id can reach the consent screen and
  cause a code to be minted; without the secret they cannot exchange it.
- **A new configuration variable is needed: `MCP_PUBLIC_URL`**, in the
  `oauth` group. The issuer, the `resource` value, and the
  `resource_metadata` URL all have to be the service's public URL, and
  nothing on Cloud Run tells the process what that is. The deploy workflow
  gains `MCP_PUBLIC_URL=${{ vars.MCP_PUBLIC_URL }}`, which can only be
  filled in after the first deploy has produced the URL.
- **#23's exempt list is longer than the issue assumes.** Everything
  `mcpAuthRouter` mounts is unauthenticated by necessity: `/authorize`,
  `/token`, both `.well-known` documents, and the consent POST, alongside
  `/healthz`. The fail-closed route test has to enumerate them.

---

## Open questions to resolve during build

- **Multi-page reads**: if callers routinely need several pages at once, a
  batch variant of `get_page_content` may be worth adding rather than
  making N calls.

Settled by the spike in issue #17, on 2026-08-18:

- **Insert at top of page** — supported, and cheap. `action: "prepend"` on
  `target: "body"` puts a block above existing content, and
  `action: "append"` with `"position": "before"` does the same thing. Both
  returned 204. Positional targeting relative to a named element also works:
  `action: "insert"` with `"position": "before"` or `"after"` against a
  `#{data-id}`. Nothing has to rewrite full page content. A `position`
  parameter on `append_to_page` is the straightforward shape.
- **`update_page_title` mechanics** — settled; see the writing notes above.
- **Ink and writes** — settled 2026-08-18 by issue #19's test,
  `test/ink-preservation.integration.test.ts`, run against a scratch page with 5
  hand-written strokes on it. `append_to_page` and `update_page_title` both left
  the stroke count and the rendered PNG bytes identical. Writes do not clobber
  handwriting.

---

## Repo hygiene

If this becomes a public repo: `.gitignore` must exclude the token cache and
any output/scratch directories. Real page dumps contain personal notes —
rendered ink images are fully legible — so no real API responses or rendered
output should ever be committed. Any example output in docs should come from
a throwaway test page with fake content.

The Azure client ID is not a secret in this design (public client, device
code flow, no client secret), so it's safe in config examples. Two things are
secret and must never appear in source: the **MSAL token cache**, which lives
only in Firestore, and the **connector-facing OAuth client secret** from
Layer 1, which lives only in GitHub repository secrets and reaches the
service as an environment variable.

This repository is public. Issues, pull requests, and Actions logs are public
with it. Do not put the Entra tenant name, tenant ID, Firestore document
contents, or any real page content into an issue or a workflow log. Mask
values in workflow output rather than echoing them.

---

## Appendix A: Working recon script

This script is validated and working. It authenticates, browses
notebook → section group(s) → section → page, fetches both the plain HTML
and the InkML, dumps everything for inspection, and renders the ink to SVG
and PNG. The auth, traversal, and ink-parsing functions here should be
lifted directly into the server.

### `package.json`

```json
{
  "name": "onenote-ink-recon",
  "version": "0.1.0",
  "private": true,
  "type": "module",
  "scripts": {
    "start": "node index.mjs"
  },
  "dependencies": {
    "@azure/msal-node": "^2.13.1",
    "@resvg/resvg-js": "^2.6.2",
    "fast-xml-parser": "^4.5.7"
  },
  "engines": {
    "node": ">=18"
  }
}
```

### `index.mjs`

```javascript
// OneNote ink reconnaissance script.
//
// 1. Signs in via device code flow (delegated auth).
// 2. Browses notebook -> tab group(s) -> section -> pages to pick one.
// 3. Pulls that page's normal HTML content, AND its content with
//    includeInkML=true.
// 4. Dumps everything raw to disk.
// 5. Finds <trace> elements in the response and renders them as SVG + PNG.

import { PublicClientApplication } from "@azure/msal-node";
import { XMLParser } from "fast-xml-parser";
import { Resvg } from "@resvg/resvg-js";
import fs from "node:fs";
import path from "node:path";
import readline from "node:readline/promises";
import { stdin, stdout } from "node:process";

// ---------------------------------------------------------------------------
// Config
// ---------------------------------------------------------------------------

const CLIENT_ID = process.env.ONENOTE_CLIENT_ID || "";
const AUTHORITY = "https://login.microsoftonline.com/common"; // personal + work/school
const SCOPES = ["https://graph.microsoft.com/Notes.Read"];
const GRAPH_ROOT = "https://graph.microsoft.com/v1.0";

if (!CLIENT_ID) {
  console.error("Missing ONENOTE_CLIENT_ID env var.");
  process.exit(1);
}

const OUT_DIR = path.join(process.cwd(), "output");
fs.mkdirSync(OUT_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// Auth (device code flow — prints a URL + code, you approve in a browser)
// ---------------------------------------------------------------------------

async function getAccessToken() {
  const pca = new PublicClientApplication({
    auth: { clientId: CLIENT_ID, authority: AUTHORITY },
  });

  const deviceCodeRequest = {
    scopes: SCOPES,
    deviceCodeCallback: (response) => {
      console.log("\n=== Sign in required ===");
      console.log(response.message);
      console.log("========================\n");
    },
  };

  const result = await pca.acquireTokenByDeviceCode(deviceCodeRequest);
  return result.accessToken;
}

// ---------------------------------------------------------------------------
// Graph helpers
// ---------------------------------------------------------------------------

// The account-wide /me/onenote/pages endpoint fails (error 20266, "maximum
// sections exceeded") once you have enough sections across all notebooks.
// Always browse: notebook -> (section group ->)* section -> pages.

async function graphGet(token, url) {
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!res.ok) {
    throw new Error(`GET ${url} failed: ${res.status} ${await res.text()}`);
  }
  return res.json();
}

async function listNotebooks(token) {
  const data = await graphGet(
    token,
    `${GRAPH_ROOT}/me/onenote/notebooks?$select=id,displayName&$orderby=displayName`
  );
  return data.value;
}

// containerKind is "notebooks" or "sectionGroups" — both expose the same
// child shapes (sections + nested sectionGroups) under those relationship names.
async function listSections(token, containerKind, containerId) {
  const data = await graphGet(
    token,
    `${GRAPH_ROOT}/me/onenote/${containerKind}/${containerId}/sections?$select=id,displayName&$orderby=displayName`
  );
  return data.value;
}

async function listSectionGroups(token, containerKind, containerId) {
  const data = await graphGet(
    token,
    `${GRAPH_ROOT}/me/onenote/${containerKind}/${containerId}/sectionGroups?$select=id,displayName&$orderby=displayName`
  );
  return data.value;
}

async function listPagesInSection(token, sectionId, top = 50) {
  const data = await graphGet(
    token,
    `${GRAPH_ROOT}/me/onenote/sections/${sectionId}/pages` +
      `?$top=${top}&$orderby=lastModifiedDateTime desc&$select=id,title,lastModifiedDateTime`
  );
  return data.value;
}

async function promptChoice(items, label, formatItem) {
  console.log(`\n${label}`);
  items.forEach((item, i) => console.log(`  [${i}] ${formatItem(item)}`));
  const rl = readline.createInterface({ input: stdin, output: stdout });
  const answer = await rl.question("\nEnter number: ");
  rl.close();
  const chosen = items[Number(answer.trim())];
  if (!chosen) throw new Error("Invalid selection.");
  return chosen;
}

// Walks notebook -> section group(s) -> section, drilling through tab groups
// (which are section groups) until landing on an actual section.
async function browseToSection(token) {
  const notebooks = await listNotebooks(token);
  const notebook = await promptChoice(notebooks, "Pick a notebook:", (n) => n.displayName);

  let containerKind = "notebooks";
  let containerId = notebook.id;

  while (true) {
    const [sections, groups] = await Promise.all([
      listSections(token, containerKind, containerId),
      listSectionGroups(token, containerKind, containerId),
    ]);

    const items = [
      ...groups.map((g) => ({ ...g, kind: "group" })),
      ...sections.map((s) => ({ ...s, kind: "section" })),
    ];

    if (items.length === 0) {
      throw new Error("No sections or section groups found here.");
    }

    const choice = await promptChoice(
      items,
      "Pick a section (or a group/tab to drill into):",
      (item) => `${item.displayName}${item.kind === "group" ? "  (tab group)" : ""}`
    );

    if (choice.kind === "section") return choice;

    containerKind = "sectionGroups";
    containerId = choice.id;
  }
}

async function fetchPageHtml(token, pageId) {
  const res = await fetch(`${GRAPH_ROOT}/me/onenote/pages/${pageId}/content`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  const text = await res.text();
  return { ok: res.ok, status: res.status, contentType: res.headers.get("content-type"), text };
}

async function fetchPageInk(token, pageId) {
  const res = await fetch(
    `${GRAPH_ROOT}/me/onenote/pages/${pageId}/content?includeInkML=true`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  const text = await res.text();
  return { ok: res.ok, status: res.status, contentType: res.headers.get("content-type"), text };
}

// ---------------------------------------------------------------------------
// Multipart handling — the includeInkML response comes back as
// multipart/mixed; split it so each part can be inspected independently.
// ---------------------------------------------------------------------------

function splitMultipart(rawText, contentType) {
  const boundaryMatch = /boundary="?([^";]+)"?/i.exec(contentType || "");
  if (!boundaryMatch) return null; // not multipart

  const boundary = "--" + boundaryMatch[1];
  const rawParts = rawText.split(boundary).filter((p) => p.trim() && p.trim() !== "--");

  return rawParts.map((part) => {
    const headerBodySplit = part.indexOf("\r\n\r\n");
    const altSplit = headerBodySplit === -1 ? part.indexOf("\n\n") : -1;
    const splitAt = headerBodySplit !== -1 ? headerBodySplit + 4 : altSplit !== -1 ? altSplit + 2 : 0;

    const headerBlock = part.slice(0, splitAt).trim();
    const body = part.slice(splitAt).trim();
    const partContentType = /content-type:\s*([^\r\n]+)/i.exec(headerBlock)?.[1] ?? "unknown";

    return { headers: headerBlock, contentType: partContentType, body };
  });
}

// ---------------------------------------------------------------------------
// Ink extraction
// ---------------------------------------------------------------------------

const xmlParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  removeNSPrefix: true, // Graph returns inkml:ink, inkml:trace, etc.
  textNodeName: "#text",
});

function himetricToPx(v) {
  // himetric = 1/100 mm. Page HTML positions elements in px @ 96dpi.
  return (v * 96) / 2540;
}

// traceGroups can nest; walk the tree and collect every <trace> found.
function collectTraces(node, acc = []) {
  if (!node || typeof node !== "object") return acc;
  if (node.trace) {
    const traces = Array.isArray(node.trace) ? node.trace : [node.trace];
    acc.push(...traces);
  }
  if (node.traceGroup) {
    const groups = Array.isArray(node.traceGroup) ? node.traceGroup : [node.traceGroup];
    groups.forEach((g) => collectTraces(g, acc));
  }
  return acc;
}

// Reads <traceFormat><channel name="X" units="himetric" .../></traceFormat>
// to get point-value order and units, instead of assuming X,Y are the first
// two numbers (real data here is X,Y,F — F is pen pressure).
function getChannelInfo(inkRoot) {
  let contexts = inkRoot.definitions?.context;
  if (!contexts) return { order: ["X", "Y"], units: "himetric" };
  if (!Array.isArray(contexts)) contexts = [contexts];
  let channels = contexts[0]?.inkSource?.traceFormat?.channel;
  if (!channels) return { order: ["X", "Y"], units: "himetric" };
  if (!Array.isArray(channels)) channels = [channels];
  const order = channels.map((c) => c["@_name"]);
  const xChan = channels.find((c) => c["@_name"] === "X");
  return { order, units: xChan?.["@_units"] ?? "himetric" };
}

function tryParseInkML(text) {
  // A page can have more than one <ink>...</ink> root — collect all of them.
  const matches = [...text.matchAll(/<(\w+:)?ink[\s>][\s\S]*?<\/(\w+:)?ink>/gi)];
  if (matches.length === 0) return null;

  const allStrokes = [];
  for (const m of matches) {
    let inkRoot;
    try {
      const parsed = xmlParser.parse(m[0]);
      inkRoot = parsed.ink || Object.values(parsed)[0];
    } catch (err) {
      console.warn("Looked like <ink> but failed to parse:", err.message);
      continue;
    }
    if (!inkRoot) continue;

    const { order, units } = getChannelInfo(inkRoot);
    const xIdx = order.indexOf("X") >= 0 ? order.indexOf("X") : 0;
    const yIdx = order.indexOf("Y") >= 0 ? order.indexOf("Y") : 1;

    const traces = collectTraces(inkRoot);
    const strokes = traces
      .map((t) => (typeof t === "string" ? t : t["#text"]))
      .filter(Boolean)
      .map((raw) =>
        raw
          .trim()
          .split(",")
          .map((pointStr) => {
            const nums = pointStr.trim().split(/\s+/).map(Number);
            const x = nums[xIdx];
            const y = nums[yIdx];
            if (Number.isNaN(x) || Number.isNaN(y)) return null;
            return units === "himetric" ? [himetricToPx(x), himetricToPx(y)] : [x, y];
          })
          .filter(Boolean)
      )
      .filter((stroke) => stroke.length > 1);

    allStrokes.push(...strokes);
  }

  return allStrokes.length > 0 ? allStrokes : null;
}

function strokesToSVG(strokes) {
  let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;
  for (const stroke of strokes) {
    for (const [x, y] of stroke) {
      if (x < minX) minX = x;
      if (y < minY) minY = y;
      if (x > maxX) maxX = x;
      if (y > maxY) maxY = y;
    }
  }
  const margin = Math.max((maxX - minX) * 0.05, (maxY - minY) * 0.05, 5);
  const w = maxX - minX + margin * 2;
  const h = maxY - minY + margin * 2;
  const viewBox = `${minX - margin} ${minY - margin} ${w} ${h}`;

  const paths = strokes
    .map((stroke) => {
      const d = stroke.map(([x, y], i) => `${i === 0 ? "M" : "L"}${x.toFixed(2)},${y.toFixed(2)}`).join(" ");
      return `  <path d="${d}" fill="none" stroke="black" stroke-width="1" stroke-linecap="round" stroke-linejoin="round" />`;
    })
    .join("\n");

  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="${viewBox}" width="${w.toFixed(0)}" height="${h.toFixed(0)}">\n  <rect x="${(minX - margin).toFixed(2)}" y="${(minY - margin).toFixed(2)}" width="${w.toFixed(2)}" height="${h.toFixed(2)}" fill="white" />\n${paths}\n</svg>\n`;
}

// ---------------------------------------------------------------------------
// Main
// ---------------------------------------------------------------------------

async function main() {
  const filterArg = process.argv[2]; // optional: substring to filter page titles

  console.log("Requesting access token (delegated, device code flow)...");
  const token = await getAccessToken();
  console.log("Signed in.\n");

  const section = await browseToSection(token);
  console.log(`\nFetching pages in "${section.displayName}"...`);
  const pages = await listPagesInSection(token, section.id);

  const candidates = filterArg
    ? pages.filter((p) => p.title?.toLowerCase().includes(filterArg.toLowerCase()))
    : pages;

  if (candidates.length === 0) {
    console.error(`No pages found in "${section.displayName}".`);
    process.exit(1);
  }

  let chosen;
  if (candidates.length === 1) {
    chosen = candidates[0];
    console.log(`Using page: "${chosen.title}"`);
  } else {
    chosen = await promptChoice(
      candidates,
      "Pick a page:",
      (p) => `${p.title}  (modified ${p.lastModifiedDateTime})`
    );
  }

  const safeName = chosen.title.replace(/[^a-z0-9_-]+/gi, "_").slice(0, 60);
  const runDir = path.join(OUT_DIR, `${safeName}_${chosen.id.slice(-8)}`);
  fs.mkdirSync(runDir, { recursive: true });

  console.log("\nFetching normal page content (HTML)...");
  const htmlResp = await fetchPageHtml(token, chosen.id);
  fs.writeFileSync(path.join(runDir, "page.html"), htmlResp.text);
  console.log(`  -> saved page.html (status ${htmlResp.status})`);

  console.log("Fetching page content with includeInkML=true...");
  const inkResp = await fetchPageInk(token, chosen.id);
  fs.writeFileSync(path.join(runDir, "ink-raw.txt"), inkResp.text);
  fs.writeFileSync(
    path.join(runDir, "ink-response-headers.json"),
    JSON.stringify({ status: inkResp.status, contentType: inkResp.contentType }, null, 2)
  );
  console.log(`  -> saved ink-raw.txt (status ${inkResp.status}, type: ${inkResp.contentType})`);

  // Search the raw body plus every multipart section for ink data.
  const blobs = [{ label: "raw-body", text: inkResp.text }];
  const parts = splitMultipart(inkResp.text, inkResp.contentType);
  if (parts) {
    console.log(`  -> multipart, split into ${parts.length} part(s)`);
    parts.forEach((part, i) => {
      const fname = `ink-part-${i}.txt`;
      fs.writeFileSync(path.join(runDir, fname), part.body);
      console.log(`     part ${i}: ${part.contentType} -> ${fname}`);
      blobs.push({ label: `part-${i}`, text: part.body });
    });
  }

  console.log("\nLooking for parsable ink/trace data...");
  let strokes = null;
  let foundIn = null;
  for (const blob of blobs) {
    const result = tryParseInkML(blob.text);
    if (result) {
      strokes = result;
      foundIn = blob.label;
      break;
    }
  }

  if (!strokes) {
    console.log(`\nNo parsable ink data found. Raw output saved in ${runDir}.`);
    return;
  }

  console.log(`Found ${strokes.length} stroke(s) in ${foundIn}. Generating SVG...`);
  const svg = strokesToSVG(strokes);
  fs.writeFileSync(path.join(runDir, "ink.svg"), svg);
  console.log(`  -> saved ink.svg (${strokes.length} strokes)`);

  const resvg = new Resvg(svg, { fitTo: { mode: "width", value: 1400 } });
  const png = resvg.render().asPng();
  fs.writeFileSync(path.join(runDir, "ink.png"), png);
  console.log(`  -> saved ink.png (${png.length} bytes)`);

  console.log(`\nAll output for this run: ${runDir}`);
}

main().catch((err) => {
  console.error("\nFailed:", err);
  process.exit(1);
});
```

### Azure app registration (one-time)

1. [entra.microsoft.com](https://entra.microsoft.com) → confirm the tenant you want is the active directory (account icon → Switch directory).
2. **Identity** → **Applications** → **App registrations** → **New registration**.
3. Supported account types: **Accounts in any organizational directory and personal Microsoft accounts**.
4. Redirect URI: leave blank — device code flow doesn't need one.
5. Copy the **Application (client) ID**.
6. **Manage** → **Authentication** → **Advanced settings** → set **Allow public client flows** to **Yes**. (Easy to miss; device code flow fails without it.)
7. **Manage** → **API permissions** → **Microsoft Graph** → **Delegated permissions** → add `Notes.Read` and `Notes.ReadWrite`.