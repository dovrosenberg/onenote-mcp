# Server conventions

Read before changing `src/server.ts`, `src/mcp-server.ts`, `src/mcp-tools.ts`, the OAuth
modules, the keepalive route, or `src/logging.ts`.

## MCP transport and the tool registry

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

## OAuth, the bearer gate, and error mapping

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

## Scheduler routes and event logging

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

**Two events exist for alerting, and their field vocabulary is fixed.**
`graph-auth-failure` and `token-cache-write-refused`, written through `logEvent` in
`src/logging.ts` to stderr. They carry a reason from a fixed set and a document path from
configuration — never an account identifier, never a cause message, never user content.
They are separate from the request line because a log-based metric keyed on a status code
cannot tell a dead refresh token from a missing page. `setEventSink` exists for tests only.
