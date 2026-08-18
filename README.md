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
server against the same document to exercise it. Nothing wires the module into
`createApp` yet; the first consumer is the Graph structure client, issue #11.

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

`npm run bootstrap` reads only `ONENOTE_CLIENT_ID`, `ONENOTE_AUTHORITY`,
`FIRESTORE_CACHE_DOC`, and `GOOGLE_CLOUD_PROJECT` — not the `MCP_OAUTH_*` values — and it
requires the last two rather than defaulting them. See [Bootstrap](#bootstrap).

## Repository hygiene

This repository is public. Do not commit real page content, rendered ink, Entra tenant
names or IDs, or Firestore document contents. `.gitignore` excludes `output/` and the
token-cache file patterns; see the "Repo hygiene" section of `project-spec.md`.
