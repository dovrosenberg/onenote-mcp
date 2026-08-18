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
| `npm test` | `node --test` over `src/**/*.test.ts` |

Tests run directly against TypeScript source using Node's native type stripping, so
`npm test` does not need a build. That constrains the source: no `enum`, no `namespace`,
no constructor parameter properties, and type-only imports must be written `import type`.
The `erasableSyntaxOnly` and `verbatimModuleSyntax` compiler options enforce this.

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
| `FIRESTORE_CACHE_DOC` | no | `tokencache/msal` | Firestore document path holding the MSAL token cache |
| `GOOGLE_CLOUD_PROJECT` | no | — | GCP project; inferred automatically on Cloud Run, needed locally |
| `PORT` | no | `8080` | Bind port. Cloud Run sets this; the server never hardcodes one. |

`npm run bootstrap` needs only `ONENOTE_CLIENT_ID`, `ONENOTE_AUTHORITY`,
`FIRESTORE_CACHE_DOC`, and `GOOGLE_CLOUD_PROJECT` — not the `MCP_OAUTH_*` values.

## Repository hygiene

This repository is public. Do not commit real page content, rendered ink, Entra tenant
names or IDs, or Firestore document contents. `.gitignore` excludes `output/` and the
token-cache file patterns; see the "Repo hygiene" section of `project-spec.md`.
