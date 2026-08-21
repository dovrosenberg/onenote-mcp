# Configuration, process wiring, and deployment

Read before adding a configuration variable, changing startup behaviour, touching the
Dockerfile, or editing `.github/workflows/deploy.yml`.

## Configuration and startup

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
data migration, and the tool modules are identical under both settings. That includes the
refresh: `createReadSyncFor` answers undefined when reads are off, so no tool call spends
a Graph request on a sync whose result it has no use for, and every answer reports
`source: "onenote"` because Graph is the only thing it read. The server asks
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

## Process-wide clients and dependencies

**One `GraphAuth` per process, built by `createGraphAuthFor` and passed to both
consumers.** `createTools` takes it as an argument rather than building it. Two MSAL
clients would each hold their own in-memory access token and their own view of the
Firestore cache, so a forced refresh through one would leave the other on a superseded
blob until its next read, and both would be writing the same document.

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

## The deploy workflow

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
