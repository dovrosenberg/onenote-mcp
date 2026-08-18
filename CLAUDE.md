# CLAUDE.md

Working notes for agents in this repository. `project-spec.md` is the authoritative
design document — read the relevant section of it before changing behaviour.

## Directory structure

```
onenote-mcp/
├── CLAUDE.md                  # this file
├── README.md                  # human-facing: quick start, scripts, config table
├── project-spec.md            # authoritative design doc — ink pipeline, OAuth layers, deploy model
├── LICENSE
│
├── package.json               # ESM, Node >= 24
├── package-lock.json          # committed; `npm ci` depends on it
├── tsconfig.json              # editor + `npm run typecheck`: src/ AND test/, noEmit
├── tsconfig.build.json        # `npm run build`: src/ only, emits to dist/
│
├── Dockerfile                 # multi-stage; runtime is node:24-slim with prod deps only (issue #6)
├── .dockerignore              # denylist for the build context; keeps host node_modules and dist out
│
├── src/                       # TypeScript source, the only thing that ships
│   ├── index.ts               # server entrypoint: validate config, bind $PORT, listen
│   ├── server.ts              # builds the Express app; does not listen
│   ├── config.ts              # env-var schema, grouped validation, ConfigError
│   ├── bootstrap.ts           # device-code CLI that seeds the Firestore token cache (placeholder, issue #9)
│   └── version.ts             # SERVICE_NAME, VERSION
│
├── test/                      # node --test suites; test/<name>.test.ts covers src/<name>.ts
│   ├── config.test.ts
│   ├── server.test.ts
│   └── version.test.ts
│
├── docs/
│   └── plans/                 # one implementation plan per issue: <issue-number>-<slug>.md
│
├── scripts/                   # operational shell scripts, not part of the built artifact
│   ├── gcp-bootstrap.sh       # one-time GCP provisioning (issue #2)
│   └── test/                  # bash tests, run by scripts/test/run.sh
│       ├── run.sh             # syntax + shellcheck + the behavioural suites
│       ├── bootstrap.test.sh  # gcp-bootstrap.sh against a stub `gcloud` on PATH
│       ├── docker.test.sh     # builds the image and runs the container (opt-in)
│       └── stubs/gcloud
│
└── dist/                      # build output; gitignored, never edited, never committed
```

A test file is named for the source file it covers: `test/config.test.ts` covers
`src/config.ts`. Not every source file has one. `src/index.ts` is a top-level entrypoint
whose behaviour is process exit codes and stderr, so it is verified by running it rather
than by a unit test — the commands are in the Final Verification section of
`docs/plans/5-repo-skeleton.md`. `src/bootstrap.ts` is a placeholder until issue #9.

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
commit. It catches the class of error a linter would here, and it is the reason the
plans list it under "lint".

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
`oauth`, `server`. The server asks for all four; `bootstrap.ts` asks for `graph` and
`firestore` only, so the operator running it locally is never made to hold
`MCP_OAUTH_CLIENT_SECRET`. Add a new variable to the `SPECS` table, not to a new
`process.env` read somewhere else.

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

## Repository hygiene

This repository is public, and so are its issues, pull requests, and Actions logs.

Never commit, and never put in an issue or a workflow log: real page content, rendered
ink images, the Entra tenant name or ID, Firestore document contents, or the Layer-1
OAuth client secret. `.gitignore` covers `output/`, `.env*`, `*.token-cache.json`, and
`.msal-cache*`. Example output in docs must come from a throwaway page with fake content.
The Azure client ID is *not* secret in this design (public client, device-code flow) and
is safe in config examples.

## Workflow

Issues are labelled by phase (`phase-1-skeleton` … `phase-5-deploy`) and by who does them
(`claude` or `manual`). Each `claude` issue gets a plan in `docs/plans/` before code,
following the structure of the existing plans: numbered acceptance criteria traced to
phases, and a Findings section recording what implementation revealed that planning
missed. Commit the plan, then implement phase by phase.
