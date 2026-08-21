# CLAUDE.md

Working notes for agents in this repository. `project-spec.md` is the authoritative
design document — read the relevant section of it before changing behaviour.
`api-overview.md` is what Microsoft Graph's OneNote endpoints actually do: the resource
paths, the query options and their limits, and the places the service contradicts its own
documentation. Read it before adding or changing a Graph call, and add anything new that
a live run reveals.

The conventions are split by area. Each file states what it covers; read the one whose
area you are changing, because most of its rules exist to stop a specific silent failure.

| File | Read it before changing |
|---|---|
| `docs/graph.md` | any Microsoft Graph call, the request budget, throttling, or the token cache |
| `docs/mirror.md` | anything under `src/mirror-*.ts`, `src/read-sync.ts`, or the sync route |
| `docs/content.md` | ink rendering, the HTML trimmer, page reads or writes, the name-lookup ladder |
| `docs/server.md` | `src/server.ts`, the MCP transport, OAuth, the keepalive route, `src/logging.ts` |
| `docs/config-deploy.md` | a configuration variable, startup behaviour, the Dockerfile, `deploy.yml` |
| `test/README.md` | any test file — what each suite asserts and what it deliberately does not |

## Directory structure

```
onenote-mcp/
├── src/                       # TypeScript source, the only thing that ships│
├── test/                      # node --test suites; test/<name>.test.ts covers src/<name>.ts
│   ├── README.md              # what each suite covers, and what it deliberately does not
│   ├── fixtures/              # hand-authored InkML and HTML; never a captured page dump
├── docs/                      # area conventions; see the table above
├── scripts/                   # operational shell scripts, not part of the built artifact
│   └── test/                  # bash tests, run by scripts/test/run.sh
├── .github/workflows/         # deploy.yml, the only CI; builds, tests, pushes, deploys
└── dist/                      # build output; gitignored, never edited, never committed
```

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
commit. It catches the class of error a linter would here.

## Conventions

These apply to every file. The rest are in the area files listed above.

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

**Nothing that matters is in memory, and that is a design position rather than an
accident.** An instance can disappear between any two requests. What is held in memory is
MSAL's access token (rebuilt from the Firestore cache, one token round trip), the memoised
Firestore and Storage clients (rebuilt lazily, no connection at construction), the request
gate's pacing state (a fresh instance may open its 4 concurrent immediately rather than
spacing from the last request), and the pending OAuth codes (single-use, 60 seconds, cap
100 — losing one costs a retry of the consent click, which is the one moment a human is
already present). Everything with a real cost to lose is elsewhere on purpose: access and
refresh tokens are signed rather than stored, the consent form rides in a signed hidden
field, the Microsoft refresh token is in Firestore, and the whole mirror is in Firestore
and GCS. Do not add a cache that a request depends on having.

## Hard rules

A violation of any of these is invisible in review and expensive in production. The file
named after each one carries the reasoning.

- **Never call the account-wide page list.** `GET /me/onenote/pages` fails with error
  20266 once the account has enough sections across all notebooks, and this one does.
  Page listing is always scoped to `/me/onenote/sections/{id}/pages`; a test in
  `test/graph-structure.test.ts` scans `src/` for the path. — `docs/graph.md`
- **Prefer `getExpandedTree` to `getFullTree`.** One request against 195, against an
  hourly budget of 400. — `docs/graph.md`
- **Never fan out with an unbounded `Promise.all` over ids.** The concurrency limit is 5
  and the safe cap is 4. — `docs/graph.md`
- **Do not add `forceRefresh` to the tool path.** It exists for the keepalive route, and
  every use is a token-endpoint round trip and a Firestore write. — `docs/graph.md`
- **The server never signs in interactively.** `acquireTokenByDeviceCode` belongs to
  `src/bootstrap.ts` alone; a test scans `src/` for it. — `docs/graph.md`
- **Do not parse the MSAL cache blob**, and do not make `isEmptyCache` look for key names.
  That change silently inverts the write guard. — `docs/graph.md`
- **Every write is three steps in this order: invalidate the mirror, write to OneNote,
  resync the mirror.** Any other order leaves superseded content served as current when
  the process stops mid-write. — `docs/mirror.md`
- **A listing difference triggers a re-fetch, never a stale mark.** A mark deletes the
  page-content document and nothing re-fetches a stale page. — `docs/mirror.md`
- **Every mirror read answers `null` on a miss, and every miss means "ask Graph".** Never
  fail a tool call because the mirror is unavailable. — `docs/mirror.md`
- **Do not add a `useLiveData` argument back**, and do not add an activity check to the
  write path. Both were removed deliberately. — `docs/mirror.md`
- **Never hardcode a port.** Cloud Run sets `PORT`. — `docs/config-deploy.md`
- **Startup failures print one readable message and exit 1** — never a stack trace.
  — `docs/config-deploy.md`
- **The container's runtime base is Debian `-slim`, never Alpine.** The resvg package
  ships no musl prebuild, and handwriting rendering is the point of the service.
  — `docs/config-deploy.md`
- **Nothing that reaches a log or a tool result may quote user content.** Not a notebook,
  section or page name, not a response body, not a query string, not stroke coordinates.
  — `docs/server.md`, `docs/graph.md`

## Repository hygiene

This repository is public, and so are its issues, pull requests, and Actions logs.

Never commit, and never put in an issue or a workflow log: real page content, rendered
ink images, the Entra tenant name or ID, Firestore document contents, or the Layer-1
OAuth client secret. `.gitignore` covers `output/`, `.env*`, `*.token-cache.json`, and
`.msal-cache*`. Example output in `README.md` must come from a throwaway page with fake content.
The Azure client ID is *not* secret in this design (public client, device-code flow) and
is safe in config examples.

## Workflow

Issues are labelled by phase (`phase-1-skeleton` … `phase-5-deploy`) and by who does them
(`claude` or `manual`). Work a `claude` issue directly; no separate plan document is
written. The issue's own task list and acceptance section are the specification, and the
commit message is where the reasoning goes — what the
implementation revealed that the issue did not anticipate belongs there, not in a
separate document.
