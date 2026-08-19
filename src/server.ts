import express, { type Application, type Request, type Response } from 'express';
import { requireBearerAuth } from '@modelcontextprotocol/sdk/server/auth/middleware/bearerAuth.js';

import type { Config } from './config.ts';
import { KEEPALIVE_PATH, keepaliveRouter } from './keepalive.ts';
import { requestLogger } from './logging.ts';
import { MCP_PATH, mcpRouter } from './mcp-server.ts';
import { createOAuthProvider } from './oauth-provider.ts';
import { oauthRouter, protectedResourceMetadataUrl } from './oauth-router.ts';
import { createGraphAuthFor, createTools } from './tools.ts';
import { SERVICE_NAME, VERSION } from './version.ts';

/**
 * Every path the health endpoint answers on.
 *
 * Exported because the fail-closed route test in test/server.test.ts enumerates what may
 * answer without a bearer token, and a second health path added here has to show up
 * there rather than be remembered.
 */
export const HEALTH_PATHS = ['/healthz', '/health'] as const;

/**
 * Build the Express application without binding a port.
 *
 * Construction is separate from listening so tests can exercise routes on an ephemeral
 * port.
 */
export function createApp(config: Config): Application {
  const app = express();

  // Cloud Run terminates TLS and forwards the original scheme and client IP in
  // X-Forwarded-*. Nothing here reads req.protocol — every absolute URL this service
  // publishes is built from MCP_PUBLIC_URL — but express-rate-limit and Express's own
  // req.ip both behave differently without it, and the OAuth mount's rate-limit keying
  // is written against this being true.
  app.set('trust proxy', true);
  app.disable('x-powered-by');

  // First, so every route below is logged including the ones that fail. What the line
  // may contain is fixed in ./logging.ts.
  app.use(requestLogger());

  // Deliberately reports no configuration. The service deploys with
  // --allow-unauthenticated, so anything this returns is public.
  //
  // Two paths, same handler. `/healthz` is the conventional name and is what Cloud Run's
  // own probes use, which reach the container directly. It is not reachable from
  // outside: measured 2026-08-19 against the deployed service, Google's frontend answers
  // `https://<service>.run.app/healthz` with its own 404 page and the request never
  // appears in the Cloud Run request log, while `/health`, `/healthz2` and `/Healthz`
  // all arrive. So `/health` is the one an external check can call.
  for (const path of HEALTH_PATHS) {
    app.get(path, (_req: Request, res: Response) => {
      res.status(200).json({ status: 'ok', service: SERVICE_NAME, version: VERSION });
    });
  }

  // Layer-1 OAuth discovery, /authorize and /token. Mounted at the root because
  // mcpAuthRouter builds its paths from the issuer URL rather than from a mount point,
  // so behind a prefix it would advertise URLs it is not listening on. Every route it
  // serves stays open: discovery happens before a client holds a token, and the
  // authorization flow is how it gets one.
  if (config.oauth === undefined) {
    throw new Error("internal: 'oauth' config group required by createApp");
  }
  const provider = createOAuthProvider(config.oauth);
  app.use(oauthRouter(config.oauth, provider));

  // One Graph auth object for the whole process, shared by the tools and by the
  // keepalive route below. See `createGraphAuthFor` for why it is not built twice.
  const auth = createGraphAuthFor(config);

  // The keepalive route, when a secret is configured for it. It sits outside the bearer
  // gate because the caller is a scheduler, which cannot run an OAuth flow: it presents
  // a shared secret instead, and ./keepalive.ts compares it in constant time before
  // doing any work. Absent configuration, the route is not mounted and the path 404s.
  const keepaliveSecret = config.server?.keepaliveSecret;
  if (keepaliveSecret !== undefined) {
    app.use(KEEPALIVE_PATH, keepaliveRouter(keepaliveSecret, auth));
  }

  // Everything below this line needs a bearer token. The middleware reads the
  // Authorization header and nothing else — a token in `?access_token=` is not seen, so
  // it answers 401 — and it refuses a token whose audience is not this server's MCP
  // endpoint, which the spike in #20 measured the SDK as otherwise accepting.
  //
  // `resourceMetadataUrl` is what puts `resource_metadata="…"` in the WWW-Authenticate
  // header of the 401. That URL is how Claude finds the authorization server and starts
  // the flow, so a 401 without it is a dead end rather than a sign-in prompt.
  //
  // No `requiredScopes`. Passing any would make the middleware enforce them and answer a
  // 403 for a token that has none; the one scope this server issues, `offline_access`,
  // is about refresh tokens rather than about what a caller may do. If a scope check is
  // ever added, the 403 has to carry `WWW-Authenticate: Bearer
  // error="insufficient_scope"` — which this middleware does — because Claude treats any
  // other 403 as terminal and prompts for nothing.
  //
  // The tools are built once and shared by every request. The MCP server around them is
  // per-request — see ./mcp-server.ts.
  app.use(
    MCP_PATH,
    requireBearerAuth({
      verifier: provider,
      resourceMetadataUrl: protectedResourceMetadataUrl(config.oauth),
    }),
    mcpRouter(createTools(auth)),
  );

  return app;
}
