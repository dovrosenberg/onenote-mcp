import express, { type Application, type Request, type Response } from 'express';

import type { Config } from './config.ts';
import { requestLogger } from './logging.ts';
import { MCP_PATH, mcpRouter } from './mcp-server.ts';
import { createTools } from './mcp-tools.ts';
import { SERVICE_NAME, VERSION } from './version.ts';

/**
 * Build the Express application without binding a port.
 *
 * Construction is separate from listening so tests can exercise routes on an ephemeral
 * port.
 */
export function createApp(config: Config): Application {
  const app = express();

  // Cloud Run terminates TLS and forwards the original scheme and client IP in
  // X-Forwarded-*. Issue #22's authorization-code redirect has to build absolute https
  // URLs, which req.protocol only reports correctly with this set.
  app.set('trust proxy', true);
  app.disable('x-powered-by');

  // First, so every route below is logged including the ones that fail. What the line
  // may contain is fixed in ./logging.ts.
  app.use(requestLogger());

  // Deliberately reports no configuration. The service deploys with
  // --allow-unauthenticated, so anything this returns is public.
  app.get('/healthz', (_req: Request, res: Response) => {
    res.status(200).json({ status: 'ok', service: SERVICE_NAME, version: VERSION });
  });

  // The tools are built once and shared by every request. The MCP server around them is
  // per-request — see ./mcp-server.ts. Issue #23's bearer-token middleware goes between
  // this path and the router; /healthz stays open.
  app.use(MCP_PATH, mcpRouter(createTools(config)));

  return app;
}
