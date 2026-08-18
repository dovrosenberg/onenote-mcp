import express, { type Application, type Request, type Response } from 'express';

import type { Config } from './config.ts';
import { SERVICE_NAME, VERSION } from './version.ts';

/**
 * Build the Express application without binding a port.
 *
 * Construction is separate from listening so tests can exercise routes on an ephemeral
 * port, and so the MCP transport in issue #14 has a seam to mount onto.
 */
export function createApp(_config: Config): Application {
  const app = express();

  // Cloud Run terminates TLS and forwards the original scheme and client IP in
  // X-Forwarded-*. Issue #22's authorization-code redirect has to build absolute https
  // URLs, which req.protocol only reports correctly with this set.
  app.set('trust proxy', true);
  app.disable('x-powered-by');

  // Deliberately reports no configuration. The service deploys with
  // --allow-unauthenticated, so anything this returns is public.
  app.get('/healthz', (_req: Request, res: Response) => {
    res.status(200).json({ status: 'ok', service: SERVICE_NAME, version: VERSION });
  });

  return app;
}
