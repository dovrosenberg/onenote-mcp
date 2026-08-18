import { exitOnConfigError, loadConfig, type Config } from './config.ts';
import { createApp } from './server.ts';

let config: Config;
try {
  config = loadConfig(['graph', 'firestore', 'oauth', 'server']);
} catch (err) {
  exitOnConfigError(err);
}

const port = config.server?.port;
if (port === undefined) {
  throw new Error("internal: 'server' config group requested but absent");
}

// app.listen returns the http.Server. The 'error' handler is attached after the call
// because a bind failure is emitted asynchronously, on the next tick at the earliest.
const server = createApp(config).listen(port, () => {
  console.log(`listening on port ${port}`);
});

// A bind failure is a startup misconfiguration like a missing variable, so it gets the
// same treatment: a readable line, exit 1, no stack trace.
server.on('error', (err: NodeJS.ErrnoException) => {
  if (err.code === 'EADDRINUSE') {
    process.stderr.write(`Port ${port} is already in use. Set PORT to a free port.\n`);
    process.exit(1);
  }
  if (err.code === 'EACCES') {
    process.stderr.write(`Not permitted to bind port ${port}. Set PORT above 1023.\n`);
    process.exit(1);
  }
  throw err;
});

// Cloud Run sends SIGTERM before terminating an instance. Without this the container is
// killed after the grace period on every revision change.
for (const signal of ['SIGTERM', 'SIGINT'] as const) {
  process.on(signal, () => {
    console.log(`${signal} received, shutting down`);
    server.close(() => process.exit(0));
  });
}
