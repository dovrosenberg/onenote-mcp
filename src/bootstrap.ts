// Placeholder for the device-code bootstrap CLI. Issue #9 implements it.
//
// It exists now so `npm run bootstrap` resolves to a real file, and so the grouped
// config loader has a second consumer proving that the MCP_OAUTH_* variables really are
// not required here.

import { exitOnConfigError, loadConfig } from './config.ts';

try {
  loadConfig(['graph', 'firestore']);
} catch (err) {
  exitOnConfigError(err);
}

process.stderr.write('bootstrap CLI not implemented yet — see issue #9\n');
process.exit(1);
