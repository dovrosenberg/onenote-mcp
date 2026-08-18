// Device-code sign-in that seeds the Firestore token cache.
//
// This is the only interactive Microsoft sign-in in the project, and it runs on the
// operator's own machine. The deployed server acquires silently from what this writes
// (src/graph-auth.ts); it has no way to prompt anyone, and Graph's OneNote endpoints do
// not support app-only auth, so when the stored refresh token dies this CLI is the only
// recovery.
//
// Everything that decides what the server can later read is shared with the server
// rather than restated here: the client id and authority from src/config.ts, GRAPH_SCOPES
// from src/graph-auth.ts, and the cache plugin from src/token-cache.ts. MSAL keys cached
// tokens by client id and by scope string, so a second spelling of either would write a
// cache that looks valid and that the server's silent acquisition misses.

import { PublicClientApplication } from '@azure/msal-node';
import type { AccountInfo, DeviceCodeRequest } from '@azure/msal-node';

import { ConfigError, exitOnConfigError, loadConfig } from './config.ts';
import { GRAPH_SCOPES } from './graph-auth.ts';
import { createFirestoreTokenCachePlugin } from './token-cache.ts';

const GRAPH_ROOT = 'https://graph.microsoft.com/v1.0';

async function main(): Promise<void> {
  // `firestore-explicit`, not `firestore`: this process authenticates to Firestore with
  // the operator's Application Default Credentials, so an unset GOOGLE_CLOUD_PROJECT or
  // FIRESTORE_CACHE_DOC would seed a real document in the wrong place and report
  // success. The server keeps the defaulting group; Cloud Run infers its own project.
  const { graph, firestore } = loadConfig(['graph', 'firestore-explicit']);
  if (graph === undefined || firestore === undefined) {
    throw new Error('internal: loadConfig returned no graph or firestore section');
  }
  const { projectId } = firestore;
  if (projectId === undefined) {
    throw new Error('internal: GOOGLE_CLOUD_PROJECT missing after validation');
  }

  const client = new PublicClientApplication({
    auth: { clientId: graph.clientId, authority: graph.authority },
    cache: { cachePlugin: createFirestoreTokenCachePlugin(firestore) },
  });

  // The cache write happens inside this call, through the plugin's afterCacheAccess.
  // Nothing here serializes anything itself.
  const result = await client.acquireTokenByDeviceCode({
    scopes: [...GRAPH_SCOPES],
    deviceCodeCallback: printDeviceCode,
  });

  if (result === null || result.account === null) {
    throw new Error('Device-code sign-in returned no account, so nothing was cached.');
  }

  const notebookCount = await countNotebooks(result.accessToken);

  printConfirmation(result.account, projectId, firestore.cacheDocumentPath, notebookCount);
}

const printDeviceCode: DeviceCodeRequest['deviceCodeCallback'] = (response) => {
  // response.message is Microsoft's own wording and already contains the verification
  // URL and the code. Rewording it risks printing a stale URL.
  //
  // MSAL invokes this callback with whatever the device-code endpoint returned, without
  // first checking that the request succeeded, so on a rejected request every field is
  // undefined despite the declared type. A wrong client id shows up here as `undefined`
  // between the two banners, and the AuthError that follows a second later does not
  // mention the client id at all.
  const message =
    typeof response.message === 'string' && response.message !== ''
      ? response.message
      : 'The device-code endpoint returned no sign-in message. Check ONENOTE_CLIENT_ID, and that the app registration has "Allow public client flows" set to Yes.';

  process.stdout.write(`\n=== Sign in required ===\n${message}\n========================\n\n`);
};

/**
 * Prove the seeded token works by making one Graph call.
 *
 * `$select=id` keeps notebook display names out of the response body entirely, and only
 * the count is printed. Names are user content, and this output is read on a terminal
 * that ends up in screenshots and pasted into issues.
 *
 * `/me/onenote/notebooks` is deliberately the probe rather than `/me/onenote/pages`,
 * which fails with error 20266 once the account has enough sections.
 */
async function countNotebooks(accessToken: string): Promise<number> {
  const response = await fetch(`${GRAPH_ROOT}/me/onenote/notebooks?$select=id`, {
    headers: { Authorization: `Bearer ${accessToken}` },
  });

  if (!response.ok) {
    // The body is not included: a Graph error carries request ids and tenant-scoped
    // detail, and the status is what says whether the grant took.
    throw new Error(
      `The token was cached, but GET /me/onenote/notebooks returned ${response.status} ${response.statusText}. Check that the app registration has Notes.Read and Notes.ReadWrite granted.`,
    );
  }

  const body: unknown = await response.json();
  const value = (body as { value?: unknown }).value;
  if (!Array.isArray(value)) {
    throw new Error('GET /me/onenote/notebooks returned no "value" array.');
  }

  return value.length;
}

function printConfirmation(
  account: AccountInfo,
  projectId: string,
  documentPath: string,
  notebookCount: number,
): void {
  // The home tenant is printed because it is the one thing a successful sign-in can get
  // wrong without any error: the device-code page will happily sign you into a personal
  // account or a second work tenant. It is also the reason for the warning line — the
  // repository is public and the tenant id does not belong in an issue or a log.
  process.stdout.write(
    [
      '',
      'Token cache seeded. The deployed server can now acquire Graph tokens silently.',
      '',
      `  Firestore project:  ${projectId}`,
      `  Firestore document: ${documentPath}`,
      `  Home tenant:        ${account.tenantId}`,
      `  Notebooks visible:  ${notebookCount}`,
      '',
      'Do not paste this output into an issue, a pull request, or a workflow log: the',
      'home tenant id identifies the Entra directory.',
      '',
    ].join('\n'),
  );
}

try {
  await main();
} catch (err) {
  if (err instanceof ConfigError) exitOnConfigError(err);

  // One readable line, no stack — the same contract the server entrypoint has. A stack
  // here would be MSAL's or Firestore's internals, which say nothing about what to fix.
  process.stderr.write(`Bootstrap failed. ${describeError(err)}\n`);
  process.exit(1);
}

function describeError(err: unknown): string {
  if (err instanceof Error) return `${err.name}: ${err.message}`;
  return String(err);
}
