// Layer 2 of the two OAuth layers: this server to Microsoft Graph.
//
// The deployed service never signs in interactively. It acquires silently from the
// Firestore cache that the bootstrap CLI seeded, and MSAL exchanges the stored refresh
// token when the access token expires. Refresh tokens rotate on every use, so the write
// half of that exchange goes back to Firestore through the cache plugin in
// ./token-cache.ts.
//
// When the refresh token is finally dead — revoked, or expired because the service sat
// idle past the ~90-day window — there is no automatic recovery. Graph's OneNote
// endpoints do not support app-only auth, so a human has to re-run the bootstrap CLI.
// That is why every failure here is a GraphAuthError whose message says so, rather than
// a raw MSAL error that reaches the caller as a bare 401 from Graph.

import { PublicClientApplication } from '@azure/msal-node';
import type { AccountInfo, AuthenticationResult, SilentFlowRequest } from '@azure/msal-node';

import type { FirestoreConfig, GraphConfig } from './config.ts';
import { createFirestoreTokenCachePlugin } from './token-cache.ts';

/**
 * The scopes this server asks for, fully qualified.
 *
 * MSAL keys cached access tokens by scope string, so the short form `Notes.Read` here
 * and the fully-qualified form in the bootstrap CLI would look like two different
 * grants and the silent lookup would miss. The bootstrap CLI imports this same
 * constant rather than repeating the strings.
 */
export const GRAPH_SCOPES: readonly string[] = [
  'https://graph.microsoft.com/Notes.Read',
  'https://graph.microsoft.com/Notes.ReadWrite',
];

const BOOTSTRAP_COMMAND = 'npm run bootstrap';

export type GraphAuthErrorReason = 'cache-unreadable' | 'no-account' | 'silent-failed';

export class GraphAuthError extends Error {
  readonly reason: GraphAuthErrorReason;
  readonly documentPath: string;

  constructor(
    reason: GraphAuthErrorReason,
    documentPath: string,
    options: { cause?: unknown } = {},
  ) {
    super(buildMessage(reason, documentPath, options.cause));
    this.name = 'GraphAuthError';
    this.reason = reason;
    this.documentPath = documentPath;
    if ('cause' in options) this.cause = options.cause;
  }
}

/**
 * The slice of `PublicClientApplication` this module calls. Declared narrowly so the
 * acquisition logic is callable from a test with a plain object; the real client
 * satisfies it structurally, so `createGraphAuth` needs no cast.
 *
 * `acquireTokenSilent` is widened to allow null. The real client's signature does not,
 * but the null branch is the one that would otherwise send an unusable token to Graph.
 */
export interface SilentTokenSource {
  getTokenCache(): { getAllAccounts(): Promise<AccountInfo[]> };
  acquireTokenSilent(request: SilentFlowRequest): Promise<AuthenticationResult | null>;
}

/**
 * Acquire a Graph access token from the seeded cache, or throw a GraphAuthError.
 *
 * `documentPath` is used only to name the cache in the error messages.
 *
 * @throws {GraphAuthError} on an unreadable cache, an empty cache, or a failed silent
 * acquisition. Nothing else escapes: a raw MSAL error would not tell the operator that
 * a human has to re-run the bootstrap CLI.
 */
export async function acquireGraphToken(
  client: SilentTokenSource,
  documentPath: string,
): Promise<string> {
  let accounts: AccountInfo[];
  try {
    // This call is what triggers beforeCacheAccess, so a TokenCacheError from the
    // decoder and a SyntaxError from MSAL's deserialize both surface here.
    accounts = await client.getTokenCache().getAllAccounts();
  } catch (err) {
    throw new GraphAuthError('cache-unreadable', documentPath, { cause: err });
  }

  // The bootstrap CLI seeds exactly one account. If a second ever appears, use the
  // first rather than guessing which one the operator meant.
  const [account] = accounts;
  if (account === undefined) {
    throw new GraphAuthError('no-account', documentPath);
  }

  let result: AuthenticationResult | null;
  try {
    result = await client.acquireTokenSilent({ account, scopes: [...GRAPH_SCOPES] });
  } catch (err) {
    throw new GraphAuthError('silent-failed', documentPath, { cause: err });
  }

  if (result === null || result.accessToken === '') {
    throw new GraphAuthError('silent-failed', documentPath);
  }

  return result.accessToken;
}

/**
 * Holds the MSAL client for the life of the process.
 *
 * That retained client is the in-process token cache: MSAL serves a still-valid access
 * token from its own in-memory store without a network call. This class deliberately
 * stores no token and no expiry of its own.
 */
export class GraphAuth {
  readonly #client: SilentTokenSource;
  readonly #documentPath: string;

  constructor(client: SilentTokenSource, documentPath: string) {
    this.#client = client;
    this.#documentPath = documentPath;
  }

  getAccessToken(): Promise<string> {
    return acquireGraphToken(this.#client, this.#documentPath);
  }
}

/**
 * Build the real client: the public-client app registration, the tenant authority, and
 * the Firestore-backed cache plugin.
 *
 * This is the only `PublicClientApplication` construction in the server path. It is
 * built once per process; constructing one per request would throw away MSAL's
 * in-memory cache and hit the token endpoint every time.
 */
export function createGraphAuth(graph: GraphConfig, firestore: FirestoreConfig): GraphAuth {
  const client = new PublicClientApplication({
    auth: { clientId: graph.clientId, authority: graph.authority },
    cache: { cachePlugin: createFirestoreTokenCachePlugin(firestore) },
  });

  return new GraphAuth(client, firestore.cacheDocumentPath);
}

function buildMessage(
  reason: GraphAuthErrorReason,
  documentPath: string,
  cause: unknown,
): string {
  // No account identifier appears in any of these. `username` is the user's UPN and
  // `homeAccountId` embeds the tenant id; both are things the repository hygiene rules
  // keep out of anything that can reach an issue or an Actions log.
  const lines: string[] = [];

  switch (reason) {
    case 'cache-unreadable':
      lines.push(
        `The Microsoft token cache at ${documentPath} could not be read: it is absent, or its contents are not a cache MSAL recognises.`,
      );
      break;
    case 'no-account':
      lines.push(
        `The Microsoft token cache at ${documentPath} was read but holds no signed-in account, so there is no refresh token to exchange.`,
      );
      break;
    case 'silent-failed':
      lines.push(
        `Silent token acquisition failed against the Microsoft token cache at ${documentPath}. The stored refresh token is expired or revoked.`,
      );
      break;
  }

  lines.push(
    `This server never signs in interactively, so it cannot recover on its own. Run \`${BOOTSTRAP_COMMAND}\` on a machine with a browser to sign in again and rewrite the cache.`,
  );

  const detail = describeCause(cause);
  if (detail !== null) lines.push(`Underlying error: ${detail}`);

  return lines.join(' ');
}

/** Cause detail without the stack, and without assuming an Error was thrown. */
function describeCause(cause: unknown): string | null {
  if (cause === undefined || cause === null) return null;
  if (cause instanceof Error) return `${cause.name}: ${cause.message}`;
  return String(cause);
}
