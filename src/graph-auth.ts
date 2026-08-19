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
//
// Not every failure means that, and telling them apart is the point of the
// `cache-unavailable` reason. Firestore is read and written inside `acquireTokenSilent`,
// through the plugin in ./token-cache.ts, so a backend outage used to arrive here as the
// same rejection a dead refresh token produces — and the message for that one sends a
// human to a browser for something a retry would have fixed. A
// `TokenCacheUnavailableError` anywhere in the cause chain is reported as retryable
// instead.
//
// `forceRefresh` is the other half of running unattended. Without it MSAL answers from
// its own in-memory access token and never touches the refresh token, so the refresh
// token's inactivity window only slides when the service is actually used. ./keepalive.ts
// calls `GraphAuth.refresh()` on a timer for that reason.

import { PublicClientApplication } from '@azure/msal-node';
import type { AccountInfo, AuthenticationResult, SilentFlowRequest } from '@azure/msal-node';

import type { FirestoreConfig, GraphConfig } from './config.ts';
import { logEvent } from './logging.ts';
import { createFirestoreTokenCachePlugin, TokenCacheUnavailableError } from './token-cache.ts';

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

export type GraphAuthErrorReason =
  | 'cache-unreadable'
  | 'cache-unavailable'
  | 'no-account'
  | 'silent-failed';

/** Reasons a retry can fix. Everything else needs a human at a browser. */
const RETRYABLE: ReadonlySet<GraphAuthErrorReason> = new Set(['cache-unavailable']);

export class GraphAuthError extends Error {
  readonly reason: GraphAuthErrorReason;
  readonly documentPath: string;
  /** True when waiting and calling again is the right response. */
  readonly retryable: boolean;

  constructor(
    reason: GraphAuthErrorReason,
    documentPath: string,
    options: { cause?: unknown } = {},
  ) {
    super(buildMessage(reason, documentPath, options.cause));
    this.name = 'GraphAuthError';
    this.reason = reason;
    this.documentPath = documentPath;
    this.retryable = RETRYABLE.has(reason);
    if ('cause' in options) this.cause = options.cause;
  }
}

/**
 * Build the error, log it as an operational event, and throw.
 *
 * The event line is what an alert policy keys on. A tool result carrying this message
 * only ever reaches the model in a Claude conversation, so without a line in the log
 * nothing tells the operator the connector has stopped working — and the longer a dead
 * grant runs, the more chances there are for the stored cache to decay further.
 */
function fail(
  reason: GraphAuthErrorReason,
  documentPath: string,
  options: { cause?: unknown } = {},
): never {
  const err = new GraphAuthError(reason, documentPath, options);
  // No cause detail and no account identifier: `documentPath` is configuration, and
  // `reason` is one of four fixed strings.
  logEvent('graph-auth-failure', {
    reason,
    documentPath,
    retryable: err.retryable ? 'true' : 'false',
  });
  throw err;
}

/** Is a TokenCacheUnavailableError anywhere in this error's cause chain? */
function causedByUnavailableCache(err: unknown): boolean {
  const seen = new Set<unknown>();
  let current: unknown = err;

  while (current !== undefined && current !== null && !seen.has(current)) {
    if (current instanceof TokenCacheUnavailableError) return true;
    seen.add(current);
    current = current instanceof Error ? current.cause : undefined;
  }

  return false;
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

/** Options for one acquisition. */
export interface AcquireOptions {
  /**
   * Skip MSAL's cached access token and exchange the refresh token.
   *
   * This is what makes a keepalive call worth anything: without it MSAL returns the
   * access token it already holds, no request reaches Entra, and the refresh token's
   * inactivity window does not move. Leave it off everywhere else — every forced refresh
   * is a token-endpoint round trip and a Firestore write.
   */
  readonly forceRefresh?: boolean;
}

/**
 * Acquire a Graph access token from the seeded cache, or throw a GraphAuthError.
 *
 * `documentPath` is used only to name the cache in the error messages.
 *
 * @throws {GraphAuthError} on an unreadable cache, an unreachable cache backend, an
 * empty cache, or a failed silent acquisition. Nothing else escapes: a raw MSAL error
 * would not tell the operator that a human has to re-run the bootstrap CLI — nor, for
 * `cache-unavailable`, that they do not.
 */
export async function acquireGraphToken(
  client: SilentTokenSource,
  documentPath: string,
  options: AcquireOptions = {},
): Promise<string> {
  let accounts: AccountInfo[];
  try {
    // This call is what triggers beforeCacheAccess, so a TokenCacheError from the
    // decoder, a TokenCacheUnavailableError from the backend, and a SyntaxError from
    // MSAL's deserialize all surface here.
    accounts = await client.getTokenCache().getAllAccounts();
  } catch (err) {
    fail(causedByUnavailableCache(err) ? 'cache-unavailable' : 'cache-unreadable', documentPath, {
      cause: err,
    });
  }

  // The bootstrap CLI seeds exactly one account. If a second ever appears, use the
  // first rather than guessing which one the operator meant.
  const [account] = accounts;
  if (account === undefined) {
    fail('no-account', documentPath);
  }

  let result: AuthenticationResult | null;
  try {
    result = await client.acquireTokenSilent({
      account,
      scopes: [...GRAPH_SCOPES],
      // Spread rather than passed as possibly-undefined: exactOptionalPropertyTypes
      // treats an explicit undefined as a different type from an absent property.
      ...(options.forceRefresh === true ? { forceRefresh: true } : {}),
    });
  } catch (err) {
    // The write half of the refresh happens inside this call too, so a Firestore outage
    // arrives here rather than as its own rejection. Reporting it as `silent-failed`
    // would tell the operator their refresh token is dead when it is not.
    fail(causedByUnavailableCache(err) ? 'cache-unavailable' : 'silent-failed', documentPath, {
      cause: err,
    });
  }

  if (result === null || result.accessToken === '') {
    fail('silent-failed', documentPath);
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

  /**
   * Exchange the refresh token even when the held access token is still good.
   *
   * Called by the keepalive route and by nothing else. Its value is the side effect
   * rather than the return: Entra issues a replacement refresh token with a fresh
   * inactivity window, and ./token-cache.ts writes it to Firestore. That is what keeps
   * an unused connector from needing `npm run bootstrap` after ~90 idle days.
   */
  refresh(): Promise<string> {
    return acquireGraphToken(this.#client, this.#documentPath, { forceRefresh: true });
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
    case 'cache-unavailable':
      // Returns early: this is the one reason that must not end in "run the bootstrap
      // CLI". Nothing is wrong with the stored credential and an interactive sign-in
      // would replace a working refresh token to fix a Firestore outage.
      lines.push(
        `The Firestore document ${documentPath} holding the Microsoft token cache could not be reached, so no token could be acquired. The stored credential is unaffected and no sign-in is needed. Check that Firestore is reachable and that the runtime service account still holds roles/datastore.user, then retry.`,
      );
      return appendCause(lines, cause);
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

  return appendCause(lines, cause);
}

function appendCause(lines: string[], cause: unknown): string {
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
