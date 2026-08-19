// Layer-1 OAuth: the metadata documents Claude discovers this server with, and the
// mount that serves /authorize and /token alongside them.
//
// Layer 1 is Claude → this server. It is unrelated to Layer 2, this server → Microsoft
// Graph, which lives in ./graph-auth.ts and uses a different client id, a different
// authority and a different flow. Nothing is shared between them.
//
// The protocol is the SDK's, not ours. `mcpAuthRouter` from @modelcontextprotocol/sdk
// carries both metadata documents, PKCE verification, redirect-URI matching, client
// authentication and the OAuth error shapes; the spike in issue #20 measured all of it
// against SDK 1.30.0 and the results are in project-spec.md under "Layer 1: what the SDK
// provides". What is left is an `OAuthServerProvider` — the consent screen, the token
// format and the code store — which is issue #22. Until that lands, `unimplemented` below
// stands in so the routes the metadata advertises exist.
//
// Three things about the mount are not free choices:
//
// - It goes at the application root. `mcpAuthRouter` builds every path from the issuer
//   URL rather than from a mount point, so behind a prefix it would advertise URLs that
//   are not where it is listening.
// - `scopesSupported` lists `offline_access`. That is the switch Claude reads to decide
//   whether to ask for a refresh token; without it the operator re-consents every time an
//   access token expires.
// - The clients store has no `registerClient`. Omitting it is what keeps
//   `registration_endpoint` out of the metadata, which is what keeps Dynamic Client
//   Registration out of the picture — the client id and secret are configured, and DCR
//   for a single-user server is machinery with no user.
//
// Protected-resource metadata is served only at the path-suffixed URL,
// `/.well-known/oauth-protected-resource/mcp`; the bare path answers 404. Measured, not
// assumed. Claude probes the suffixed path first and issue #23's 401 names it explicitly,
// so nothing here depends on the bare one.

import { Router, type Response } from 'express';
import { metadataHandler } from '@modelcontextprotocol/sdk/server/auth/handlers/metadata.js';
import {
  createOAuthMetadata,
  mcpAuthRouter,
} from '@modelcontextprotocol/sdk/server/auth/router.js';
import type { OAuthRegisteredClientsStore } from '@modelcontextprotocol/sdk/server/auth/clients.js';
import { ServerError } from '@modelcontextprotocol/sdk/server/auth/errors.js';
import type {
  AuthorizationParams,
  OAuthServerProvider,
} from '@modelcontextprotocol/sdk/server/auth/provider.js';
import type { AuthInfo } from '@modelcontextprotocol/sdk/server/auth/types.js';
import type { OAuthClientInformationFull, OAuthTokens } from '@modelcontextprotocol/sdk/shared/auth.js';

import type { OAuthConfig } from './config.ts';
import { MCP_PATH } from './mcp-server.ts';

/** Where the authorization-server metadata document is served. RFC 8414. */
export const AUTHORIZATION_SERVER_METADATA_PATH = '/.well-known/oauth-authorization-server';

/**
 * The scopes advertised in both metadata documents. `offline_access` is the one that
 * matters: Claude appends it to an authorization request only when the metadata lists
 * it, and it is what makes the flow issue a refresh token.
 */
const SCOPES_SUPPORTED = ['offline_access'];

/**
 * Every redirect URI the one registered client may send a code to.
 *
 * The first is the callback for every hosted Claude surface — claude.ai, Desktop, mobile
 * and Cowork. The other two are Claude Code, which is a native client on an RFC 8252
 * loopback redirect and binds whatever ephemeral port the OS gives it. The port is
 * absent here on purpose: the SDK's matcher ignores the port on a loopback URI and
 * requires an exact match on everything else, which the spike in #20 confirmed against a
 * request to `http://localhost:3118/callback`.
 *
 * This allowlist is most of what protects the endpoint. Anyone who learns the client id
 * can reach the consent screen, but every code minted goes to claude.ai or to the
 * operator's own machine.
 */
const REDIRECT_URIS = [
  'https://claude.ai/api/mcp/auth_callback',
  'http://localhost/callback',
  'http://127.0.0.1/callback',
];

/**
 * The rate-limit configuration for /authorize and /token.
 *
 * The SDK applies `express-rate-limit` to both — 100 requests per 15 minutes on
 * /authorize, 50 on /token — and keys on the client address. Two things make its default
 * keying wrong here. `trust proxy` is true, which this service needs behind Cloud Run, so
 * the default key is whatever X-Forwarded-For carries: forgeable, and the library logs a
 * ValidationError stack trace on every request rather than a line. Behind Cloud Run the
 * untrusted alternative is the front end's own address, which is the same for everyone.
 *
 * So the key is a constant and the bucket is global. That is the truth for a one-user
 * server, and it says so rather than looking per-caller. The cost is stated in
 * project-spec.md and accepted: anyone who finds the URL can spend the token endpoint's
 * 50 requests and block the operator's own connect flow for 15 minutes. No per-IP keying
 * fixes that when every request arrives from one proxy.
 */
const RATE_LIMIT = {
  keyGenerator: () => 'global',
  // The keying above is what the trustProxy validator warns about, and it no longer
  // applies: the key does not come from the request at all.
  validate: { trustProxy: false, keyGeneratorIpFallback: false },
};

/** The canonical URL of the MCP endpoint — the `resource` a token is bound to. */
export function mcpResourceUrl(oauth: OAuthConfig): URL {
  return new URL(`${oauth.publicUrl}${MCP_PATH}`);
}

/**
 * The URL of the protected-resource metadata document.
 *
 * Issue #23 hands this to `requireBearerAuth` as `resourceMetadataUrl`, which is what
 * puts it in the `WWW-Authenticate` header of a 401. Nothing here emits that header.
 */
export function protectedResourceMetadataUrl(oauth: OAuthConfig): string {
  return new URL(`/.well-known/oauth-protected-resource${MCP_PATH}`, oauth.publicUrl).href;
}

/**
 * The clients store: exactly one client, from configuration.
 *
 * No `registerClient`. See the note at the top of this file — its absence is what keeps
 * `registration_endpoint` out of the metadata.
 *
 * The secret is present, which is what makes this a confidential client: the SDK's
 * `authenticateClient` middleware makes a secret mandatory whenever the client record
 * carries one. It compares with `!==` rather than in constant time and that cannot be
 * changed without replacing the whole token router — see the note in #22 and in
 * project-spec.md.
 */
export function createClientsStore(oauth: OAuthConfig): OAuthRegisteredClientsStore {
  const client: OAuthClientInformationFull = {
    client_id: oauth.clientId,
    client_secret: oauth.clientSecret,
    redirect_uris: REDIRECT_URIS,
    // Declared for truthfulness rather than enforcement: the SDK's token handler does
    // not consult either list.
    grant_types: ['authorization_code', 'refresh_token'],
    response_types: ['code'],
    token_endpoint_auth_method: 'client_secret_post',
    scope: SCOPES_SUPPORTED.join(' '),
    client_name: 'Claude',
  };

  return {
    getClient(clientId: string): OAuthClientInformationFull | undefined {
      return clientId === client.client_id ? client : undefined;
    },
  };
}

/**
 * The provider that stands in until issue #22 writes the real one.
 *
 * Every method fails with `ServerError`, which the SDK turns into a 500 carrying
 * `error: "server_error"`. The routes therefore exist and answer in OAuth's own shape,
 * which is what this issue's acceptance asks for; no token can be issued and no request
 * can be authorised, which is the right failure while the provider is missing.
 */
export function unimplementedProvider(oauth: OAuthConfig): OAuthServerProvider {
  const clientsStore = createClientsStore(oauth);
  const refuse = (): never => {
    throw new ServerError('The authorization server is not configured yet');
  };

  return {
    get clientsStore() {
      return clientsStore;
    },
    authorize(_client: OAuthClientInformationFull, _params: AuthorizationParams, _res: Response) {
      return Promise.resolve(refuse());
    },
    challengeForAuthorizationCode(): Promise<string> {
      return Promise.resolve(refuse());
    },
    exchangeAuthorizationCode(): Promise<OAuthTokens> {
      return Promise.resolve(refuse());
    },
    exchangeRefreshToken(): Promise<OAuthTokens> {
      return Promise.resolve(refuse());
    },
    verifyAccessToken(): Promise<AuthInfo> {
      return Promise.resolve(refuse());
    },
  };
}

/**
 * Build the Layer-1 OAuth mount: both metadata documents, `/authorize` and `/token`.
 *
 * Mount it at the application root — see the note at the top of this file. `provider` is
 * what issue #22 supplies; the default refuses everything.
 */
export function oauthRouter(
  oauth: OAuthConfig,
  provider: OAuthServerProvider = unimplementedProvider(oauth),
): Router {
  const router = Router();
  const issuerUrl = new URL(oauth.publicUrl);

  // Registered ahead of mcpAuthRouter, so this document is the one served and the SDK's
  // copy of the same route is never reached.
  //
  // The only difference is `token_endpoint_auth_methods_supported`. The SDK builds that
  // array as ['client_secret_post', 'none'] from nothing the caller controls, so a
  // server that requires a client secret still advertises that it accepts clients with
  // none. Enforcement is correct either way — the client record above carries a secret,
  // which makes the secret mandatory — so this is a wrong statement in a public document
  // rather than a hole. It is rewritten rather than left, because the document is what a
  // client reads to decide whether to send the secret at all.
  //
  // Everything else is the SDK's own `createOAuthMetadata` output, so the two documents
  // cannot drift apart as the SDK changes; `metadataHandler` is the SDK's handler too,
  // which keeps the CORS headers and the GET/OPTIONS-only method handling identical to
  // the route it replaces.
  const metadata = {
    ...createOAuthMetadata({ provider, issuerUrl, scopesSupported: SCOPES_SUPPORTED }),
    token_endpoint_auth_methods_supported: ['client_secret_post'],
  };
  router.use(AUTHORIZATION_SERVER_METADATA_PATH, metadataHandler(metadata));

  router.use(
    mcpAuthRouter({
      provider,
      issuerUrl,
      // The protected resource is the MCP endpoint, not the origin. That is what puts
      // the document at /.well-known/oauth-protected-resource/mcp and what makes
      // `resource` carry the /mcp path, which is the value Claude sends on every
      // authorization and token request and the audience #23 checks a token against.
      resourceServerUrl: mcpResourceUrl(oauth),
      resourceName: 'OneNote MCP',
      scopesSupported: SCOPES_SUPPORTED,
      authorizationOptions: { rateLimit: RATE_LIMIT },
      tokenOptions: { rateLimit: RATE_LIMIT },
    }),
  );

  return router;
}
