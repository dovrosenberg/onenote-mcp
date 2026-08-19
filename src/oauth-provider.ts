// Layer-1 OAuth: the consent screen, the token format, and the authorization-code store.
//
// The protocol around this file is the SDK's. `mcpAuthRouter` validates `client_id`, the
// `redirect_uri` against the registered list, `response_type=code` and
// `code_challenge_method=S256`, authenticates the client at `/token`, and verifies the
// PKCE `code_verifier` against the challenge this module returns. None of that is
// re-implemented here. What is here is what the SDK issues nothing for: a page with an
// Approve button, an HMAC token format, and a Map of pending codes.
//
// Five things are decisions rather than details:
//
// - **Tokens are stateless.** An access token and a refresh token are each a signed
//   payload and nothing else; no store is consulted to verify one. That is what makes a
//   Cloud Run revision replacement invisible to a connected client — an in-memory token
//   store would force a reconnect on every deploy.
// - **Refresh tokens slide, and they are not rotated in the security sense.** Every
//   refresh mints a new refresh token carrying a fresh 30-day expiry, so a connector in
//   regular use never returns to the consent screen — the point of this service is to
//   run unattended, and Claude refreshes on its own. What sliding does *not* do is
//   invalidate the previous refresh token: it stays good until the expiry stamped in it,
//   because a stateless token has no record to mark as spent. Real rotation needs a
//   store, and a store is what makes a revision replacement force a reconnect. Rotation
//   is required for *public* clients in any case, and the secret in Claude's Advanced
//   settings makes this a confidential one. The 30 days therefore bound idleness rather
//   than the connection: a connector unused for 30 days needs one more click, and one
//   used weekly needs none.
// - **There is no `revokeToken`, deliberately.** The SDK advertises
//   `revocation_endpoint` only when the provider implements it, and with stateless
//   tokens there is nothing to revoke: no record exists to delete and no later request
//   consults one. The operator's actual lever is rotating `MCP_TOKEN_SIGNING_KEY`, which
//   invalidates every outstanding access token, refresh token and consent form at once.
// - **Authorization codes live in memory**, single-use, for 60 seconds. This service runs
//   `--max-instances=1`, so the instance that serves the consent POST is the one that
//   serves the token exchange a second later. A revision replacement in that window
//   loses the code and costs a retry of the consent click, which is the one moment in
//   the flow when a human is already present.
// - **The client secret is compared with `!==`,** inside the SDK's `authenticateClient`
//   middleware. That cannot be replaced without replacing the whole token router, so the
//   constant-time comparison originally asked for in issue #22 is not implemented. The
//   channel is a string comparison behind TLS, reachable only from the public internet,
//   under the token endpoint's 50-request rate limit. The signature comparisons in this
//   file are constant-time, because those this module does own.
//
// The consent screen authenticates nobody. Anyone holding the client id can reach it and
// cause a code to be minted; what protects the endpoint is that `/token` requires the
// client secret, that every code goes to a redirect URI on the allowlist in
// ./oauth-router.ts — claude.ai or loopback — and that PKCE binds the code to the client
// that started the flow.

import express, { Router, type Request, type Response } from 'express';
import { createHmac, randomBytes, timingSafeEqual } from 'node:crypto';

import { redirectUriMatches } from '@modelcontextprotocol/sdk/server/auth/handlers/authorize.js';
import { InvalidGrantError, InvalidTokenError } from '@modelcontextprotocol/sdk/server/auth/errors.js';
import type { AuthorizationParams } from '@modelcontextprotocol/sdk/server/auth/provider.js';
import type { AuthInfo } from '@modelcontextprotocol/sdk/server/auth/types.js';
import type {
  OAuthClientInformationFull,
  OAuthTokens,
} from '@modelcontextprotocol/sdk/shared/auth.js';

import type { OAuthConfig } from './config.ts';
import { createClientsStore, mcpResourceUrl, type Layer1Provider } from './oauth-router.ts';

/**
 * Where the consent form posts back to.
 *
 * It needs a route of this repository's own: the SDK's `/authorize` router answers the
 * request that renders the form and owns nothing that resumes the flow afterwards. The
 * path is deliberately not under `/authorize`, so the request never passes through that
 * router's rate limiter on its way here.
 */
export const CONSENT_PATH = '/consent';

/** Authorization-code lifetime. The code is exchanged seconds after it is minted. */
const AUTHORIZATION_CODE_TTL_MS = 60_000;

/**
 * How long a rendered consent form stays usable. It bounds the window in which an
 * abandoned browser tab can still mint a code, and it is generous because the person on
 * the other side of it may be reading.
 */
const CONSENT_TTL_MS = 10 * 60_000;

/** Access-token lifetime, in seconds. Claude refreshes reactively on a 401. */
const ACCESS_TOKEN_TTL_S = 60 * 60;

/**
 * Refresh-token lifetime, in seconds, counted from each refresh rather than from the
 * consent click — see the note at the top. It bounds how long a connector may sit idle
 * before a human has to approve again, not how long the connection may live.
 */
const REFRESH_TOKEN_TTL_S = 30 * 24 * 60 * 60;

/**
 * How many unexchanged codes may be held at once.
 *
 * The consent POST is reachable by anyone who holds the client id, and each request it
 * accepts costs a Map entry for 60 seconds. Expired entries are pruned on every insert;
 * this cap bounds what a burst inside one 60-second window can cost, by evicting the
 * oldest pending code. The eviction is invisible in normal use — one person clicks
 * Approve once, and the flow that loses a code retries the click.
 */
const MAX_PENDING_CODES = 100;

/** Payload discriminators. One byte each, because these ride in a URL and a form field. */
const KIND_ACCESS = 'a';
const KIND_REFRESH = 'r';
const KIND_CONSENT = 'q';

interface PendingCode {
  readonly clientId: string;
  readonly redirectUri: string;
  readonly codeChallenge: string;
  readonly scopes: readonly string[];
  readonly expiresAt: number;
}

/**
 * Build the Layer-1 provider. Pass it to `oauthRouter`, which mounts its consent route
 * alongside the SDK's `/authorize` and `/token`.
 */
export function createOAuthProvider(oauth: OAuthConfig): Layer1Provider {
  const clientsStore = createClientsStore(oauth);
  const key = oauth.tokenSigningKey;

  // The audience every token carries. It is this server's own resource identifier, taken
  // from configuration rather than from the `resource` parameter the request supplied —
  // a token minted here is good for this MCP endpoint and for nothing else, whatever was
  // asked for. Issue #23 compares this value against the endpoint being called.
  const audience = mcpResourceUrl(oauth).href;

  const codes = new Map<string, PendingCode>();

  function mintCode(pending: PendingCode): string {
    const now = Date.now();
    for (const [existing, record] of codes) {
      if (record.expiresAt <= now) codes.delete(existing);
    }
    while (codes.size >= MAX_PENDING_CODES) {
      // Insertion order: the oldest pending code goes first.
      const oldest = codes.keys().next();
      if (oldest.done === true) break;
      codes.delete(oldest.value);
    }

    const code = randomBytes(32).toString('base64url');
    codes.set(code, pending);
    return code;
  }

  /**
   * Read a pending code without consuming it. The SDK calls
   * `challengeForAuthorizationCode` and then `exchangeAuthorizationCode` on the same
   * code, so only the second consumes it.
   */
  function readCode(client: OAuthClientInformationFull, code: string): PendingCode {
    const pending = codes.get(code);
    if (pending === undefined) {
      // Also the answer for a code that was already exchanged. A caller cannot tell the
      // two apart, and neither can be acted on.
      throw new InvalidGrantError('Authorization code is invalid, expired or already used');
    }
    if (pending.expiresAt <= Date.now()) {
      codes.delete(code);
      throw new InvalidGrantError('Authorization code is invalid, expired or already used');
    }
    if (pending.clientId !== client.client_id) {
      throw new InvalidGrantError('Authorization code was issued to another client');
    }
    return pending;
  }

  function mintRefreshToken(clientId: string, scopes: readonly string[]): string {
    return sign(key, {
      k: KIND_REFRESH,
      c: clientId,
      a: audience,
      p: scopes.join(' '),
      x: expiryEpochSeconds(REFRESH_TOKEN_TTL_S),
      n: randomBytes(9).toString('base64url'),
    });
  }

  function issueTokens(clientId: string, scopes: readonly string[], refreshToken: string): OAuthTokens {
    return {
      access_token: sign(key, {
        k: KIND_ACCESS,
        c: clientId,
        a: audience,
        p: scopes.join(' '),
        x: expiryEpochSeconds(ACCESS_TOKEN_TTL_S),
        // Two tokens minted in the same second for the same client would otherwise be
        // the same bytes. Nothing depends on them differing; this makes them differ.
        n: randomBytes(9).toString('base64url'),
      }),
      token_type: 'Bearer',
      expires_in: ACCESS_TOKEN_TTL_S,
      scope: scopes.join(' '),
      refresh_token: refreshToken,
    };
  }

  const router = Router();
  router.post(CONSENT_PATH, express.urlencoded({ extended: false }), async (req: Request, res: Response) => {
    res.setHeader('Cache-Control', 'no-store');

    const body: unknown = req.body;
    const submitted =
      typeof body === 'object' && body !== null ? (body as Record<string, unknown>)['request'] : undefined;
    const payload = typeof submitted === 'string' ? unsign(key, submitted) : null;

    // A form field that was edited, signed with another key, or left in a tab for more
    // than CONSENT_TTL_MS. Answered here rather than redirected: the redirect URI is
    // part of what failed to verify, so there is nowhere trustworthy to send the caller,
    // and no code is minted.
    if (payload === null || readString(payload, 'k') !== KIND_CONSENT || isExpired(payload)) {
      res.status(400).type('html').send(refusalPage());
      return;
    }

    const clientId = readString(payload, 'c');
    const redirectUri = readString(payload, 'r');
    const codeChallenge = readString(payload, 'h');
    if (clientId === null || redirectUri === null || codeChallenge === null) {
      res.status(400).type('html').send(refusalPage());
      return;
    }

    // Re-checked rather than trusted from the signature. The signature proves this
    // server built the field; it does not prove the allowlist still holds the URI, and
    // the key outlives a change to it.
    const registered = await clientsStore.getClient(clientId);
    if (
      registered === undefined ||
      !registered.redirect_uris.some((allowed) => redirectUriMatches(redirectUri, allowed))
    ) {
      res.status(400).type('html').send(refusalPage());
      return;
    }

    const state = readString(payload, 's') ?? '';
    const scopeString = readString(payload, 'p') ?? '';
    const code = mintCode({
      clientId,
      redirectUri,
      codeChallenge,
      scopes: splitScopes(scopeString),
      expiresAt: Date.now() + AUTHORIZATION_CODE_TTL_MS,
    });

    const target = new URL(redirectUri);
    target.searchParams.set('code', code);
    if (state !== '') target.searchParams.set('state', state);
    res.redirect(302, target.href);
  });

  return {
    get clientsStore() {
      return clientsStore;
    },

    consentRouter: router,

    /**
     * Renders the consent page rather than redirecting.
     *
     * Everything needed to resume the flow crosses the page in one signed hidden field:
     * the client id, the redirect URI, the PKCE challenge, the state and the scopes. It
     * is signed with `MCP_TOKEN_SIGNING_KEY` so the form cannot be edited, and it is
     * carried in the form rather than held in memory so an instance replacement between
     * rendering and clicking does not break the flow.
     *
     * The `resource` the request asked for is not carried: this server mints tokens for
     * its own MCP endpoint and no other.
     */
    async authorize(
      client: OAuthClientInformationFull,
      params: AuthorizationParams,
      res: Response,
    ): Promise<void> {
      const request = sign(key, {
        k: KIND_CONSENT,
        c: client.client_id,
        r: params.redirectUri,
        h: params.codeChallenge,
        s: params.state ?? '',
        p: (params.scopes ?? []).join(' '),
        x: expiryEpochSeconds(Math.floor(CONSENT_TTL_MS / 1000)),
      });

      res
        .status(200)
        .type('html')
        .send(consentPage(`${oauth.publicUrl}${CONSENT_PATH}`, request, new URL(params.redirectUri).host));
    },

    async challengeForAuthorizationCode(
      client: OAuthClientInformationFull,
      authorizationCode: string,
    ): Promise<string> {
      return readCode(client, authorizationCode).codeChallenge;
    },

    async exchangeAuthorizationCode(
      client: OAuthClientInformationFull,
      authorizationCode: string,
      _codeVerifier?: string,
      redirectUri?: string,
    ): Promise<OAuthTokens> {
      // The verifier is not read here: `skipLocalPkceValidation` is left off, so the SDK
      // has already checked it against the challenge returned above.
      const pending = readCode(client, authorizationCode);
      // Single use. Consumed before anything else can fail, so a request that gets as
      // far as a real error still spends the code.
      codes.delete(authorizationCode);

      if (redirectUri !== undefined && redirectUri !== pending.redirectUri) {
        throw new InvalidGrantError('redirect_uri does not match the authorization request');
      }

      return issueTokens(
        client.client_id,
        pending.scopes,
        mintRefreshToken(client.client_id, pending.scopes),
      );
    },

    /**
     * Exchanges a refresh token for a new access token and a new refresh token, the
     * second carrying a fresh 30-day expiry — see the note at the top of this file. The
     * token handed in stays valid until its own expiry; nothing here can revoke it.
     *
     * Every failure is `invalid_grant`, never `invalid_request` and never a custom code.
     * That is the code Claude keys its re-authentication on: anything else reaches the
     * operator as a broken connector rather than as a prompt to sign in again.
     */
    async exchangeRefreshToken(
      client: OAuthClientInformationFull,
      refreshToken: string,
      scopes?: string[],
    ): Promise<OAuthTokens> {
      const payload = unsign(key, refreshToken);
      if (payload === null || readString(payload, 'k') !== KIND_REFRESH || isExpired(payload)) {
        throw new InvalidGrantError('Refresh token is invalid or expired');
      }
      if (readString(payload, 'c') !== client.client_id) {
        throw new InvalidGrantError('Refresh token was issued to another client');
      }
      if (readString(payload, 'a') !== audience) {
        throw new InvalidGrantError('Refresh token was issued for another resource');
      }

      const granted = splitScopes(readString(payload, 'p') ?? '');
      // A refresh may narrow the scopes but never widen them (RFC 6749 section 6). The
      // narrowing applies to the access token only: the new refresh token carries the
      // originally granted scopes, so one narrow request does not permanently narrow the
      // connection.
      const requested = scopes === undefined ? granted : scopes.filter((scope) => granted.includes(scope));

      return issueTokens(client.client_id, requested, mintRefreshToken(client.client_id, granted));
    },

    /**
     * Verifies an access token. Issue #23 wires this to the MCP route through the SDK's
     * `requireBearerAuth`, which turns `InvalidTokenError` into the 401 that carries
     * `WWW-Authenticate`.
     *
     * The audience is checked here because the payload carries it: a token signed by
     * this key for another resource identifier is refused rather than accepted and
     * compared later.
     */
    async verifyAccessToken(token: string): Promise<AuthInfo> {
      const payload = unsign(key, token);
      if (payload === null || readString(payload, 'k') !== KIND_ACCESS || isExpired(payload)) {
        throw new InvalidTokenError('Access token is invalid or expired');
      }
      const clientId = readString(payload, 'c');
      if (clientId === null || readString(payload, 'a') !== audience) {
        throw new InvalidTokenError('Access token is invalid or expired');
      }

      return {
        token,
        clientId,
        scopes: splitScopes(readString(payload, 'p') ?? ''),
        expiresAt: readNumber(payload, 'x') ?? 0,
        resource: new URL(audience),
      };
    },
  };
}

/**
 * Sign a compact payload: `base64url(JSON).base64url(HMAC-SHA256)`.
 *
 * Not a JWT. Nothing outside this server ever reads one of these, so there is no
 * interoperability to buy and no library to add — `node:crypto` and a one-letter key per
 * field. The keys are short because an access token rides in every request header.
 */
function sign(key: string, payload: Record<string, string | number>): string {
  const body = Buffer.from(JSON.stringify(payload)).toString('base64url');
  return `${body}.${mac(key, body)}`;
}

/**
 * Verify a signed payload and return its fields, or null if the signature does not
 * match, the token is malformed, or the body is not a JSON object.
 *
 * Returning null rather than throwing is deliberate: every caller turns a failure into
 * its own OAuth error, and which one differs by endpoint.
 */
function unsign(key: string, token: string): Record<string, unknown> | null {
  const separator = token.indexOf('.');
  if (separator <= 0) return null;

  const body = token.slice(0, separator);
  const provided = Buffer.from(token.slice(separator + 1), 'utf8');
  const expected = Buffer.from(mac(key, body), 'utf8');
  // Length is checked first because timingSafeEqual throws on a mismatch. The length of
  // an HMAC-SHA256 digest is not a secret.
  if (provided.length !== expected.length || !timingSafeEqual(provided, expected)) return null;

  try {
    const parsed: unknown = JSON.parse(Buffer.from(body, 'base64url').toString('utf8'));
    if (typeof parsed !== 'object' || parsed === null || Array.isArray(parsed)) return null;
    return parsed as Record<string, unknown>;
  } catch {
    return null;
  }
}

function mac(key: string, body: string): string {
  return createHmac('sha256', key).update(body).digest('base64url');
}

function expiryEpochSeconds(ttlSeconds: number): number {
  return Math.floor(Date.now() / 1000) + ttlSeconds;
}

/** True when the payload has no `x` field or its expiry has passed. */
function isExpired(payload: Record<string, unknown>): boolean {
  const expiry = readNumber(payload, 'x');
  return expiry === null || expiry * 1000 <= Date.now();
}

function readString(payload: Record<string, unknown>, field: string): string | null {
  const value = payload[field];
  return typeof value === 'string' ? value : null;
}

function readNumber(payload: Record<string, unknown>, field: string): number | null {
  const value = payload[field];
  return typeof value === 'number' && Number.isFinite(value) ? value : null;
}

function splitScopes(value: string): string[] {
  return value === '' ? [] : value.split(' ');
}

/**
 * The consent page.
 *
 * One button. It names what is being granted and the host the code will be sent to,
 * because that host is the only part of the request a person can meaningfully check —
 * `claude.ai` for a hosted Claude surface, `localhost` for Claude Code. Everything else
 * in the request is opaque machinery.
 *
 * `escapeHtml` is applied to both interpolations. The signed field is base64url and the
 * host comes from a URL the SDK has already matched against the allowlist, so neither
 * can carry markup today; escaping them is what keeps that true if either ever changes.
 */
function consentPage(formAction: string, request: string, redirectHost: string): string {
  return page(
    'Approve access',
    `<h1>Approve access to your OneNote</h1>
    <p>
      Approving lets this connector read and write the notebooks, sections and pages of
      the OneNote account this server was set up with, including rendered handwriting.
    </p>
    <p class="detail">
      The authorization code will be sent to <strong>${escapeHtml(redirectHost)}</strong>.
      If that is not where you started signing in, close this page.
    </p>
    <form method="post" action="${escapeHtml(formAction)}">
      <input type="hidden" name="request" value="${escapeHtml(request)}">
      <button type="submit">Approve</button>
    </form>`,
  );
}

/** Shown when the signed field fails to verify. No code has been minted at this point. */
function refusalPage(): string {
  return page(
    'Request refused',
    `<h1>Request refused</h1>
    <p>
      This approval form could not be verified. It may have been altered, or it may have
      been open for too long.
    </p>
    <p class="detail">Start the connection again from the client that sent you here.</p>`,
  );
}

function page(title: string, body: string): string {
  return `<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>${escapeHtml(title)}</title>
<style>
  body { font-family: system-ui, sans-serif; margin: 0 auto; max-width: 34rem; padding: 3rem 1.5rem; line-height: 1.5; }
  h1 { font-size: 1.4rem; }
  .detail { color: #555; font-size: 0.9rem; }
  button { font-size: 1rem; padding: 0.6rem 1.6rem; margin-top: 1rem; cursor: pointer; }
</style>
</head>
<body>
${body}
</body>
</html>
`;
}

function escapeHtml(value: string): string {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}
