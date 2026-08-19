// Drives the whole Layer-1 flow over a real HTTP server through createApp, the way
// test/oauth-router.test.ts and test/mcp-server.test.ts do. Nothing here is a unit test
// of a handler: the redirect, the consent POST and the token exchange are only
// meaningful over the wire, because what is being checked is which status a request
// gets, which OAuth error code comes back in the body, and whether a Location header is
// present — and a direct call to the provider bypasses the SDK middleware that produces
// all three.
//
// The token format is re-implemented at the bottom of this file rather than imported.
// That is deliberate: it makes the wire format an assertion. A change to how a payload
// is signed breaks these tests, which is the point, because the only other thing that
// reads a token is a running instance of this server.
//
// What no test here can check is whether Claude accepts any of it. That waits for a real
// connect against the deployed URL.

import test, { after } from 'node:test';
import assert from 'node:assert/strict';
import { createHash, createHmac, randomBytes } from 'node:crypto';
import type { AddressInfo } from 'node:net';

import type { Config } from '../src/config.ts';
import { createOAuthProvider } from '../src/oauth-provider.ts';
import { createApp } from '../src/server.ts';

const PUBLIC_URL = 'https://onenote-mcp.example.run.app';
const CLIENT_ID = 'mcp-client';
const CLIENT_SECRET = 'mcp-secret';
const SIGNING_KEY = 'k'.repeat(32);
const REDIRECT_URI = 'https://claude.ai/api/mcp/auth_callback';

const OAUTH_CONFIG = {
  clientId: CLIENT_ID,
  clientSecret: CLIENT_SECRET,
  tokenSigningKey: SIGNING_KEY,
  publicUrl: PUBLIC_URL,
};

const STUB_CONFIG: Config = {
  graph: { clientId: 'client-id', authority: 'https://login.microsoftonline.com/common' },
  firestore: { cacheDocumentPath: 'tokencache/msal', projectId: 'proj' },
  oauth: OAUTH_CONFIG,
  server: { port: 0 },
};

const server = createApp(STUB_CONFIG).listen(0);
const ready = new Promise<void>((resolve) => server.once('listening', () => resolve()));

after(() => {
  server.close();
});

async function request(path: string, init: RequestInit = {}): Promise<Response> {
  await ready;
  const { port } = server.address() as AddressInfo;
  return fetch(`http://127.0.0.1:${port}${path}`, { redirect: 'manual', ...init });
}

function form(fields: Record<string, string>): RequestInit {
  return {
    method: 'POST',
    headers: { 'content-type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams(fields).toString(),
  };
}

/** A PKCE pair, S256 — the only method the metadata advertises. */
function pkce(): { verifier: string; challenge: string } {
  const verifier = randomBytes(32).toString('base64url');
  return { verifier, challenge: createHash('sha256').update(verifier).digest('base64url') };
}

/** GET /authorize, and pull the signed hidden field out of the consent page. */
async function consentForm(
  challenge: string,
  options: { redirectUri?: string; state?: string } = {},
): Promise<{ page: string; request: string }> {
  const query = new URLSearchParams({
    client_id: CLIENT_ID,
    redirect_uri: options.redirectUri ?? REDIRECT_URI,
    response_type: 'code',
    code_challenge: challenge,
    code_challenge_method: 'S256',
    scope: 'offline_access',
  });
  if (options.state !== undefined) query.set('state', options.state);

  const res = await request(`/authorize?${query.toString()}`);
  assert.equal(res.status, 200, 'the authorize endpoint did not render a page');
  const page = await res.text();

  const field = /name="request" value="([^"]+)"/.exec(page);
  assert.ok(field !== null, 'the consent page carries no signed request field');
  return { page, request: field[1]! };
}

/** POST the consent form and return the redirect it answers with. */
async function approve(signedRequest: string): Promise<Response> {
  return request('/consent', form({ request: signedRequest }));
}

/** Walk GET /authorize → POST /consent and return the minted code. */
async function codeFor(challenge: string, state?: string): Promise<string> {
  const { request: signed } = await consentForm(
    challenge,
    state === undefined ? {} : { state },
  );
  const res = await approve(signed);
  assert.equal(res.status, 302, 'the consent POST did not redirect');
  const location = res.headers.get('location');
  assert.ok(location !== null, 'the consent POST redirected with no Location');
  const code = new URL(location).searchParams.get('code');
  assert.ok(code !== null, `no code in the redirect: ${location}`);
  return code;
}

async function tokenRequest(fields: Record<string, string>): Promise<{
  status: number;
  body: Record<string, unknown>;
}> {
  const res = await request('/token', form({ client_id: CLIENT_ID, ...fields }));
  return { status: res.status, body: (await res.json()) as Record<string, unknown> };
}

test('the authorization request renders a consent page rather than redirecting', async () => {
  const { challenge } = pkce();
  const { page } = await consentForm(challenge);

  // The host is the only part of the request a person can meaningfully check, so it has
  // to be on the page. The button is what Anthropic requires to exist at all.
  assert.match(page, /claude\.ai/);
  assert.match(page, /<button[^>]*>\s*Approve\s*<\/button>/);

  // The whole request crosses the page in the signed field. Anything that leaked the
  // signing key into the page would make the field forgeable.
  assert.ok(!page.includes(SIGNING_KEY), 'the consent page carries the signing key');
  assert.ok(!page.includes(CLIENT_SECRET), 'the consent page carries the client secret');
});

test('the full authorization-code flow issues a token bound to the MCP endpoint', async () => {
  const { verifier, challenge } = pkce();
  const state = 'opaque-state-value';

  const { request: signed } = await consentForm(challenge, { state });
  const redirected = await approve(signed);
  assert.equal(redirected.status, 302);

  const location = new URL(redirected.headers.get('location')!);
  assert.equal(`${location.origin}${location.pathname}`, REDIRECT_URI);
  // Returned verbatim: Claude matches it against what it sent to detect a crossed flow.
  assert.equal(location.searchParams.get('state'), state);
  const code = location.searchParams.get('code')!;

  const { status, body } = await tokenRequest({
    grant_type: 'authorization_code',
    code,
    code_verifier: verifier,
    client_secret: CLIENT_SECRET,
    redirect_uri: REDIRECT_URI,
  });

  assert.equal(status, 200, JSON.stringify(body));
  assert.equal(body['token_type'], 'Bearer');
  assert.equal(body['expires_in'], 3600);
  assert.equal(typeof body['access_token'], 'string');
  // offline_access is in the metadata, so Claude asks for it and expects a refresh token
  // back. Without one the operator re-consents every hour.
  assert.equal(typeof body['refresh_token'], 'string');

  // The audience is the whole point of the token: it names the MCP endpoint including
  // its path, which is the value issue #23 compares the called URL against.
  const info = await createOAuthProvider(OAUTH_CONFIG).verifyAccessToken(body['access_token'] as string);
  assert.equal(info.resource?.href, `${PUBLIC_URL}/mcp`);
  assert.equal(info.clientId, CLIENT_ID);
  assert.deepEqual(info.scopes, ['offline_access']);
});

test('a wrong client secret is invalid_client, and no token is issued', async () => {
  const { verifier, challenge } = pkce();
  const code = await codeFor(challenge);

  const { status, body } = await tokenRequest({
    grant_type: 'authorization_code',
    code,
    code_verifier: verifier,
    client_secret: 'not-the-secret',
  });

  assert.equal(status, 400);
  assert.equal(body['error'], 'invalid_client');
  assert.equal('access_token' in body, false);
});

test('an authorization code cannot be used twice', async () => {
  const { verifier, challenge } = pkce();
  const code = await codeFor(challenge);

  const first = await tokenRequest({
    grant_type: 'authorization_code',
    code,
    code_verifier: verifier,
    client_secret: CLIENT_SECRET,
  });
  assert.equal(first.status, 200, JSON.stringify(first.body));

  const second = await tokenRequest({
    grant_type: 'authorization_code',
    code,
    code_verifier: verifier,
    client_secret: CLIENT_SECRET,
  });

  // invalid_grant, not invalid_request: a spent code is a grant that is no longer good,
  // and it is indistinguishable here from one that never existed.
  assert.equal(second.status, 400);
  assert.equal(second.body['error'], 'invalid_grant');
  assert.equal('access_token' in second.body, false);
});

test('a code_verifier that does not match the challenge is invalid_grant', async () => {
  const { challenge } = pkce();
  const other = pkce();
  const code = await codeFor(challenge);

  const { status, body } = await tokenRequest({
    grant_type: 'authorization_code',
    code,
    code_verifier: other.verifier,
    client_secret: CLIENT_SECRET,
  });

  assert.equal(status, 400);
  assert.equal(body['error'], 'invalid_grant');
  assert.equal('access_token' in body, false);
});

test('a redirect_uri that does not match the one the code was minted for is invalid_grant', async () => {
  const { verifier, challenge } = pkce();
  const code = await codeFor(challenge);

  const { status, body } = await tokenRequest({
    grant_type: 'authorization_code',
    code,
    code_verifier: verifier,
    client_secret: CLIENT_SECRET,
    // Registered, so the SDK accepts it as a URI; not the one this code came from.
    redirect_uri: 'http://localhost:3118/callback',
  });

  assert.equal(status, 400);
  assert.equal(body['error'], 'invalid_grant');
});

test('an edited hidden field is refused before any code is minted', async () => {
  const { challenge } = pkce();
  const { request: signed } = await consentForm(challenge);

  const [payload, signature] = signed.split('.');
  const decoded = JSON.parse(Buffer.from(payload!, 'base64url').toString('utf8')) as Record<
    string,
    unknown
  >;
  // The interesting edit: send the code somewhere else. Everything else on the form is
  // machinery, but this field decides who receives the authorization code.
  decoded['r'] = 'https://evil.example.com/callback';
  const edited = `${Buffer.from(JSON.stringify(decoded)).toString('base64url')}.${signature}`;

  const res = await approve(edited);

  // 400 rather than a redirect: the redirect URI is part of what failed to verify, so
  // there is nowhere to send the caller, and no code exists to send.
  assert.equal(res.status, 400);
  assert.equal(res.headers.get('location'), null);
  const page = await res.text();
  assert.ok(!page.includes('evil.example.com'), 'the refusal page echoed the edited value');
});

test('a hidden field signed with another key is refused', async () => {
  const { challenge } = pkce();
  const { request: signed } = await consentForm(challenge);

  const payload = signed.split('.')[0]!;
  const forged = `${payload}.${createHmac('sha256', 'a'.repeat(32)).update(payload).digest('base64url')}`;

  const res = await approve(forged);
  assert.equal(res.status, 400);
  assert.equal(res.headers.get('location'), null);
});

test('an expired consent form mints nothing', async () => {
  const { challenge } = pkce();

  // Signed with the real key, so only the expiry makes it bad. That is what separates
  // "this form is old" from "this form was tampered with".
  const stale = sign({
    k: 'q',
    c: CLIENT_ID,
    r: REDIRECT_URI,
    h: challenge,
    s: '',
    p: 'offline_access',
    x: Math.floor(Date.now() / 1000) - 1,
  });

  const res = await approve(stale);
  assert.equal(res.status, 400);
  assert.equal(res.headers.get('location'), null);
});

test('an unregistered redirect_uri is invalid_request, with no redirect', async () => {
  const { challenge } = pkce();
  const query = new URLSearchParams({
    client_id: CLIENT_ID,
    redirect_uri: 'https://evil.example.com/callback',
    response_type: 'code',
    code_challenge: challenge,
    code_challenge_method: 'S256',
  });

  const res = await request(`/authorize?${query.toString()}`);
  const body = (await res.json()) as Record<string, unknown>;

  // Refused directly rather than redirected: redirecting an error to an unregistered URI
  // is how an authorization server becomes an open redirector.
  assert.equal(res.status, 400);
  assert.equal(body['error'], 'invalid_request');
  assert.equal(res.headers.get('location'), null);
});

test('a refresh token buys a new access token and a fresh 30-day window', async () => {
  const { verifier, challenge } = pkce();
  const code = await codeFor(challenge);

  const first = await tokenRequest({
    grant_type: 'authorization_code',
    code,
    code_verifier: verifier,
    client_secret: CLIENT_SECRET,
  });
  assert.equal(first.status, 200, JSON.stringify(first.body));
  const refreshToken = first.body['refresh_token'] as string;

  const refreshed = await tokenRequest({
    grant_type: 'refresh_token',
    refresh_token: refreshToken,
    client_secret: CLIENT_SECRET,
  });

  assert.equal(refreshed.status, 200, JSON.stringify(refreshed.body));
  assert.equal(typeof refreshed.body['access_token'], 'string');

  // A new refresh token, not the one that was handed in. This is what makes the server
  // run unattended: the 30 days are counted from each refresh, so a connector Claude
  // refreshes every hour never returns to the consent screen, and only one left idle for
  // 30 days does.
  const renewed = refreshed.body['refresh_token'] as string;
  assert.notEqual(renewed, refreshToken);
  assert.ok(expiryOf(renewed) >= expiryOf(refreshToken), 'the new refresh token expires no later');

  const info = await createOAuthProvider(OAUTH_CONFIG).verifyAccessToken(
    refreshed.body['access_token'] as string,
  );
  assert.equal(info.resource?.href, `${PUBLIC_URL}/mcp`);

  // The renewed token works, and so does the one it replaced. That second half is not an
  // oversight: a stateless token carries its own expiry and there is no record to mark as
  // spent, so sliding the window cannot invalidate the previous token. Asserted rather
  // than left implicit, because it is the difference between this and real rotation.
  for (const token of [renewed, refreshToken]) {
    const again = await tokenRequest({
      grant_type: 'refresh_token',
      refresh_token: token,
      client_secret: CLIENT_SECRET,
    });
    assert.equal(again.status, 200, JSON.stringify(again.body));
  }
});

test('an unreadable or expired refresh token is invalid_grant', async () => {
  const expired = sign({
    k: 'r',
    c: CLIENT_ID,
    a: `${PUBLIC_URL}/mcp`,
    p: 'offline_access',
    x: Math.floor(Date.now() / 1000) - 1,
    n: 'nonce',
  });

  for (const refreshToken of ['not-a-token', expired]) {
    const { status, body } = await tokenRequest({
      grant_type: 'refresh_token',
      refresh_token: refreshToken,
      client_secret: CLIENT_SECRET,
    });

    // The code Claude keys its re-authentication on. invalid_request here would reach
    // the operator as a broken connector rather than as a prompt to sign in again.
    assert.equal(status, 400);
    assert.equal(body['error'], 'invalid_grant');
    assert.equal('access_token' in body, false);
  }
});

test('an access token is not accepted as a refresh token, and the reverse', async () => {
  const { verifier, challenge } = pkce();
  const code = await codeFor(challenge);
  const { body } = await tokenRequest({
    grant_type: 'authorization_code',
    code,
    code_verifier: verifier,
    client_secret: CLIENT_SECRET,
  });

  const asRefresh = await tokenRequest({
    grant_type: 'refresh_token',
    refresh_token: body['access_token'] as string,
    client_secret: CLIENT_SECRET,
  });
  assert.equal(asRefresh.status, 400);
  assert.equal(asRefresh.body['error'], 'invalid_grant');

  // Both are signed with the same key over the same shape, so the one-letter kind field
  // is the only thing separating them.
  const provider = createOAuthProvider(OAUTH_CONFIG);
  await assert.rejects(
    () => provider.verifyAccessToken(body['refresh_token'] as string),
    isInvalidToken,
  );
});

test('a token signed for another resource is refused', async () => {
  const provider = createOAuthProvider(OAUTH_CONFIG);
  const foreign = sign({
    k: 'a',
    c: CLIENT_ID,
    a: 'https://someone-else.example.run.app/mcp',
    p: 'offline_access',
    x: Math.floor(Date.now() / 1000) + 3600,
    n: 'nonce',
  });

  // Same signing key, right shape, wrong audience. The SDK carries AuthInfo.resource but
  // compares nothing; refusing here is what stops a token minted for another server
  // being replayed against this one.
  await assert.rejects(() => provider.verifyAccessToken(foreign), isInvalidToken);
});

test('the pending-code store is bounded, and drops the oldest code first', async () => {
  const { verifier, challenge } = pkce();
  const evicted = await codeFor(challenge);

  // MAX_PENDING_CODES is 100 in src/oauth-provider.ts. One consent POST per code, which
  // is the only way in — the store is not reachable any other way.
  //
  // The 100 filler forms are signed here rather than fetched from /authorize, because
  // that endpoint's rate limit is 100 requests per 15 minutes across every caller and
  // this one test would spend the whole file's budget. A form this test signs is byte
  // for byte what /authorize renders; the tests above are what check that.
  for (let i = 0; i < 100; i += 1) {
    const filler = pkce();
    const res = await approve(
      sign({
        k: 'q',
        c: CLIENT_ID,
        r: REDIRECT_URI,
        h: filler.challenge,
        s: '',
        p: 'offline_access',
        x: Math.floor(Date.now() / 1000) + 600,
      }),
    );
    assert.equal(res.status, 302);
  }

  const { status, body } = await tokenRequest({
    grant_type: 'authorization_code',
    code: evicted,
    code_verifier: verifier,
    client_secret: CLIENT_SECRET,
  });

  // The bound is what stops anyone who knows the client id from growing the map without
  // limit. Losing a code costs a retry of the consent click.
  assert.equal(status, 400);
  assert.equal(body['error'], 'invalid_grant');
});

test('both consent responses carry the headers that keep them out of caches and frames', async () => {
  const { challenge } = pkce();

  // The rendered form holds the signed authorization request; the POST's answer is a
  // redirect carrying an authorization code. Neither may sit in a shared cache, and
  // neither may leave the query string of /authorize — which carries `state` and the
  // PKCE challenge — in a Referer on the way to claude.ai.
  const rendered = await request(
    `/authorize?${new URLSearchParams({
      client_id: CLIENT_ID,
      redirect_uri: REDIRECT_URI,
      response_type: 'code',
      code_challenge: challenge,
      code_challenge_method: 'S256',
      scope: 'offline_access',
    }).toString()}`,
  );
  const refused = await approve('not-a-signed-field');

  for (const [label, res] of [
    ['the rendered form', rendered],
    ['the consent POST', refused],
  ] as const) {
    assert.equal(res.headers.get('cache-control'), 'no-store', label);
    assert.equal(res.headers.get('referrer-policy'), 'no-referrer', label);
    assert.equal(res.headers.get('x-frame-options'), 'DENY', label);
    assert.equal(res.headers.get('x-content-type-options'), 'nosniff', label);

    const csp = res.headers.get('content-security-policy') ?? '';
    assert.match(csp, /default-src 'none'/, label);
    assert.match(csp, /frame-ancestors 'none'/, label);
    // No form-action: browsers have disagreed about whether it is checked against a
    // redirect target, and the consent POST answers with a redirect to claude.ai. A
    // directive that might refuse that redirect would break the one page a human has to
    // get through.
    assert.doesNotMatch(csp, /form-action/, label);
  }
});

test('the consent route is rate limited, above the pending-code cap', async () => {
  // The route needs a limiter of its own: it is mounted ahead of the SDK's /authorize
  // limiter on purpose, and a rendered form stays postable for CONSENT_TTL_MS, so one
  // trip through /authorize yields a field that can be posted again and again.
  //
  // Its own server, because exhausting the limit on the shared one would leave every
  // later test in this file rate limited. The forms are invalid on purpose — a request
  // that is refused for its signature still counts against the limit, which is the
  // property worth having, since an attacker's requests are the invalid ones.
  const own = createApp(STUB_CONFIG).listen(0);
  await new Promise<void>((resolve) => own.once('listening', () => resolve()));
  const { port } = own.address() as AddressInfo;

  try {
    const post = (): Promise<Response> =>
      fetch(`http://127.0.0.1:${port}/consent`, { redirect: 'manual', ...form({ request: 'no' }) });

    const first = await post();
    assert.equal(first.status, 400);
    // Above MAX_PENDING_CODES, so the store's own eviction — the bound whose behaviour
    // is specified — is what a burst runs into first, and the limiter is the backstop.
    assert.equal(first.headers.get('x-ratelimit-limit'), '200');

    let refused: Response | undefined;
    for (let i = 0; i < 205 && refused === undefined; i += 1) {
      const res = await post();
      if (res.status === 429) refused = res;
    }

    assert.ok(refused !== undefined, 'the consent route accepted more than its limit');
  } finally {
    own.close();
  }
});

test('the metadata offers no revocation endpoint', async () => {
  const res = await request('/.well-known/oauth-authorization-server');
  const doc = (await res.json()) as Record<string, unknown>;

  // The SDK advertises it only when the provider implements revokeToken, and this one
  // does not: there is no record of a stateless token to delete. The operator's lever is
  // rotating MCP_TOKEN_SIGNING_KEY, which invalidates every outstanding token at once.
  assert.equal('revocation_endpoint' in doc, false);
});

/** The `x` field of a signed payload: its expiry, in seconds since the epoch. */
function expiryOf(token: string): number {
  const payload = JSON.parse(Buffer.from(token.split('.')[0]!, 'base64url').toString('utf8')) as Record<
    string,
    unknown
  >;
  assert.equal(typeof payload['x'], 'number', 'the payload carries no numeric expiry');
  return payload['x'] as number;
}

/** Matches the error the SDK's bearer middleware turns into a 401. */
function isInvalidToken(err: unknown): true {
  assert.ok(err instanceof Error, `not an error: ${String(err)}`);
  assert.equal((err as { errorCode?: string }).errorCode, 'invalid_token');
  return true;
}

/**
 * The token format, re-implemented: base64url(JSON) + '.' + base64url(HMAC-SHA256).
 *
 * Deliberately not imported from src/oauth-provider.ts — a test that signs with the
 * implementation's own function proves only that it agrees with itself.
 */
function sign(payload: Record<string, string | number>): string {
  const body = Buffer.from(JSON.stringify(payload)).toString('base64url');
  return `${body}.${createHmac('sha256', SIGNING_KEY).update(body).digest('base64url')}`;
}
