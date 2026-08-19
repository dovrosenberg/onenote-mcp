// Environment-variable schema and validation.
//
// Two entrypoints need different subsets of these variables: the server needs the
// Layer-1 OAuth credentials, the bootstrap CLI runs on the operator's own machine and
// must not require them. So validation is grouped rather than all-or-nothing.
//
// The two entrypoints also differ on how strict the Firestore variables are, which is
// why there are two Firestore groups. The server asks for `firestore`, where
// FIRESTORE_CACHE_DOC defaults and GOOGLE_CLOUD_PROJECT may be absent because Cloud Run
// infers the project from the metadata server. The bootstrap CLI asks for
// `firestore-explicit`, where both are required: it runs against whatever project the
// operator's Application Default Credentials happen to point at, and seeding the cache
// into the wrong project or the wrong document path fails silently — the CLI reports
// success and the deployed server still finds no account.
//
// Nothing here reads process.env at module scope and nothing here exits the process
// except exitOnConfigError, which is only ever called from an entrypoint.

export type ConfigGroup = 'graph' | 'firestore' | 'firestore-explicit' | 'oauth' | 'server';

export class ConfigError extends Error {
  readonly missing: readonly string[];
  readonly invalid: readonly string[];

  /**
   * `purposes` maps a missing variable's name to the description printed beside it. It
   * is passed in rather than looked up from SPECS because a name can appear in more
   * than one group with a different reason for being required — FIRESTORE_CACHE_DOC is
   * optional in `firestore` and required in `firestore-explicit` — and a lookup by name
   * alone would print whichever entry happens to come first in the table.
   */
  constructor(missing: string[], invalid: string[], purposes: ReadonlyMap<string, string>) {
    super(formatConfigError(missing, invalid, purposes));
    this.name = 'ConfigError';
    this.missing = missing;
    this.invalid = invalid;
  }
}

interface VarSpec {
  readonly name: string;
  readonly group: ConfigGroup;
  readonly required: boolean;
  readonly fallback?: string;
  readonly purpose: string;
  /** Returns null when the value is well-formed, or a reason string when it is not. */
  readonly check?: (value: string) => string | null;
}

const MIN_SIGNING_KEY_LENGTH = 32;

const SPECS: readonly VarSpec[] = [
  {
    name: 'ONENOTE_CLIENT_ID',
    group: 'graph',
    required: true,
    purpose: 'Azure app registration client ID (public client)',
  },
  {
    name: 'ONENOTE_AUTHORITY',
    group: 'graph',
    required: true,
    purpose: 'Entra ID authority URL for the tenant',
    check: checkHttpsUrl,
  },
  {
    name: 'FIRESTORE_CACHE_DOC',
    group: 'firestore',
    required: false,
    fallback: 'tokencache/msal',
    purpose: 'Firestore document path holding the MSAL token cache',
    check: checkDocumentPath,
  },
  {
    name: 'GOOGLE_CLOUD_PROJECT',
    group: 'firestore',
    required: false,
    purpose: 'GCP project; inferred automatically on Cloud Run, needed when running locally',
  },
  {
    name: 'FIRESTORE_CACHE_DOC',
    group: 'firestore-explicit',
    required: true,
    purpose:
      'Firestore document path holding the MSAL token cache; the bootstrap CLI requires it rather than defaulting, so a local run cannot seed a document the deployed service does not read',
    check: checkDocumentPath,
  },
  {
    name: 'GOOGLE_CLOUD_PROJECT',
    group: 'firestore-explicit',
    required: true,
    purpose:
      'GCP project; the bootstrap CLI requires it rather than inferring, so a local run cannot seed the cache into whichever project Application Default Credentials point at',
  },
  {
    name: 'MCP_OAUTH_CLIENT_ID',
    group: 'oauth',
    required: true,
    purpose: 'Layer-1 OAuth client ID that Claude presents',
  },
  {
    name: 'MCP_OAUTH_CLIENT_SECRET',
    group: 'oauth',
    required: true,
    purpose: 'Layer-1 OAuth client secret',
  },
  {
    name: 'MCP_TOKEN_SIGNING_KEY',
    group: 'oauth',
    required: true,
    purpose: 'Key used to sign issued access tokens',
    check: checkSigningKey,
  },
  {
    name: 'MCP_PUBLIC_URL',
    group: 'oauth',
    required: true,
    purpose:
      "The service's own public URL, with no trailing slash; the OAuth issuer, the resource identifier and the resource-metadata URL are all derived from it",
    check: checkPublicUrl,
  },
  {
    name: 'PORT',
    group: 'server',
    required: false,
    fallback: '8080',
    purpose: 'Bind port; Cloud Run sets this',
    check: checkPort,
  },
  {
    name: 'MCP_KEEPALIVE_SECRET',
    group: 'server',
    required: false,
    purpose:
      'Shared secret a scheduler presents to POST /keepalive, which refreshes the Microsoft token so it does not expire from disuse; the route is not mounted when this is unset',
    check: checkSecretLength,
  },
];

export interface GraphConfig {
  readonly clientId: string;
  readonly authority: string;
}

export interface FirestoreConfig {
  readonly cacheDocumentPath: string;
  readonly projectId: string | undefined;
}

export interface OAuthConfig {
  readonly clientId: string;
  readonly clientSecret: string;
  readonly tokenSigningKey: string;
  /**
   * The service's own public origin, no trailing slash. Nothing on Cloud Run tells the
   * process what URL it is reached at, so it is configured rather than derived from a
   * request — a value taken from Host or X-Forwarded-Host is whatever the caller sent.
   */
  readonly publicUrl: string;
}

export interface ServerConfig {
  readonly port: number;
  /**
   * The secret POST /keepalive requires, or undefined to not serve that route at all.
   *
   * Optional because the service runs correctly without it — it just needs a human at a
   * browser again if the connector goes unused long enough for Microsoft's refresh token
   * to lapse. See the keepalive section of README.md.
   */
  readonly keepaliveSecret?: string;
}

export interface Config {
  readonly graph?: GraphConfig;
  readonly firestore?: FirestoreConfig;
  readonly oauth?: OAuthConfig;
  readonly server?: ServerConfig;
}

/**
 * Read and validate the variables belonging to `groups`.
 *
 * Accumulates every problem before throwing, so one run reports the full list rather
 * than making the operator rediscover them one at a time.
 *
 * @throws {ConfigError} if any required variable is absent or any present value is malformed.
 */
export function loadConfig(
  groups: readonly ConfigGroup[],
  env: NodeJS.ProcessEnv = process.env,
): Config {
  const wanted = new Set(groups);
  const missing: string[] = [];
  const invalid: string[] = [];
  const values = new Map<string, string>();
  const purposes = new Map<string, string>();

  for (const spec of SPECS) {
    if (!wanted.has(spec.group)) continue;
    purposes.set(spec.name, spec.purpose);

    const raw = env[spec.name];
    const trimmed = raw === undefined ? '' : raw.trim();

    if (trimmed === '') {
      // A variable set to whitespace is treated as unset. That is what an empty
      // GitHub secret produces once Cloud Run injects it.
      if (spec.required) {
        missing.push(spec.name);
      } else if (spec.fallback !== undefined) {
        values.set(spec.name, spec.fallback);
      }
      continue;
    }

    const reason = spec.check?.(trimmed) ?? null;
    if (reason !== null) {
      invalid.push(`${spec.name}: ${reason}`);
      continue;
    }

    values.set(spec.name, trimmed);
  }

  if (missing.length > 0 || invalid.length > 0) {
    throw new ConfigError(missing, invalid, purposes);
  }

  const config: {
    graph?: GraphConfig;
    firestore?: FirestoreConfig;
    oauth?: OAuthConfig;
    server?: ServerConfig;
  } = {};

  if (wanted.has('graph')) {
    config.graph = {
      clientId: required(values, 'ONENOTE_CLIENT_ID'),
      authority: required(values, 'ONENOTE_AUTHORITY'),
    };
  }
  if (wanted.has('firestore') || wanted.has('firestore-explicit')) {
    // Both groups produce the same shape. They differ only in whether the two names are
    // required, so asking for both at once is a caller error rather than a merge.
    config.firestore = {
      cacheDocumentPath: required(values, 'FIRESTORE_CACHE_DOC'),
      projectId: values.get('GOOGLE_CLOUD_PROJECT'),
    };
  }
  if (wanted.has('oauth')) {
    config.oauth = {
      clientId: required(values, 'MCP_OAUTH_CLIENT_ID'),
      clientSecret: required(values, 'MCP_OAUTH_CLIENT_SECRET'),
      tokenSigningKey: required(values, 'MCP_TOKEN_SIGNING_KEY'),
      publicUrl: required(values, 'MCP_PUBLIC_URL'),
    };
  }
  if (wanted.has('server')) {
    const keepaliveSecret = values.get('MCP_KEEPALIVE_SECRET');
    config.server = {
      port: Number(required(values, 'PORT')),
      // Spread rather than assigned as possibly-undefined: exactOptionalPropertyTypes
      // treats an explicit undefined as a different type from an absent property.
      ...(keepaliveSecret === undefined ? {} : { keepaliveSecret }),
    };
  }

  return config;
}

/**
 * Print a ConfigError and exit 1. Any other error is rethrown, so an unexpected
 * failure is never disguised as a configuration problem.
 */
export function exitOnConfigError(err: unknown): never {
  if (err instanceof ConfigError) {
    process.stderr.write(`${err.message}\n`);
    process.exit(1);
  }
  throw err;
}

function required(values: Map<string, string>, name: string): string {
  const value = values.get(name);
  if (value === undefined) {
    // Unreachable: loadConfig throws before this point if a required name is absent.
    throw new Error(`internal: ${name} missing after validation`);
  }
  return value;
}

function formatConfigError(
  missing: string[],
  invalid: string[],
  purposes: ReadonlyMap<string, string>,
): string {
  // Deliberately does not say "the server did not start" — the bootstrap CLI uses this
  // same error, and it is not a server.
  const lines: string[] = ['Configuration error. Startup aborted.', ''];

  if (missing.length > 0) {
    lines.push(`Missing required environment variable${missing.length === 1 ? '' : 's'}:`);
    for (const name of missing) {
      lines.push(`  ${name} — ${purposes.get(name) ?? 'no description available'}`);
    }
    lines.push('');
  }

  if (invalid.length > 0) {
    lines.push(`Invalid environment variable${invalid.length === 1 ? '' : 's'}:`);
    for (const entry of invalid) {
      lines.push(`  ${entry}`);
    }
    lines.push('');
  }

  lines.push('See the Configuration section of README.md.');
  return lines.join('\n');
}

function checkHttpsUrl(value: string): string | null {
  let url: URL;
  try {
    url = new URL(value);
  } catch {
    return `expected an absolute https URL, got ${JSON.stringify(value)}`;
  }
  if (url.protocol !== 'https:') {
    return `expected an https URL, got protocol ${JSON.stringify(url.protocol)}`;
  }
  return null;
}

/**
 * The issuer identifier of an OAuth authorization server: https, no query, no fragment
 * (RFC 8414 section 2). The trailing slash is rejected on top of that, because every URL
 * this service publishes is this value with a path concatenated onto it and a trailing
 * slash would land in the middle of the result.
 *
 * A path is not rejected, but the service is expected to sit at the root of an origin:
 * the SDK's OAuth router builds `/authorize` and `/token` as absolute paths from the
 * issuer, so a value carrying a path would advertise them one level above where they are
 * mounted. Cloud Run gives the service an origin of its own.
 */
function checkPublicUrl(value: string): string | null {
  const notHttps = checkHttpsUrl(value);
  if (notHttps !== null) return notHttps;

  if (value.endsWith('/')) {
    return `expected no trailing slash, got ${JSON.stringify(value)}`;
  }

  const url = new URL(value);
  if (url.search !== '') {
    return `expected no query string, got ${JSON.stringify(url.search)}`;
  }
  if (url.hash !== '') {
    return `expected no fragment, got ${JSON.stringify(url.hash)}`;
  }
  return null;
}

function checkDocumentPath(value: string): string | null {
  const segments = value.split('/');
  if (segments.some((segment) => segment === '')) {
    return `expected a Firestore document path with no empty segments, got ${JSON.stringify(value)}`;
  }
  if (segments.length % 2 !== 0) {
    // An odd segment count is a collection reference, not a document. Issue #7 writes
    // a single document, so this would only fail later, at the first token refresh.
    return `expected a document path with an even number of segments (collection/doc), got ${segments.length} in ${JSON.stringify(value)}`;
  }
  return null;
}

function checkSigningKey(value: string): string | null {
  return checkSecretLength(value);
}

/**
 * The one length rule both shared secrets are held to. It is a floor on how much work a
 * guess costs, and nothing here can check that the characters were chosen randomly.
 */
function checkSecretLength(value: string): string | null {
  if (value.length < MIN_SIGNING_KEY_LENGTH) {
    return `expected at least ${MIN_SIGNING_KEY_LENGTH} characters, got ${value.length}`;
  }
  return null;
}

function checkPort(value: string): string | null {
  if (!/^\d+$/.test(value)) {
    return `expected an integer, got ${JSON.stringify(value)}`;
  }
  const port = Number(value);
  if (port < 1 || port > 65535) {
    return `expected a port in 1-65535, got ${port}`;
  }
  return null;
}
