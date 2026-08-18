// Environment-variable schema and validation.
//
// Two entrypoints need different subsets of these variables: the server needs the
// Layer-1 OAuth credentials, the bootstrap CLI (issue #9) runs on the operator's own
// machine and must not require them. So validation is grouped rather than all-or-nothing.
//
// Nothing here reads process.env at module scope and nothing here exits the process
// except exitOnConfigError, which is only ever called from an entrypoint.

export type ConfigGroup = 'graph' | 'firestore' | 'oauth' | 'server';

export class ConfigError extends Error {
  readonly missing: readonly string[];
  readonly invalid: readonly string[];

  constructor(missing: string[], invalid: string[]) {
    super(formatConfigError(missing, invalid));
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
    name: 'PORT',
    group: 'server',
    required: false,
    fallback: '8080',
    purpose: 'Bind port; Cloud Run sets this',
    check: checkPort,
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
}

export interface ServerConfig {
  readonly port: number;
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

  for (const spec of SPECS) {
    if (!wanted.has(spec.group)) continue;

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
    throw new ConfigError(missing, invalid);
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
  if (wanted.has('firestore')) {
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
    };
  }
  if (wanted.has('server')) {
    config.server = { port: Number(required(values, 'PORT')) };
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

function formatConfigError(missing: string[], invalid: string[]): string {
  const lines: string[] = ['Configuration error. The server did not start.', ''];

  if (missing.length > 0) {
    lines.push(`Missing required environment variable${missing.length === 1 ? '' : 's'}:`);
    for (const name of missing) {
      lines.push(`  ${name} — ${purposeOf(name)}`);
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

function purposeOf(name: string): string {
  return SPECS.find((spec) => spec.name === name)?.purpose ?? 'no description available';
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
