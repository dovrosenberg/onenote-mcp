// The bootstrap CLI's testable behaviour is what it does before it needs a human and a
// network: it refuses to run on an incomplete environment. That is a process exit code
// and a message on stderr, so the test spawns it rather than importing it — the module
// signs in at import time and has no exported function to call.
//
// Nothing here covers the sign-in itself. That needs a browser, a real Entra tenant, and
// Firestore, and no credential that could stand in for one may be committed.

import test from 'node:test';
import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import path from 'node:path';
import { promisify } from 'node:util';

const execFileAsync = promisify(execFile);

const CLI = path.join(import.meta.dirname, '..', 'src', 'bootstrap.ts');

interface RunResult {
  readonly code: number;
  readonly stdout: string;
  readonly stderr: string;
}

async function runBootstrap(env: NodeJS.ProcessEnv): Promise<RunResult> {
  try {
    const { stdout, stderr } = await execFileAsync(process.execPath, [CLI], {
      // A bare env, plus PATH so the child can find its own runtime. Inheriting the
      // ambient environment would let a real ONENOTE_CLIENT_ID on the developer's
      // machine turn this into a test that tries to sign in.
      env: { PATH: process.env['PATH'] ?? '', ...env },
      timeout: 30_000,
    });
    return { code: 0, stdout, stderr };
  } catch (err) {
    const failure = err as { code?: number; stdout?: string; stderr?: string };
    return {
      code: failure.code ?? -1,
      stdout: failure.stdout ?? '',
      stderr: failure.stderr ?? '',
    };
  }
}

test('an empty environment exits 1 naming every missing variable, with no stack trace', async () => {
  const result = await runBootstrap({});

  assert.equal(result.code, 1);
  for (const name of [
    'ONENOTE_CLIENT_ID',
    'ONENOTE_AUTHORITY',
    'FIRESTORE_CACHE_DOC',
    'GOOGLE_CLOUD_PROJECT',
  ]) {
    assert.match(result.stderr, new RegExp(name), `expected ${name} in the error`);
  }
  assert.doesNotMatch(result.stderr, /\n\s*at .*\(/, 'a stack trace reached the operator');
});

test('the CLI does not require the Layer-1 OAuth credentials', async () => {
  // The operator runs this on their own machine and is never made to hold the client
  // secret. Only the Firestore names are reported as missing here.
  const result = await runBootstrap({
    ONENOTE_CLIENT_ID: 'client-id',
    ONENOTE_AUTHORITY: 'https://login.microsoftonline.com/common',
  });

  assert.equal(result.code, 1);
  assert.doesNotMatch(result.stderr, /MCP_OAUTH_CLIENT_ID/);
  assert.doesNotMatch(result.stderr, /MCP_OAUTH_CLIENT_SECRET/);
  assert.doesNotMatch(result.stderr, /MCP_TOKEN_SIGNING_KEY/);
  assert.match(result.stderr, /FIRESTORE_CACHE_DOC/);
  assert.match(result.stderr, /GOOGLE_CLOUD_PROJECT/);
});

test('an unset FIRESTORE_CACHE_DOC is refused rather than defaulted', async () => {
  // The server's group defaults this name to tokencache/msal. If the CLI inherited that
  // default it would seed a document the deployed service may not be reading, and would
  // print a success line saying so.
  const result = await runBootstrap({
    ONENOTE_CLIENT_ID: 'client-id',
    ONENOTE_AUTHORITY: 'https://login.microsoftonline.com/common',
    GOOGLE_CLOUD_PROJECT: 'some-project',
  });

  assert.equal(result.code, 1);
  assert.match(result.stderr, /FIRESTORE_CACHE_DOC/);
  assert.doesNotMatch(result.stdout, /Token cache seeded/);
});
