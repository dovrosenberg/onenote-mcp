// The registry. This constructs the real MSAL and Firestore clients, which is what the
// deployed server does at startup; neither opens a connection until a token is asked
// for, so no credential and no backend is needed to run this.

import test from 'node:test';
import assert from 'node:assert/strict';

import type { Config } from '../src/config.ts';
import { indexTools } from '../src/mcp-tools.ts';
import { createGraphAuthFor, createTools } from '../src/tools.ts';

const STUB_CONFIG: Config = {
  graph: { clientId: 'client-id', authority: 'https://login.microsoftonline.com/common' },
  firestore: { cacheDocumentPath: 'tokencache/msal', projectId: 'proj' },
  oauth: {
    clientId: 'mcp-client',
    clientSecret: 'mcp-secret',
    tokenSigningKey: 'x'.repeat(32),
    publicUrl: 'https://onenote-mcp.example.run.app',
  },
  server: { port: 0 },
};

test('createTools exposes the browsing, reading and writing tools', () => {
  const tools = createTools(createGraphAuthFor(STUB_CONFIG));
  assert.deepEqual(
    tools.map((tool) => tool.name),
    [
      'list_notebooks',
      'list_sections',
      'list_pages',
      'search_pages',
      'find_page_by_name',
      'list_pages_by_name',
      'get_page_content',
      'append_to_page',
      'create_page',
      'update_page_title',
    ],
  );
  assert.doesNotThrow(() => indexTools(tools));
});

test('createGraphAuthFor refuses a config that was loaded without the Graph groups', () => {
  assert.throws(() => createGraphAuthFor({ server: { port: 0 } }), /graph/);
});

test('one auth object serves the whole process', () => {
  // The keepalive route and the tools share it. Two MSAL clients would mean two
  // in-memory access tokens and two writers of the same Firestore document.
  const auth = createGraphAuthFor(STUB_CONFIG);
  assert.equal(typeof auth.getAccessToken, 'function');
  assert.equal(typeof auth.refresh, 'function');
});
