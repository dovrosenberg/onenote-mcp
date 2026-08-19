// The memoised Firestore client.
//
// This constructs the real `@google-cloud/firestore` client, which needs no credential
// and no backend because it connects lazily on the first read or write. That is the same
// property test/tools.test.ts already leans on, and if it ever stops being true both
// files fail together.
//
// What no test here covers is anything the client then does. Every Firestore call in
// this repository is untested for the reason CLAUDE.md gives about src/token-cache.ts:
// there is no emulator on this machine, and an in-memory fake would assert the fake
// rather than Firestore's transaction behaviour.

import test from 'node:test';
import assert from 'node:assert/strict';

import { firestoreFor, resetFirestoreClients } from '../src/firestore.ts';

test('one client is built per project and reused', () => {
  resetFirestoreClients();

  const first = firestoreFor({ cacheDocumentPath: 'tokencache/msal', projectId: 'proj-a' });
  const second = firestoreFor({ cacheDocumentPath: 'mirror/default', projectId: 'proj-a' });

  // The whole point: the token cache and the page mirror ask for different documents and
  // must get the same connection. Two clients would mean two gRPC channels and two
  // credential refreshers against one database.
  assert.equal(first, second);
});

test('a different project id gets a different client', () => {
  resetFirestoreClients();

  const a = firestoreFor({ cacheDocumentPath: 'tokencache/msal', projectId: 'proj-a' });
  const b = firestoreFor({ cacheDocumentPath: 'tokencache/msal', projectId: 'proj-b' });

  assert.notEqual(a, b);
});

test('an absent project id is its own key, not a wildcard', () => {
  resetFirestoreClients();

  // On Cloud Run the project comes from the metadata server, so projectId is undefined
  // and the client infers it. That inferred client must not be handed to a caller that
  // named a project explicitly.
  const inferred = firestoreFor({ cacheDocumentPath: 'tokencache/msal', projectId: undefined });
  const named = firestoreFor({ cacheDocumentPath: 'tokencache/msal', projectId: 'proj-a' });

  assert.notEqual(inferred, named);
  assert.equal(
    inferred,
    firestoreFor({ cacheDocumentPath: 'x/y', projectId: undefined }),
    'the inferred client is memoised too',
  );
});

test('resetFirestoreClients drops what was built', () => {
  resetFirestoreClients();

  const before = firestoreFor({ cacheDocumentPath: 'tokencache/msal', projectId: 'proj-a' });
  resetFirestoreClients();
  const after = firestoreFor({ cacheDocumentPath: 'tokencache/msal', projectId: 'proj-a' });

  assert.notEqual(before, after);
});
