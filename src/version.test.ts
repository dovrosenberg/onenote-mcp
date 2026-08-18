import test from 'node:test';
import assert from 'node:assert/strict';

import { SERVICE_NAME, VERSION } from './version.ts';

test('service name is stable', () => {
  assert.equal(SERVICE_NAME, 'onenote-mcp');
});

test('version is semver-shaped', () => {
  assert.match(VERSION, /^\d+\.\d+\.\d+$/);
});
