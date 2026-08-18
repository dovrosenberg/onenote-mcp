import test from 'node:test';
import assert from 'node:assert/strict';

import { GraphAuthError } from '../src/graph-auth.ts';
import { GraphRequestError, GraphResponseError } from '../src/graph-structure.ts';
import { InkRenderError } from '../src/ink.ts';
import {
  ToolInputError,
  indexTools,
  optionalInteger,
  optionalString,
  requiredString,
  toolDescriptors,
  toolErrorResult,
  type ToolDefinition,
} from '../src/mcp-tools.ts';

function tool(name: string): ToolDefinition {
  return {
    name,
    title: name,
    description: `The ${name} tool.`,
    inputSchema: { type: 'object', properties: {} },
    handle: async () => ({ content: [] }),
  };
}

function errorText(err: unknown): string {
  const result = toolErrorResult('a_tool', err);
  assert.equal(result.isError, true);
  const first = result.content[0];
  assert.equal(first?.type, 'text');
  return first?.type === 'text' ? first.text : '';
}

test('toolDescriptors drops the handler and keeps the declared shape', () => {
  const descriptors = toolDescriptors([tool('list_notebooks')]);
  assert.deepEqual(descriptors, [
    {
      name: 'list_notebooks',
      title: 'list_notebooks',
      description: 'The list_notebooks tool.',
      inputSchema: { type: 'object', properties: {} },
    },
  ]);
});

test('indexTools refuses two tools with the same name', () => {
  assert.throws(() => indexTools([tool('same'), tool('same')]), /same/);
});

test('an input error names the argument and the tool', () => {
  assert.match(errorText(new ToolInputError('pageId', 'is required')), /a_tool.*'pageId'/);
});

test('an auth failure keeps the instruction to re-run the bootstrap CLI', () => {
  const text = errorText(new GraphAuthError('no-account', 'tokencache/msal'));
  assert.match(text, /npm run bootstrap/);
});

test('a Graph request failure reports the status and the OData error code', () => {
  const body = JSON.stringify({ error: { code: '20266', message: 'maximum sections exceeded' } });
  const text = errorText(new GraphRequestError('https://graph/x', 400, 'Bad Request', body));
  assert.match(text, /400 Bad Request/);
  assert.match(text, /20266: maximum sections exceeded/);
});

test('a Graph body that is not an OData error is dropped rather than quoted', () => {
  const text = errorText(
    new GraphRequestError('https://graph/x', 502, 'Bad Gateway', '<html>proxy said no</html>'),
  );
  assert.match(text, /502 Bad Gateway/);
  assert.ok(!text.includes('proxy said no'), text);
});

test('a Graph error message longer than the cap is truncated', () => {
  const body = JSON.stringify({ error: { code: 'x', message: 'y'.repeat(1000) } });
  const text = errorText(new GraphRequestError('https://graph/x', 400, 'Bad Request', body));
  assert.ok(text.length < 500, `message not truncated: ${text.length} chars`);
});

test('a malformed Graph response and an ink failure keep their own messages', () => {
  assert.match(errorText(new GraphResponseError('body has no value array', 'u')), /no value array/);
  assert.match(errorText(new InkRenderError('resvg rejected the document')), /resvg rejected/);
});

test('an unmodelled error is reported without its message', () => {
  const text = errorText(new Error('PATCH body: the quick brown fox'));
  assert.ok(!text.includes('quick brown fox'), text);
  assert.match(text, /unexpected error/);
});

test('requiredString rejects a missing, blank, or non-string argument', () => {
  assert.equal(requiredString({ pageId: 'p-1' }, 'pageId'), 'p-1');
  for (const value of [undefined, null, '', '   ', 7, {}]) {
    assert.throws(() => requiredString({ pageId: value }, 'pageId'), ToolInputError);
  }
});

test('optionalString distinguishes absent from blank', () => {
  assert.equal(optionalString({}, 'sectionId'), undefined);
  assert.equal(optionalString({ sectionId: null }, 'sectionId'), undefined);
  assert.equal(optionalString({ sectionId: 's-1' }, 'sectionId'), 's-1');
  assert.throws(() => optionalString({ sectionId: '  ' }, 'sectionId'), ToolInputError);
});

test('optionalInteger enforces whole numbers inside the range', () => {
  const range = { min: 1, max: 100 };
  assert.equal(optionalInteger({}, 'top', range), undefined);
  assert.equal(optionalInteger({ top: 25 }, 'top', range), 25);
  for (const value of [0, 101, 2.5, '25']) {
    assert.throws(() => optionalInteger({ top: value }, 'top', range), ToolInputError);
  }
});
