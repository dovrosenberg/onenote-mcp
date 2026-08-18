import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile, readdir } from 'node:fs/promises';
import path from 'node:path';

import {
  GRAPH_ROOT,
  GraphRequestError,
  GraphResponseError,
  GraphStructure,
  graphGet,
  type FetchLike,
  type TokenSource,
} from '../src/graph-structure.ts';
import { createGate } from '../src/graph-throttle.ts';

const TOKEN = 'fake-access-token';
const NODE_QUERY = '$select=id,displayName&$orderby=displayName';

const tokens: TokenSource = { getAccessToken: () => Promise.resolve(TOKEN) };

interface Call {
  url: string;
  authorization: string | undefined;
}

/**
 * A fetch whose routes are keyed by the exact URL. An unrouted URL fails the test, so
 * every assertion about behaviour is also an assertion about the URL that was built.
 */
function fakeFetch(routes: Record<string, () => Response>): {
  fetchImpl: FetchLike;
  calls: Call[];
} {
  const calls: Call[] = [];
  const fetchImpl: FetchLike = (url, init) => {
    calls.push({ url, authorization: init.headers['Authorization'] });
    const route = routes[url];
    if (route === undefined) {
      return Promise.reject(new Error(`no route for ${url}\nrouted: ${Object.keys(routes).join('\n        ')}`));
    }
    return Promise.resolve(route());
  };
  return { fetchImpl, calls };
}

function json(body: unknown, init: ResponseInit = {}): Response {
  return new Response(JSON.stringify(body), {
    status: 200,
    headers: { 'content-type': 'application/json' },
    ...init,
  });
}

function childUrl(kind: 'notebooks' | 'sectionGroups', id: string, child: string): string {
  return `${GRAPH_ROOT}/me/onenote/${kind}/${id}/${child}?${NODE_QUERY}`;
}

/**
 * One container with both child relationships inlined. `listContainerChildren` asks for
 * this instead of the two `childUrl` listings, which halves the cost of any walk.
 */
function containerUrl(kind: 'notebooks' | 'sectionGroups', id: string): string {
  return (
    `${GRAPH_ROOT}/me/onenote/${kind}/${id}?$select=id,displayName` +
    `&$expand=sections($select=id,displayName),sectionGroups($select=id,displayName)`
  );
}

/** The container response shape: the entity itself, with the children inlined. */
function container(
  id: string,
  displayName: string,
  sections: { id: string; displayName: string }[],
  sectionGroups: { id: string; displayName: string }[],
): Response {
  return json({ id, displayName, sections, sectionGroups });
}

function pagesUrl(sectionId: string, top: number): string {
  return (
    `${GRAPH_ROOT}/me/onenote/sections/${sectionId}/pages` +
    `?$top=${top}&$orderby=lastModifiedDateTime desc&$select=id,title,lastModifiedDateTime`
  );
}

async function caught(promise: Promise<unknown>): Promise<unknown> {
  try {
    await promise;
  } catch (err) {
    return err;
  }
  assert.fail('expected the call to reject');
}

test('graphGet sends the bearer token and returns the parsed object', async () => {
  const url = `${GRAPH_ROOT}/me/onenote/notebooks`;
  const { fetchImpl, calls } = fakeFetch({ [url]: () => json({ value: [] }) });

  assert.deepEqual(await graphGet(url, TOKEN, fetchImpl), { value: [] });
  assert.equal(calls.length, 1);
  assert.equal(calls[0]?.authorization, `Bearer ${TOKEN}`);
});

test('a non-2xx response throws GraphRequestError carrying the status and the body', async () => {
  // Shaped like the account-wide page list failure this module exists to avoid, because
  // that error is only distinguishable from any other 400 by the body text.
  const body = JSON.stringify({
    error: { code: '20266', message: 'Maximum number of sections exceeded.' },
  });
  const url = `${GRAPH_ROOT}/me/onenote/notebooks`;
  const { fetchImpl } = fakeFetch({
    [url]: () => new Response(body, { status: 400, statusText: 'Bad Request' }),
  });

  const err = await caught(graphGet(url, TOKEN, fetchImpl));

  assert.ok(err instanceof GraphRequestError, `expected GraphRequestError, got ${String(err)}`);
  assert.equal(err.name, 'GraphRequestError');
  assert.equal(err.status, 400);
  assert.equal(err.statusText, 'Bad Request');
  assert.equal(err.body, body);
  assert.match(err.message, /400 Bad Request/);
  assert.match(err.message, /20266/);
});

test('a 2xx body that is not a JSON object throws GraphResponseError', async () => {
  const url = `${GRAPH_ROOT}/me/onenote/notebooks`;
  const cases: Array<() => Response> = [
    () => new Response('<html>sign in</html>', { status: 200 }),
    () => json([{ id: 'nb-1' }]),
  ];

  for (const route of cases) {
    const { fetchImpl } = fakeFetch({ [url]: route });
    const err = await caught(graphGet(url, TOKEN, fetchImpl));
    assert.ok(err instanceof GraphResponseError, `expected GraphResponseError, got ${String(err)}`);
    assert.equal(err.url, url);
  }
});

test('listNotebooks follows @odata.nextLink instead of stopping at the first page', async () => {
  const first = `${GRAPH_ROOT}/me/onenote/notebooks?${NODE_QUERY}`;
  const second = `${GRAPH_ROOT}/me/onenote/notebooks?${NODE_QUERY}&$skiptoken=fake-cursor`;
  const { fetchImpl, calls } = fakeFetch({
    [first]: () =>
      json({
        value: [{ id: 'nb-1', displayName: '2025' }],
        '@odata.nextLink': second,
      }),
    [second]: () => json({ value: [{ id: 'nb-2', displayName: '2026' }] }),
  });

  const notebooks = await new GraphStructure(tokens, fetchImpl).listNotebooks();

  assert.deepEqual(notebooks, [
    { id: 'nb-1', displayName: '2025' },
    { id: 'nb-2', displayName: '2026' },
  ]);
  assert.deepEqual(
    calls.map((call) => call.url),
    [first, second],
  );
});

test('a nextLink that never stops is abandoned rather than followed forever', async () => {
  const url = `${GRAPH_ROOT}/me/onenote/notebooks?${NODE_QUERY}`;
  const { fetchImpl, calls } = fakeFetch({
    [url]: () => json({ value: [{ id: 'nb-1', displayName: '2025' }], '@odata.nextLink': url }),
  });

  const err = await caught(new GraphStructure(tokens, fetchImpl).listNotebooks());

  assert.ok(err instanceof GraphResponseError, `expected GraphResponseError, got ${String(err)}`);
  assert.match(err.message, /@odata.nextLink/);
  assert.ok(calls.length < 100, `expected the walk to stop, made ${calls.length} requests`);
});

test('listSections and listSectionGroups address both container kinds and encode the id', async () => {
  // Real OneNote ids contain characters that must not reach the path raw.
  const groupId = '1-abc!def/ghi';
  const encoded = encodeURIComponent(groupId);
  const { fetchImpl, calls } = fakeFetch({
    [childUrl('notebooks', 'nb-1', 'sections')]: () =>
      json({ value: [{ id: 's-1', displayName: 'Daily todo' }] }),
    [childUrl('sectionGroups', encoded, 'sectionGroups')]: () =>
      json({ value: [{ id: 'g-2', displayName: 'Week 1' }] }),
  });
  const client = new GraphStructure(tokens, fetchImpl);

  assert.deepEqual(await client.listSections('notebooks', 'nb-1'), [
    { id: 's-1', displayName: 'Daily todo' },
  ]);
  assert.deepEqual(await client.listSectionGroups('sectionGroups', groupId), [
    { id: 'g-2', displayName: 'Week 1' },
  ]);
  assert.equal(calls.length, 2);
  // The slash is the character that matters: raw, it would add a path segment and
  // address a different resource.
  assert.match(calls[1]?.url ?? '', /sectionGroups\/1-abc!def%2Fghi\/sectionGroups\?/);
});

test('listPagesInSection scopes to the section and sorts by last modified, newest first', async () => {
  const url = pagesUrl('s-1', 50);
  const { fetchImpl, calls } = fakeFetch({
    [url]: () =>
      json({
        value: [
          { id: 'p-1', title: 'Fake page', lastModifiedDateTime: '2026-02-01T00:00:00Z' },
          { id: 'p-2', title: '', lastModifiedDateTime: '2026-01-01T00:00:00Z' },
        ],
      }),
  });

  const pages = await new GraphStructure(tokens, fetchImpl).listPagesInSection('s-1');

  assert.deepEqual(pages, [
    { id: 'p-1', title: 'Fake page', lastModifiedDateTime: '2026-02-01T00:00:00Z' },
    { id: 'p-2', title: '', lastModifiedDateTime: '2026-01-01T00:00:00Z' },
  ]);
  const requested = calls[0]?.url ?? '';
  assert.match(requested, /\/sections\/s-1\/pages\?/);
  assert.match(requested, /\$orderby=lastModifiedDateTime desc/);
});

test('listPagesInSection stops once top items are in hand', async () => {
  const first = pagesUrl('s-1', 3);
  const second = `${first}&$skiptoken=fake-cursor`;
  const page = (id: string): Record<string, string> => ({
    id,
    title: `Fake ${id}`,
    lastModifiedDateTime: '2026-02-01T00:00:00Z',
  });
  const { fetchImpl, calls } = fakeFetch({
    [first]: () => json({ value: [page('p-1'), page('p-2')], '@odata.nextLink': second }),
    [second]: () => json({ value: [page('p-3'), page('p-4')], '@odata.nextLink': `${second}&more` }),
  });

  const pages = await new GraphStructure(tokens, fetchImpl).listPagesInSection('s-1', 3);

  assert.deepEqual(
    pages.map((p) => p.id),
    ['p-1', 'p-2', 'p-3'],
  );
  assert.equal(calls.length, 2, 'the third page must not be requested');
});

test('listPagesInSection rejects a top that is not a positive integer', async () => {
  const { fetchImpl, calls } = fakeFetch({});
  const client = new GraphStructure(tokens, fetchImpl);

  for (const top of [0, -1, 1.5, Number.NaN]) {
    assert.ok((await caught(client.listPagesInSection('s-1', top))) instanceof RangeError);
  }
  assert.equal(calls.length, 0);
});

test('getNotebookTree returns the full tree through a nested section group', async () => {
  // The acceptance case from issue #11: a notebook whose section group contains a
  // further section group. A walk that stopped one level down would miss s-3 entirely.
  const { fetchImpl, calls } = fakeFetch({
    [containerUrl('notebooks', 'nb-1')]: () =>
      container('nb-1', '2026', [{ id: 's-1', displayName: 'Year overview' }], [
        { id: 'g-1', displayName: 'March' },
      ]),
    [containerUrl('sectionGroups', 'g-1')]: () =>
      container('g-1', 'March', [{ id: 's-2', displayName: 'Monthly calendar' }], [
        { id: 'g-2', displayName: 'Week 1' },
      ]),
    [containerUrl('sectionGroups', 'g-2')]: () =>
      container('g-2', 'Week 1', [{ id: 's-3', displayName: 'Daily todo' }], []),
  });

  const tree = await new GraphStructure(tokens, fetchImpl).getNotebookTree({
    id: 'nb-1',
    displayName: '2026',
  });

  assert.deepEqual(tree, {
    id: 'nb-1',
    displayName: '2026',
    sections: [{ id: 's-1', displayName: 'Year overview' }],
    sectionGroups: [
      {
        id: 'g-1',
        displayName: 'March',
        sections: [{ id: 's-2', displayName: 'Monthly calendar' }],
        sectionGroups: [
          {
            id: 'g-2',
            displayName: 'Week 1',
            sections: [{ id: 's-3', displayName: 'Daily todo' }],
            sectionGroups: [],
          },
        ],
      },
    ],
  });

  // One request per container, not two: the children arrive with the container.
  assert.equal(calls.length, 3);
});

test('a section group that contains itself is abandoned rather than recursed forever', async () => {
  const { fetchImpl, calls } = fakeFetch({
    [containerUrl('notebooks', 'nb-1')]: () =>
      container('nb-1', '2026', [], [{ id: 'g-1', displayName: 'March' }]),
    // The group reports itself as its own child, which no real structure does.
    [containerUrl('sectionGroups', 'g-1')]: () =>
      container('g-1', 'March', [], [{ id: 'g-1', displayName: 'March' }]),
  });

  const err = await caught(
    new GraphStructure(tokens, fetchImpl).getNotebookTree({ id: 'nb-1', displayName: '2026' }),
  );

  assert.ok(err instanceof GraphResponseError, `expected GraphResponseError, got ${String(err)}`);
  assert.match(err.message, /nested deeper/);
  assert.ok(calls.length < 100, `expected the walk to stop, made ${calls.length} requests`);
});

test('a collection body with no value array, or an unusable item, throws GraphResponseError', async () => {
  const url = `${GRAPH_ROOT}/me/onenote/notebooks?${NODE_QUERY}`;
  const bodies: unknown[] = [
    { notValue: [] },
    { value: [{ displayName: 'no id' }] },
    { value: ['not an object'] },
  ];

  for (const body of bodies) {
    const { fetchImpl } = fakeFetch({ [url]: () => json(body) });
    const err = await caught(new GraphStructure(tokens, fetchImpl).listNotebooks());
    assert.ok(err instanceof GraphResponseError, `expected GraphResponseError, got ${String(err)}`);
  }
});

test('no module under src/ calls the account-wide page list', async () => {
  // The hard constraint of issue #11. `/me/onenote/pages` (the collection) fails with
  // error 20266 once the account has enough sections, and nothing at runtime can observe
  // a call that is absent, so this is a source-text check like the device-code one in
  // test/graph-auth.test.ts.
  //
  // `/me/onenote/pages/{id}/...` is a different endpoint and stays allowed: it addresses
  // one page, and issue #12's content fetch needs it. So the pattern bans the path only
  // when nothing follows the segment.
  const banned = /\/me\/onenote\/pages(?![\w/])/;

  const srcDir = path.join(import.meta.dirname, '..', 'src');
  const files = (await readdir(srcDir)).filter((name) => name.endsWith('.ts'));
  assert.ok(files.includes('graph-structure.ts'), 'expected to have scanned graph-structure.ts');

  for (const name of files) {
    const source = await readFile(path.join(srcDir, name), 'utf8');
    const code = stripComments(source);
    assert.doesNotMatch(
      code,
      banned,
      `src/${name} must not call the account-wide /me/onenote/pages list; scope page listing to /sections/{id}/pages`,
    );
  }

  // The stripper must not be what makes the check pass.
  assert.match(stripComments("const url = 'https://graph.microsoft.com/v1.0/me/onenote/pages'"), banned);
  assert.doesNotMatch(stripComments('// never call /me/onenote/pages'), banned);
  assert.doesNotMatch(stripComments('/* never call /me/onenote/pages */'), banned);
});

/**
 * Remove comments so a prose mention of the banned path is not read as a call site.
 *
 * `//` is only treated as a line comment when it is not preceded by a colon, so the
 * `https://` inside a URL literal does not truncate the rest of its own line — which is
 * the line a real violation would be on.
 */
function stripComments(source: string): string {
  return source.replace(/\/\*[\s\S]*?\*\//g, '').replace(/(^|[^:])\/\/.*$/gm, '$1');
}

// ---------------------------------------------------------------------------
// getExpandedTree. The URL is the assertion that matters most here: the $select inside
// each expand clause is worth 5.7x on the response, and the separator between $select
// and $expand inside one clause is a semicolon. Both were measured and confirmed against
// the real tenant; a comma there returns a 400 that no unit test would otherwise catch.
// ---------------------------------------------------------------------------

const EXPANDED_URL =
  `${GRAPH_ROOT}/me/onenote/notebooks?$select=id,displayName` +
  `&$expand=sections($select=id,displayName),` +
  `sectionGroups($select=id,displayName;$expand=sections($select=id,displayName))`;

test('getExpandedTree asks for the tree in one request and parses it', async () => {
  const { fetchImpl, calls } = fakeFetch({
    [EXPANDED_URL]: () =>
      json({
        value: [
          {
            id: 'nb-1',
            displayName: '2026',
            sections: [{ id: 'sec-1', displayName: 'Inbox' }],
            sectionGroups: [
              {
                id: 'grp-1',
                displayName: 'March',
                sections: [{ id: 'sec-2', displayName: 'Daily' }],
              },
            ],
          },
        ],
      }),
  });

  const tree = await new GraphStructure(tokens, fetchImpl).getExpandedTree();

  assert.equal(calls.length, 1, 'the whole point is that this is one request');
  assert.deepEqual(tree, [
    {
      id: 'nb-1',
      displayName: '2026',
      sections: [{ id: 'sec-1', displayName: 'Inbox' }],
      sectionGroups: [
        { id: 'grp-1', displayName: 'March', sections: [{ id: 'sec-2', displayName: 'Daily' }] },
      ],
    },
  ]);
});

test('an expanded relationship Graph omitted reads as empty, not as a fault', async () => {
  // Graph leaves the property off a relationship that holds nothing, and a notebook with
  // no section groups is ordinary.
  const { fetchImpl } = fakeFetch({
    [EXPANDED_URL]: () => json({ value: [{ id: 'nb-1', displayName: '2026' }] }),
  });

  const tree = await new GraphStructure(tokens, fetchImpl).getExpandedTree();
  assert.deepEqual(tree, [{ id: 'nb-1', displayName: '2026', sections: [], sectionGroups: [] }]);
});

test('an expanded relationship that is not an array is a GraphResponseError', async () => {
  const { fetchImpl } = fakeFetch({
    [EXPANDED_URL]: () => json({ value: [{ id: 'nb-1', displayName: '2026', sections: {} }] }),
  });

  await assert.rejects(
    () => new GraphStructure(tokens, fetchImpl).getExpandedTree(),
    GraphResponseError,
  );
});

// ---------------------------------------------------------------------------
// The two filtered lookups. Both URLs were confirmed against the real tenant, and both
// carry a limit that is not documented anywhere: `tolower()` is accepted on a section's
// pages and rejected on sections, and the account-wide sections endpoint answers 500
// unless a $filter is present.
// ---------------------------------------------------------------------------

const SECTIONS_BY_NAME =
  `${GRAPH_ROOT}/me/onenote/sections?$select=id,displayName` +
  `&$expand=parentNotebook($select=id,displayName),parentSectionGroup($select=id,displayName)` +
  `&$filter=`;

/** `contains(tolower(...))`, because `tolower(...) eq` answers 500 on this endpoint. */
function byNameUrl(lowered: string): string {
  return `${SECTIONS_BY_NAME}${encodeURIComponent(`contains(tolower(displayName), '${lowered}')`)}`;
}

test('findSectionsByName filters case-insensitively and returns the parents', async () => {
  const url = byNameUrl('monthly log');
  const { fetchImpl, calls } = fakeFetch({
    [url]: () =>
      json({
        value: [
          {
            id: 'sec-1',
            displayName: 'Monthly Log',
            parentNotebook: { id: 'nb-1', displayName: '2026' },
            parentSectionGroup: { id: 'grp-1', displayName: 'February' },
          },
          {
            id: 'sec-2',
            displayName: 'Monthly Log',
            parentNotebook: { id: 'nb-2', displayName: '2025' },
            parentSectionGroup: null,
          },
        ],
      }),
  });

  const sections = await new GraphStructure(tokens, fetchImpl).findSectionsByName('Monthly Log');

  assert.equal(calls.length, 1, 'one request reaches a section at any nesting depth');
  assert.deepEqual(sections[0]?.parentSectionGroup, { id: 'grp-1', displayName: 'February' });
  // A section directly under a notebook has no parent group, and that is not a fault.
  assert.equal(sections[1]?.parentSectionGroup, null);
  assert.deepEqual(sections[1]?.parentNotebook, { id: 'nb-2', displayName: '2025' });
});

test('findSectionsByName doubles a quote in the name rather than breaking the filter', async () => {
  const url = byNameUrl("bob''s notes");
  const { fetchImpl } = fakeFetch({ [url]: () => json({ value: [] }) });

  // An unescaped apostrophe would end the OData string literal early and produce a 400.
  await new GraphStructure(tokens, fetchImpl).findSectionsByName("Bob's notes");
});

test('findPagesByTitle lets Graph compare titles, case-insensitively', async () => {
  const filter = encodeURIComponent("tolower(title) eq 'monthly log'");
  const url =
    `${GRAPH_ROOT}/me/onenote/sections/sec-1/pages` +
    `?$select=id,title,lastModifiedDateTime&$filter=${filter}`;
  const { fetchImpl, calls } = fakeFetch({
    [url]: () =>
      json({
        value: [
          { id: 'p-1', title: 'Monthly Log', lastModifiedDateTime: '2026-03-04T10:00:00Z' },
        ],
      }),
  });

  // The caller's capitals are lowered before the comparison, because tolower() applies to
  // the stored title and not to the literal.
  const pages = await new GraphStructure(tokens, fetchImpl).findPagesByTitle('sec-1', 'MONTHLY Log');

  assert.equal(calls.length, 1);
  assert.deepEqual(
    pages.map((page) => page.id),
    ['p-1'],
  );
});

// ---------------------------------------------------------------------------
// The gate, seen from the client. The policy itself is tested in
// test/graph-throttle.test.ts; what matters here is that the client's requests actually
// go through it, and that a 429 carries the Retry-After the gate needs.
// ---------------------------------------------------------------------------

test('a 429 from Graph is retried through the gate rather than thrown at the caller', async () => {
  const url = `${GRAPH_ROOT}/me/onenote/notebooks?${NODE_QUERY}`;
  let attempts = 0;
  const { fetchImpl } = fakeFetch({
    [url]: () => {
      attempts += 1;
      if (attempts === 1) {
        return new Response('{"error":{"code":"10007"}}', {
          status: 429,
          headers: { 'content-type': 'application/json', 'retry-after': '3' },
        });
      }
      return json({ value: [{ id: 'nb-1', displayName: '2026' }] });
    },
  });

  const slept: number[] = [];
  const gate = createGate({
    minIntervalMs: 0,
    sleep: (ms: number) => {
      slept.push(ms);
      return Promise.resolve();
    },
  });

  const notebooks = await new GraphStructure(tokens, fetchImpl, gate).listNotebooks();

  assert.deepEqual(notebooks, [{ id: 'nb-1', displayName: '2026' }]);
  assert.equal(attempts, 2);
  assert.deepEqual(slept, [3_000], 'the Retry-After header is what the wait comes from');
});

test('a 404 is not retried, and reaches the caller as it is', async () => {
  const url = `${GRAPH_ROOT}/me/onenote/notebooks?${NODE_QUERY}`;
  let attempts = 0;
  const { fetchImpl } = fakeFetch({
    [url]: () => {
      attempts += 1;
      return new Response('{"error":{"code":"20112"}}', { status: 404 });
    },
  });

  const err = await caught(new GraphStructure(tokens, fetchImpl, createGate({ minIntervalMs: 0 })).listNotebooks());

  assert.ok(err instanceof GraphRequestError);
  assert.equal(err.status, 404);
  assert.equal(attempts, 1, 'repeating a 404 would spend quota to fail again');
});

test('listPagesInSection never asks Graph for more than 100, and pages to reach top', async () => {
  // $top=200 is a 400 with code 20129, so a caller asking for 150 has to be served by
  // following @odata.nextLink rather than by a bigger request.
  const first =
    `${GRAPH_ROOT}/me/onenote/sections/s-1/pages` +
    `?$top=100&$orderby=lastModifiedDateTime desc&$select=id,title,lastModifiedDateTime`;
  const next = `${GRAPH_ROOT}/me/onenote/sections/s-1/pages?$skiptoken=abc`;

  const pageAt = (index: number): Record<string, string> => ({
    id: `p-${index}`,
    title: `Page ${index}`,
    lastModifiedDateTime: '2026-03-04T10:00:00Z',
  });

  const { fetchImpl, calls } = fakeFetch({
    [first]: () =>
      json({
        value: Array.from({ length: 100 }, (_unused, index) => pageAt(index)),
        '@odata.nextLink': next,
      }),
    [next]: () =>
      json({ value: Array.from({ length: 100 }, (_unused, index) => pageAt(100 + index)) }),
  });

  const pages = await new GraphStructure(tokens, fetchImpl).listPagesInSection('s-1', 150);

  assert.equal(pages.length, 150);
  assert.equal(calls.length, 2);
  assert.ok(calls.every((call) => !call.url.includes('$top=150')));
});
