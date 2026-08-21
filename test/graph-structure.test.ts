import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile, readdir } from 'node:fs/promises';
import path from 'node:path';

import {
  GRAPH_REQUEST_TIMEOUT_MS,
  GRAPH_ROOT,
  withRequestTimeout,
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

test('a nextLink pointing off the Graph origin is refused before the token is sent', async () => {
  // Every request this module makes carries the Graph access token in an Authorization
  // header, and a nextLink is the one URL that comes out of a response body rather than
  // being built here. Graph is what writes these links, so this is not expected to fire;
  // it is here so that the token cannot leave graph.microsoft.com whether or not that
  // stays true.
  const first = `${GRAPH_ROOT}/me/onenote/notebooks?${NODE_QUERY}`;

  const elsewhere = [
    'https://graph.microsoft.com.attacker.example/v1.0/me/onenote/notebooks',
    'http://graph.microsoft.com/v1.0/me/onenote/notebooks',
    'https://graph.microsoft.com:8443/v1.0/me/onenote/notebooks',
    'https://attacker.example/collect',
    '/me/onenote/notebooks?$skiptoken=relative',
  ];

  for (const next of elsewhere) {
    const { fetchImpl, calls } = fakeFetch({
      [first]: () => json({ value: [{ id: 'nb-1', displayName: '2025' }], '@odata.nextLink': next }),
    });

    const err = await caught(new GraphStructure(tokens, fetchImpl).listNotebooks());

    assert.ok(err instanceof GraphResponseError, `${next}: got ${String(err)}`);
    // Only the first request was made. The refusal happens before the second is issued,
    // not after a response comes back from somewhere else.
    assert.deepEqual(
      calls.map((call) => call.url),
      [first],
      next,
    );
    // The link itself is not quoted: it came from a response body, and this output can
    // reach a public log.
    assert.ok(!err.message.includes(next), `the message quoted the link: ${err.message}`);
  }
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

// lastModifiedDateTime is selected at the top level and inside both `sections` clauses,
// and deliberately not inside `sectionGroups` — nothing compares a section group's
// timestamp, and the field costs about 40% of the response (111,615 bytes against
// 79,660, measured 2026-08-19). It is what lets the #30 mirror visit only the sections
// that changed: a page create, edit or delete each move it, and nothing else does.
const EXPANDED_URL =
  `${GRAPH_ROOT}/me/onenote/notebooks?$select=id,displayName,lastModifiedDateTime` +
  `&$expand=sections($select=id,displayName,lastModifiedDateTime),` +
  `sectionGroups($select=id,displayName;$expand=sections($select=id,displayName,lastModifiedDateTime))`;

test('getExpandedTree asks for the tree in one request and parses it', async () => {
  const { fetchImpl, calls } = fakeFetch({
    [EXPANDED_URL]: () =>
      json({
        value: [
          {
            id: 'nb-1',
            displayName: '2026',
            lastModifiedDateTime: '2026-08-19T10:00:00Z',
            sections: [
              { id: 'sec-1', displayName: 'Inbox', lastModifiedDateTime: '2026-08-19T09:00:00Z' },
            ],
            sectionGroups: [
              {
                id: 'grp-1',
                displayName: 'March',
                sections: [
                  {
                    id: 'sec-2',
                    displayName: 'Daily',
                    lastModifiedDateTime: '2026-08-18T08:00:00Z',
                  },
                ],
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
      lastModifiedDateTime: '2026-08-19T10:00:00Z',
      sections: [
        { id: 'sec-1', displayName: 'Inbox', lastModifiedDateTime: '2026-08-19T09:00:00Z' },
      ],
      sectionGroups: [
        {
          id: 'grp-1',
          displayName: 'March',
          // No timestamp on the group itself: the URL does not ask for one.
          sections: [
            { id: 'sec-2', displayName: 'Daily', lastModifiedDateTime: '2026-08-18T08:00:00Z' },
          ],
        },
      ],
    },
  ]);
});

test('a section with no lastModifiedDateTime decodes without the key, not with undefined', async () => {
  // Absent is a real state, not a fault. The mirror reads it as "cannot tell whether this
  // section changed" and visits the section anyway, which is the same branch it would
  // take if the service stopped returning the field. `exactOptionalPropertyTypes` is why
  // the key must be absent rather than present-and-undefined.
  const { fetchImpl } = fakeFetch({
    [EXPANDED_URL]: () =>
      json({
        value: [
          {
            id: 'nb-1',
            displayName: '2026',
            sections: [{ id: 'sec-1', displayName: 'Inbox' }],
          },
        ],
      }),
  });

  const tree = await new GraphStructure(tokens, fetchImpl).getExpandedTree();

  assert.equal('lastModifiedDateTime' in tree[0]!, false);
  assert.equal('lastModifiedDateTime' in tree[0]!.sections[0]!, false);
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

// ---------------------------------------------------------------------------
// The two calls the #30 mirror adds. Both go through #collect, so both follow
// @odata.nextLink and both pass the shared gate.
// ---------------------------------------------------------------------------

test('listPagesChangedSince sends the datetime filter and no $orderby', async () => {
  // The filter is what makes the second tier one request per changed section rather than
  // a read-and-discard of the section's first hundred pages. Measured accepted
  // 2026-08-19; api-overview.md carries the run.
  //
  // No $orderby: the documented default for a section's pages is already
  // lastModifiedTime desc, and adding one would be a second unverified query option on a
  // call whose filter is the thing under suspicion.
  const url =
    `${GRAPH_ROOT}/me/onenote/sections/s-1/pages` +
    `?$select=id,title,lastModifiedDateTime&$top=100` +
    `&$filter=${encodeURIComponent('lastModifiedDateTime ge 2026-08-19T00:00:00Z')}`;

  const { fetchImpl, calls } = fakeFetch({
    [url]: () =>
      json({
        value: [
          { id: 'p-1', title: 'Monday', lastModifiedDateTime: '2026-08-19T12:00:00Z' },
          { id: 'p-2', title: 'Tuesday', lastModifiedDateTime: '2026-08-19T09:00:00Z' },
        ],
      }),
  });

  const pages = await new GraphStructure(tokens, fetchImpl).listPagesChangedSince(
    's-1',
    '2026-08-19T00:00:00Z',
  );

  assert.equal(calls.length, 1);
  assert.equal(calls[0]!.url.includes('$orderby'), false);
  assert.deepEqual(pages, [
    { id: 'p-1', title: 'Monday', lastModifiedDateTime: '2026-08-19T12:00:00Z' },
    { id: 'p-2', title: 'Tuesday', lastModifiedDateTime: '2026-08-19T09:00:00Z' },
  ]);
});

test('listPagesChangedSince does no date arithmetic of its own', async () => {
  // The watermark overlap belongs to the sync, which is the only thing that knows when
  // its pass started. This method sends the string it was handed, so a test of the
  // overlap rule cannot accidentally pass because two off-by-one hours cancelled out.
  const sinceIso = '1999-12-31T23:59:59Z';
  const url =
    `${GRAPH_ROOT}/me/onenote/sections/s-1/pages` +
    `?$select=id,title,lastModifiedDateTime&$top=100` +
    `&$filter=${encodeURIComponent(`lastModifiedDateTime ge ${sinceIso}`)}`;

  const { fetchImpl, calls } = fakeFetch({ [url]: () => json({ value: [] }) });

  await new GraphStructure(tokens, fetchImpl).listPagesChangedSince('s-1', sinceIso);
  assert.ok(calls[0]!.url.includes(encodeURIComponent(sinceIso)));
});

test('listPageSummaries asks for the three fields and follows every nextLink', async () => {
  // A sweep that stopped early would report pages as deleted that are merely past the
  // cutoff, which is the one mistake here that destroys data. So there is no `top`
  // argument and no bound but @odata.nextLink running out.
  const first =
    `${GRAPH_ROOT}/me/onenote/sections/s-1/pages?$select=id,title,lastModifiedDateTime&$top=100`;
  const next = `${GRAPH_ROOT}/me/onenote/sections/s-1/pages?$skiptoken=abc`;

  const { fetchImpl, calls } = fakeFetch({
    [first]: () =>
      json({
        value: Array.from({ length: 100 }, (_unused, index) => ({
          id: `p-${index}`,
          title: `Page ${index}`,
          lastModifiedDateTime: '2026-08-19T11:00:00Z',
        })),
        '@odata.nextLink': next,
      }),
    [next]: () =>
      json({
        value: [
          { id: 'p-100', title: 'Hundred', lastModifiedDateTime: '2026-08-19T12:00:00Z' },
          { id: 'p-101', title: 'Hundred and one', lastModifiedDateTime: '2026-08-19T13:00:00Z' },
        ],
      }),
  });

  const pages = await new GraphStructure(tokens, fetchImpl).listPageSummaries('s-1');

  assert.equal(calls.length, 2);
  assert.equal(pages.length, 102);
  assert.deepEqual(pages[0], {
    id: 'p-0',
    title: 'Page 0',
    lastModifiedDateTime: '2026-08-19T11:00:00Z',
  });
  // The title and the timestamp are why this is not `$select=id`: the sweep has no other
  // source for a discovered page's title, and no other way to notice content drift.
  assert.deepEqual(pages[101], {
    id: 'p-101',
    title: 'Hundred and one',
    lastModifiedDateTime: '2026-08-19T13:00:00Z',
  });
});

test('listPageSummaries rejects an item with no usable id rather than dropping it', async () => {
  // A dropped id reads as "this page is gone" and deletes the mirrored copy. Failing the
  // whole sweep for that section is the safe direction.
  const url =
    `${GRAPH_ROOT}/me/onenote/sections/s-1/pages?$select=id,title,lastModifiedDateTime&$top=100`;
  const { fetchImpl } = fakeFetch({
    [url]: () => json({ value: [{ id: 'p-1' }, { title: 'no id here' }] }),
  });

  await assert.rejects(
    () => new GraphStructure(tokens, fetchImpl).listPageSummaries('s-1'),
    GraphResponseError,
  );
});

test('the real fetch is given a timeout, so a hung call cannot hold a gate slot', async () => {
  // Node's fetch is undici, whose headersTimeout and bodyTimeout both default to 300
  // seconds — the same as Cloud Run's request timeout, and longer than the mirror sync's
  // entire 240-second budget. Four hung calls would hold every slot the gate has, and
  // neither budget can help: both are checked before an operation starts, and the
  // operation is the thing hanging.
  let seen: AbortSignal | undefined;
  const timed = withRequestTimeout(
    (_url, init) => {
      seen = init.signal;
      return Promise.resolve(new Response('{}', { status: 200 }));
    },
    50,
  );

  await timed('https://graph.microsoft.com/v1.0/x', { headers: {} });

  assert.ok(seen instanceof AbortSignal);
  assert.equal(seen.aborted, false);
  await new Promise((resolve) => setTimeout(resolve, 80));
  assert.equal(seen.aborted, true, 'the signal fires on its own, without the caller');
});

test('the timeout aborts a call that never answers, and it is not retried', async () => {
  // A service that has stopped answering is not helped by being asked again inside the
  // same run. The sync logs the page or section as failed, leaves its watermark alone,
  // and the next scheduled run tries it.
  const hangs = withRequestTimeout(
    (_url, init) =>
      new Promise<Response>((_resolve, reject) => {
        init.signal?.addEventListener('abort', () => {
          reject(new DOMException('The operation was aborted due to timeout', 'TimeoutError'));
        });
      }),
    20,
  );

  const gate = createGate({ maxConcurrent: 2, minIntervalMs: 0, maxRetries: 3 });
  let attempts = 0;

  await assert.rejects(
    () =>
      gate.run(() => {
        attempts += 1;
        return hangs('https://graph.microsoft.com/v1.0/x', { headers: {} });
      }),
    (err: unknown) => err instanceof DOMException && err.name === 'TimeoutError',
  );

  assert.equal(attempts, 1, 'a timeout carries no status, so retryWait declines it');
});

test('the timeout is far below both budgets it exists to protect', () => {
  // The property, not the number: 60s against a 240s sync budget and a 300s Cloud Run
  // request timeout. Raising it past either would put the wedge back.
  assert.ok(GRAPH_REQUEST_TIMEOUT_MS < 240_000);
  assert.ok(GRAPH_REQUEST_TIMEOUT_MS * 4 < 300_000, 'four hung calls still clear inside a request');
});
