// The read side of the mirror, and the branch every covered tool goes through.
//
// Fakes throughout: a store that answers from plain maps, and a blob reader that answers
// from a map of bytes. No Firestore and no GCS, for the reason CLAUDE.md gives about
// src/mirror-store.ts — what those modules do is call, and what they decide lives here
// and in src/mirror-schema.ts.
//
// Almost every assertion below is about a **miss**, because a miss is the mirror's most
// important behaviour: it is what sends a read to Graph, and every way of getting it
// wrong produces a confident wrong answer rather than an error. The list, roughly in
// order of how bad the failure would be:
//
// - a page whose stored ink object has gone answers "no handwriting" instead of missing
// - a section group whose nested groups were never enumerated answers a short list that
//   looks complete
// - a write-marked-stale page serves pre-write content
// - a Firestore outage becomes a tool error instead of a slower answer

import test from 'node:test';
import assert from 'node:assert/strict';

import { setEventSink } from '../src/logging.ts';
import type { MirrorBlobReader } from '../src/mirror-blobs.ts';
import type {
  MirrorNotebook,
  MirrorPage,
  MirrorPageContent,
  MirrorSection,
  MirrorSectionGroup,
} from '../src/mirror-schema.ts';
import {
  MirrorReader,
  MirrorStructureEmptyError,
  mirrorLookupStructure,
  readSourced,
  type MirrorReadStore,
} from '../src/mirror-reader.ts';
import type { ScanResult } from '../src/mirror-store.ts';
import type { ReadFreshness, ReadSync } from '../src/read-sync.ts';

setEventSink(() => {});

const NB = 'nb-1';

function notebook(overrides: Partial<MirrorNotebook> = {}): MirrorNotebook {
  return {
    id: NB,
    displayName: '2026',
    mirrored: true,
    sectionCount: 1,
    sectionGroupCount: 0,
    graphLastModifiedDateTime: null,
    ...overrides,
  };
}

function section(overrides: Partial<MirrorSection> = {}): MirrorSection {
  return {
    id: 'sec-1',
    displayName: 'Daily',
    notebookId: NB,
    parentId: NB,
    parentKind: 'notebook',
    path: '2026 / Daily',
    mirrored: true,
    graphLastModifiedDateTime: null,
    pagesSyncedThrough: '2026-08-19T12:00:00Z',
    pageCount: 0,
    ...overrides,
  };
}

function group(overrides: Partial<MirrorSectionGroup> = {}): MirrorSectionGroup {
  return {
    id: 'grp-1',
    displayName: '062 - February',
    notebookId: NB,
    parentId: NB,
    parentKind: 'notebook',
    mirrored: true,
    path: '2026 / 062 - February',
    childGroupsKnown: true,
    ...overrides,
  };
}

function page(overrides: Partial<MirrorPage> = {}): MirrorPage {
  return {
    id: 'p-1',
    title: 'Monday',
    titleLower: 'monday',
    sectionId: 'sec-1',
    notebookId: NB,
    sectionPath: '2026 / Daily',
    lastModifiedDateTime: '2026-08-19T11:00:00Z',
    contentState: 'present',
    contentHash: 'abc',
    htmlLocation: 'firestore',
    htmlObject: null,
    htmlBytes: 20,
    ink: null,
    ...overrides,
  };
}

interface Fixture {
  notebooks: MirrorNotebook[];
  sections: MirrorSection[];
  groups: MirrorSectionGroup[];
  pages: MirrorPage[];
  content: Map<string, MirrorPageContent>;
  scanTruncated: boolean;
  throwOn?: string;
}

function fakeStore(fx: Partial<Fixture> = {}): { store: MirrorReadStore; data: Fixture } {
  const data: Fixture = {
    notebooks: [notebook()],
    sections: [section()],
    groups: [],
    pages: [page()],
    content: new Map([['p-1', { pageId: 'p-1', html: '<p>hi</p>', bytes: 9, contentHash: 'abc' }]]),
    scanTruncated: false,
    ...fx,
  };

  const guard = <T>(name: string, value: T): Promise<T> =>
    data.throwOn === name ? Promise.reject(new Error('firestore down')) : Promise.resolve(value);

  const store: MirrorReadStore = {
    listNotebooks: () => guard('listNotebooks', data.notebooks),
    listSectionsUnder: (parentId) =>
      guard('listSectionsUnder', data.sections.filter((s) => s.parentId === parentId)),
    listSectionGroupsUnder: (parentId) =>
      guard('listSectionGroupsUnder', data.groups.filter((g) => g.parentId === parentId)),
    listAllSections: () => guard('listAllSections', data.sections),
    listHeldSections: () =>
      guard(
        'listHeldSections',
        data.sections.filter((s) => (s.pendingWrites ?? 0) > 0),
      ),
    listAllSectionGroups: () => guard('listAllSectionGroups', data.groups),
    getNotebook: (id) => guard('getNotebook', data.notebooks.find((n) => n.id === id) ?? null),
    getSection: (id) => guard('getSection', data.sections.find((s) => s.id === id) ?? null),
    getSectionGroup: (id) => guard('getSectionGroup', data.groups.find((g) => g.id === id) ?? null),
    getPage: (id) => guard('getPage', data.pages.find((p) => p.id === id) ?? null),
    getPageContent: (id) => guard('getPageContent', data.content.get(id) ?? null),
    listPagesInSection: (sectionId, limit) =>
      guard(
        'listPagesInSection',
        data.pages.filter((p) => p.sectionId === sectionId).slice(0, limit),
      ),
    scanPages: (scope = {}) =>
      guard<ScanResult>('scanPages', {
        pages: data.pages.filter(
          (p) =>
            (scope.sectionId === undefined || p.sectionId === scope.sectionId) &&
            (scope.notebookId === undefined || p.notebookId === scope.notebookId),
        ),
        truncated: data.scanTruncated,
      }),
  };

  return { store, data };
}

function fakeBlobs(objects: Record<string, Uint8Array | string> = {}): MirrorBlobReader {
  return {
    getInk: (pageId) => {
      const value = objects[`ink/${pageId}`];
      return Promise.resolve(value instanceof Uint8Array ? value : null);
    },
    getHtml: (pageId) => {
      const value = objects[`html/${pageId}`];
      return Promise.resolve(typeof value === 'string' ? value : null);
    },
  };
}

function reader(fx: Partial<Fixture> = {}, objects: Record<string, Uint8Array | string> = {}): {
  reader: MirrorReader;
  data: Fixture;
} {
  const { store, data } = fakeStore(fx);
  return { reader: new MirrorReader(store, fakeBlobs(objects)), data };
}

// ---------------------------------------------------------------------------
// Structure
// ---------------------------------------------------------------------------

test('list_notebooks reports every notebook, flagging which have their pages held', () => {
  const { reader: r } = reader({
    notebooks: [notebook(), notebook({ id: 'nb-2', displayName: 'Other', mirrored: false })],
  });

  return r.listNotebooks().then((notebooks) => {
    assert.deepEqual(notebooks, [
      { id: NB, displayName: '2026', pagesMirrored: true },
      { id: 'nb-2', displayName: 'Other', pagesMirrored: false },
    ]);
  });
});

test('an empty notebook collection is a miss, not an empty account', async () => {
  // The sync has never run. Answering "no notebooks" would be indistinguishable from an
  // account that genuinely has none, and this tool takes no argument that could make the
  // difference detectable.
  const { reader: r } = reader({ notebooks: [] });
  assert.equal(await r.listNotebooks(), null);
});

test('an unknown container is a miss', async () => {
  const { reader: r } = reader();
  assert.equal(await r.listContainerChildren('notebooks', 'nb-nope'), null);
  assert.equal(await r.listContainerChildren('sectionGroups', 'grp-nope'), null);
});

test('a section group whose nested groups were never enumerated is a miss', async () => {
  // $expand reaches one level of section group, so a first-level group's nested groups
  // are absent from the tree response rather than known to be empty. Answering from that
  // document would omit them and look complete.
  const { reader: r } = reader({ groups: [group({ childGroupsKnown: false })] });
  assert.equal(await r.listContainerChildren('sectionGroups', 'grp-1'), null);

  const { reader: known } = reader({
    groups: [group({ childGroupsKnown: true })],
    sections: [section({ id: 'sec-2', parentId: 'grp-1', parentKind: 'sectionGroup' })],
  });
  const children = await known.listContainerChildren('sectionGroups', 'grp-1');
  assert.deepEqual(children?.sections, [{ id: 'sec-2', displayName: 'Daily' }]);
});

test('a section the mirror does not hold pages for is a miss', async () => {
  const { reader: r } = reader({ sections: [section({ mirrored: false })] });
  assert.equal(await r.listPagesInSection('sec-1', 10), null);
  assert.equal(await r.listPagesInSection('sec-nope', 10), null);
});

test('a section with a write in flight is a miss for its whole page listing', async () => {
  // What the hold covers that a stale page marker cannot: `create_page` adds a page the
  // mirror has no document for, and `update_page_title` changes the title every listing
  // and by-name lookup matches on. Serving the stored listing through either window is a
  // confident wrong answer -- "the page was not created", or the old title.
  const held = section({ pendingWrites: 1, pendingWritesSince: new Date().toISOString() });

  const { reader: r } = reader({ sections: [held] });
  assert.equal(await r.listPagesInSection('sec-1', 10), null);
  assert.equal(await r.searchTitles(() => true, { sectionId: 'sec-1' }), null);
  assert.equal(await r.searchTitles(() => true), null, 'an unscoped search covers it too');
  assert.equal(
    await r.searchTitles(() => true, { notebookId: NB }),
    null,
    'and so does a search scoped to the notebook holding it',
  );
});

test('a hold in another notebook does not spoil a search scoped away from it', async () => {
  // Precision is worth a query here: an unscoped search_pages answered by Graph costs up
  // to 61 requests, a seventh of the hourly budget.
  const { reader: r } = reader({
    notebooks: [notebook(), notebook({ id: 'nb-2' })],
    sections: [
      section(),
      section({
        id: 'sec-2',
        notebookId: 'nb-2',
        parentId: 'nb-2',
        pendingWrites: 1,
        pendingWritesSince: new Date().toISOString(),
      }),
    ],
  });

  assert.notEqual(await r.searchTitles(() => true, { notebookId: NB }), null);
  assert.notEqual(await r.listPagesInSection('sec-1', 10), null, 'sec-1 holds nothing back');
  assert.equal(await r.listPagesInSection('sec-2', 10), null, 'sec-2 does');
});

test('a hold a dead process left behind expires rather than wedging the section', async () => {
  // endWrite runs in a `finally`, so a hold outlives its write only when the process
  // stops between the two. Nothing runs after that to lower the count, and a hold nothing
  // can clear would send every listing for that section to Graph forever.
  const stale = section({
    pendingWrites: 1,
    pendingWritesSince: new Date(Date.now() - 3_600_000).toISOString(),
  });

  const { reader: r } = reader({ sections: [stale] });
  assert.notEqual(await r.listPagesInSection('sec-1', 10), null);
  assert.notEqual(await r.searchTitles(() => true), null);
});

test('list_pages reports an exact total, not a >= heuristic', async () => {
  // The Graph path can only say "I asked for top and got top, so probably more". The
  // mirror asks for one extra and counts.
  const pages = Array.from({ length: 5 }, (_u, i) => page({ id: `p-${i}` }));
  const { reader: r } = reader({ pages });

  const exactly = await r.listPagesInSection('sec-1', 5);
  assert.equal(exactly?.pages.length, 5);
  assert.equal(exactly?.total, 5, 'five held, five asked for: there is no sixth');

  const fewer = await r.listPagesInSection('sec-1', 3);
  assert.equal(fewer?.pages.length, 3);
  assert.equal(fewer?.total, 4, 'one more than asked for is enough to know there are more');
});

// ---------------------------------------------------------------------------
// Page content
// ---------------------------------------------------------------------------

test('page HTML is trimmed at read time, not at sync time', async () => {
  // Stored untrimmed so the trimmer can change without re-fetching every page from
  // Graph. This is the assertion that the stored form really is the raw one.
  const raw = '<div style="position:absolute;left:48px"><span style="font-weight:bold">hi</span></div>';
  const { reader: r } = reader({
    content: new Map([['p-1', { pageId: 'p-1', html: raw, bytes: raw.length, contentHash: 'abc' }]]),
  });

  const content = await r.getPageContent('p-1');
  assert.ok(content);
  assert.notEqual(content.html, raw, 'the stored form is raw');
  assert.ok(content.html?.includes('hi'), 'and the text survives the trim');
  assert.equal(content.html?.includes('font-weight'), false);
});

test('a stale or missing page is a miss', async () => {
  // `stale` is a write tool saying the stored copy is superseded. Serving it would hand
  // a model pre-write content with nothing saying so.
  for (const contentState of ['stale', 'missing'] as const) {
    const { reader: r } = reader({ pages: [page({ contentState })] });
    assert.equal(await r.getPageContent('p-1'), null, contentState);
  }
});

test('a page whose stored HTML has gone is a miss, not half a page', async () => {
  const { reader: r } = reader({ content: new Map() });
  assert.equal(await r.getPageContent('p-1'), null);
});

test('spilled HTML is fetched from the bucket', async () => {
  const { reader: r } = reader(
    { pages: [page({ htmlLocation: 'gcs', htmlObject: 'html/p-1.html', htmlBytes: 800_000 })] },
    { 'html/p-1': '<p>from the bucket</p>' },
  );

  const content = await r.getPageContent('p-1');
  assert.ok(content?.html?.includes('from the bucket'));
});

test('a page whose ink object has gone is a miss for the whole page', async () => {
  // The single worst failure available here. The document says there is handwriting and
  // the object is not there; answering "no handwriting" is a lie a model cannot detect,
  // and it would silently drop the only copy of what the page actually says.
  const { reader: r } = reader({
    pages: [
      page({
        ink: {
          objectName: 'ink/p-1.png',
          inkmlObjectName: 'inkml/p-1.xml',
          width: 1400,
          height: 900,
          strokeCount: 42,
          bytes: 1234,
        },
      }),
    ],
  });

  assert.equal(await r.getPageContent('p-1'), null);
});

test('stored ink is returned with the metadata the document holds', async () => {
  const png = new Uint8Array([137, 80, 78, 71, 1, 2, 3]);
  const { reader: r } = reader(
    {
      pages: [
        page({
          ink: {
            objectName: 'ink/p-1.png',
            inkmlObjectName: 'inkml/p-1.xml',
            width: 1400,
            height: 900,
            strokeCount: 42,
            bytes: png.byteLength,
          },
        }),
      ],
    },
    { 'ink/p-1': png },
  );

  const content = await r.getPageContent('p-1');
  assert.deepEqual(content?.ink?.png, png);
  assert.equal(content?.ink?.strokeCount, 42);
  assert.equal(content?.ink?.width, 1400);
});

// ---------------------------------------------------------------------------
// Search
// ---------------------------------------------------------------------------

test('a title scan reports its own scope, matched and scanned counts', async () => {
  const { reader: r } = reader({
    notebooks: [notebook(), notebook({ id: 'nb-2', mirrored: false })],
    pages: [page({ id: 'p-1', title: 'Monday notes' }), page({ id: 'p-2', title: 'Tuesday' })],
  });

  const found = await r.searchTitles((title) => title.toLowerCase().includes('mon'));

  assert.equal(found?.totalMatches, 1);
  assert.equal(found?.matches[0]?.pageId, 'p-1');
  assert.equal(found?.pagesScanned, 2);
  // The scope is the thing a model has to be told: page content is held for one of two
  // notebooks, so "no match" does not mean "no such page".
  assert.equal(found?.notebooksSearched, 1);
  assert.equal(found?.notebooksInAccount, 2);
});

test('a truncated scan says so', async () => {
  const { reader: r } = reader({ scanTruncated: true });
  const found = await r.searchTitles(() => true);
  assert.equal(found?.scanTruncated, true);
});

test('a scoped search on an unknown section is a miss', async () => {
  const { reader: r } = reader();
  assert.equal(await r.searchTitles(() => true, { sectionId: 'sec-nope' }), null);
});

// ---------------------------------------------------------------------------
// The lookup structure
// ---------------------------------------------------------------------------

test('the expanded tree is reassembled with sections under their real parents', async () => {
  const { reader: r } = reader({
    groups: [group()],
    sections: [
      section({ id: 'sec-1', parentId: NB, parentKind: 'notebook' }),
      section({ id: 'sec-2', displayName: 'Log', parentId: 'grp-1', parentKind: 'sectionGroup' }),
    ],
  });

  const tree = await r.expandedTree();

  assert.equal(tree?.length, 1);
  assert.deepEqual(tree?.[0]?.sections, [{ id: 'sec-1', displayName: 'Daily' }]);
  assert.deepEqual(tree?.[0]?.sectionGroups[0]?.sections, [{ id: 'sec-2', displayName: 'Log' }]);
});

test('an empty mirror throws rather than resolving to an empty tree', async () => {
  // resolveSection would turn an empty tree into a NameLookupError listing no siblings,
  // which reads to a model as "that section does not exist" rather than "the mirror is
  // empty". The caller catches this and goes to Graph.
  const lookup = mirrorLookupStructure(reader({ notebooks: [], sections: [] }).reader);

  await assert.rejects(() => lookup.getExpandedTree(), MirrorStructureEmptyError);
  await assert.rejects(() => lookup.findSectionsByName('Daily'), MirrorStructureEmptyError);
});

test('findSectionsByName matches a substring and carries the parents', async () => {
  const { reader: r } = reader({
    groups: [group()],
    sections: [section({ id: 'sec-2', displayName: 'Daily log', parentId: 'grp-1', parentKind: 'sectionGroup' })],
  });

  const found = await r.findSectionsByName('daily');

  assert.equal(found?.length, 1);
  assert.equal(found?.[0]?.parentNotebook?.displayName, '2026');
  assert.equal(found?.[0]?.parentSectionGroup?.displayName, '062 - February');
});

// ---------------------------------------------------------------------------
// readSourced
//
// Three steps in one function: refresh the mirror, read the mirror, report what the
// answer is worth. Most of what matters is the third — a wrong label is a confident wrong
// answer, where a wrong branch is only a slower right one.
// ---------------------------------------------------------------------------

interface Counts {
  mirror: number;
  graph: number;
  refresh: number;
}

function counting(counts: Counts, hit: unknown | null): {
  fromMirror: () => Promise<unknown | null>;
  fromGraph: () => Promise<unknown>;
} {
  return {
    fromMirror: () => {
      counts.mirror += 1;
      return Promise.resolve(hit);
    },
    fromGraph: () => {
      counts.graph += 1;
      return Promise.resolve('from graph');
    },
  };
}

/** A refresh that answers whatever it is told to, and counts how often it was asked. */
function fakeSync(counts: Counts, freshness: ReadFreshness): ReadSync {
  return {
    refresh: () => {
      counts.refresh += 1;
      return Promise.resolve(freshness);
    },
  };
}

test('a refreshed mirror answers as OneNote and never touches Graph', async () => {
  const counts: Counts = { mirror: 0, graph: 0, refresh: 0 };
  const { fromMirror, fromGraph } = counting(counts, 'from mirror');
  const { reader: r } = reader();

  const answer = await readSourced({
    tool: 'a_tool',
    mirror: r,
    sync: fakeSync(counts, 'current'),
    fromMirror,
    fromGraph,
  });

  // The label is the claim the caller acts on, and `origin` is what selects the shape.
  // They are separate fields because this case is both: mirror data, OneNote's content.
  assert.equal(answer.source, 'onenote');
  assert.equal(answer.origin, 'mirror');
  assert.equal(answer.data, 'from mirror');
  assert.equal(counts.graph, 0);
  assert.equal(counts.refresh, 1);
});

test('a refresh that did not finish downgrades the answer to mirror', async () => {
  // The budget ran out, the lease was held, the tree read failed — from here they are one
  // outcome, and reporting `onenote` for any of them would claim the copy is current when
  // nothing checked.
  const counts: Counts = { mirror: 0, graph: 0, refresh: 0 };
  const { fromMirror, fromGraph } = counting(counts, 'from mirror');
  const { reader: r } = reader();

  const answer = await readSourced({
    tool: 'a_tool',
    mirror: r,
    sync: fakeSync(counts, 'behind'),
    fromMirror,
    fromGraph,
  });

  assert.equal(answer.source, 'mirror');
  assert.equal(answer.origin, 'mirror');
  assert.equal(counts.graph, 0, 'a stale label is still an answer, not a fallback');
});

test('no sync bound is the same as a refresh that did not finish', async () => {
  const counts: Counts = { mirror: 0, graph: 0, refresh: 0 };
  const { fromMirror, fromGraph } = counting(counts, 'from mirror');
  const { reader: r } = reader();

  const answer = await readSourced({
    tool: 'a_tool',
    mirror: r,
    sync: undefined,
    fromMirror,
    fromGraph,
  });

  assert.equal(answer.source, 'mirror');
});

test('staleTracked reports OneNote even when the refresh did not finish', async () => {
  // get_page_content and nothing else. A page document carries its own contentState and
  // every write marks its page stale before touching OneNote, so a hit is a page nothing
  // has superseded — which is a fact about that page rather than about the account-wide
  // refresh.
  const counts: Counts = { mirror: 0, graph: 0, refresh: 0 };
  const { fromMirror, fromGraph } = counting(counts, 'from mirror');
  const { reader: r } = reader();

  const answer = await readSourced({
    tool: 'get_page_content',
    mirror: r,
    sync: fakeSync(counts, 'behind'),
    staleTracked: true,
    fromMirror,
    fromGraph,
  });

  assert.equal(answer.source, 'onenote');
  assert.equal(answer.origin, 'mirror');
});

test('the refresh runs before the mirror is read, not after', async () => {
  // A refresh that landed afterwards would have answered the question the caller already
  // asked with data it never saw.
  const order: string[] = [];
  const { reader: r } = reader();

  await readSourced({
    tool: 'a_tool',
    mirror: r,
    sync: {
      refresh: () => {
        order.push('refresh');
        return Promise.resolve('current');
      },
    },
    fromMirror: () => {
      order.push('read');
      return Promise.resolve('held');
    },
    fromGraph: () => Promise.resolve('live'),
  });

  assert.deepEqual(order, ['refresh', 'read']);
});

test('a miss falls through to Graph, which is OneNote by definition', async () => {
  const counts: Counts = { mirror: 0, graph: 0, refresh: 0 };
  const { fromMirror, fromGraph } = counting(counts, null);
  const { reader: r } = reader();

  const answer = await readSourced({
    tool: 'a_tool',
    mirror: r,
    sync: fakeSync(counts, 'current'),
    fromMirror,
    fromGraph,
  });

  assert.equal(answer.source, 'onenote');
  assert.equal(answer.origin, 'graph');
  assert.equal(counts.mirror, 1);
  assert.equal(counts.graph, 1);
});

test('a Firestore failure falls through to Graph rather than failing the call', async () => {
  // The mirror is an optimisation and Graph is the ground truth. Refusing a tool call
  // because a cache is down would be strictly worse than the behaviour before the mirror
  // existed.
  const counts: Counts = { mirror: 0, graph: 0, refresh: 0 };
  const { reader: r } = reader();

  const answer = await readSourced({
    tool: 'a_tool',
    mirror: r,
    sync: fakeSync(counts, 'current'),
    fromMirror: () => Promise.reject(new Error('firestore down')),
    fromGraph: () => {
      counts.graph += 1;
      return Promise.resolve('from graph');
    },
  });

  assert.equal(answer.source, 'onenote');
  assert.equal(answer.origin, 'graph');
  assert.equal(answer.data, 'from graph');
  assert.equal(counts.graph, 1);
});

test('no mirror at all means Graph, and no refresh is attempted', async () => {
  // MIRROR_READ_ENABLED=false. Every tool module treats an absent mirror as "always
  // Graph", which is what makes the flag a complete rollback — including the promise that
  // no tool call spends a Graph request on a sync it has no use for.
  const counts: Counts = { mirror: 0, graph: 0, refresh: 0 };
  const { fromMirror, fromGraph } = counting(counts, 'from mirror');

  const answer = await readSourced({
    tool: 'a_tool',
    mirror: undefined,
    sync: fakeSync(counts, 'current'),
    fromMirror,
    fromGraph,
  });

  assert.equal(answer.source, 'onenote');
  assert.equal(answer.origin, 'graph');
  assert.equal(counts.mirror, 0);
  assert.equal(counts.refresh, 0);
});

test('mirroredAt is reported on a hit and absent on a fallback', async () => {
  const { reader: r } = reader();
  const counts: Counts = { mirror: 0, graph: 0, refresh: 0 };

  const hit = await readSourced({
    tool: 'a_tool',
    mirror: r,
    sync: fakeSync(counts, 'current'),
    fromMirror: () => Promise.resolve('held'),
    fromGraph: () => Promise.resolve('live'),
    mirroredAt: () => Promise.resolve('2026-08-19T12:00:00.000Z'),
  });
  assert.equal(hit.mirroredAt, '2026-08-19T12:00:00.000Z');

  const missed = await readSourced({
    tool: 'a_tool',
    mirror: r,
    sync: fakeSync(counts, 'current'),
    fromMirror: () => Promise.resolve(null),
    fromGraph: () => Promise.resolve('live'),
    mirroredAt: () => Promise.resolve('2026-08-19T12:00:00.000Z'),
  });
  assert.equal(missed.mirroredAt, undefined, 'a Graph answer has no sync time');
});

test('a failure reading mirroredAt does not lose the answer', async () => {
  // It is a nicety. Losing the whole hit because a timestamp could not be read would
  // trade a real saving for a cosmetic field.
  const { reader: r } = reader();
  const counts: Counts = { mirror: 0, graph: 0, refresh: 0 };

  const answer = await readSourced({
    tool: 'a_tool',
    mirror: r,
    sync: fakeSync(counts, 'current'),
    fromMirror: () => Promise.resolve('held'),
    fromGraph: () => Promise.resolve('live'),
    mirroredAt: () => Promise.reject(new Error('firestore down')),
  });

  assert.equal(answer.source, 'onenote');
  assert.equal(answer.origin, 'mirror');
  assert.equal(answer.data, 'held');
  assert.equal(answer.mirroredAt, undefined);
});
