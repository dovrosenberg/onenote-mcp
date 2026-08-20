// The read side of the mirror, and the one helper every covered tool branches through.
//
// **Every method here answers `null` on a miss**, and a miss always means "ask Graph".
// That is the whole contract. The mirror is an optimisation and Graph is the ground
// truth, so a page this does not hold, a section it has not seen, a Firestore outage, and
// a caller that asked for live data all end in the same place: the tool's existing Graph
// path, with the result saying so.
//
// A miss is only detectable when the tool names something the mirror can fail to find.
// `list_notebooks` takes no arguments and unscoped `search_pages` takes only a query, so
// neither can miss — which is why the sync stores structure for the **whole account**
// rather than only the selected notebooks. Page *content* follows the selection; notebook
// and section names do not. Without that, both tools would answer confidently and
// partially from a mirror holding three notebooks out of fifty-five, and a partial answer
// that cannot be detected as partial is the failure CLAUDE.md names about truncated
// searches. Unscoped `search_pages` still reports `notebooksSearched` against
// `notebooksInAccount`, because its *pages* really are a subset.
//
// The branch lives in the tool modules rather than behind an adapter implementing
// `StructureClient`. An adapter returning `PageSummary[]` has nowhere to say who
// answered; reporting the source would need either state on an object `createTools`
// shares across every request — unsafe the moment two calls overlap — or a widening of
// every narrow interface's return type, which would change every existing fake in
// test/structure-tools.test.ts, test/page-tools.test.ts and test/page-search.test.ts.
// And "which source answered" is part of the tool's contract with the model, exactly like
// `moreAvailable`, `stoppedEarly` and `deepSearchUsed`.

import type {
  ContainerChildren,
  ContainerKind,
  ExpandedNotebook,
  Notebook,
  PageSummary,
  Section,
  SectionWithParents,
} from './graph-structure.ts';
import { fitInkToByteBudget, type InkImage } from './ink.ts';
import { logEvent } from './logging.ts';
import type { MirrorBlobReader } from './mirror-blobs.ts';
import {
  timestampToIso,
  type MirrorNotebook,
  type MirrorPage,
  type MirrorPageContent,
  type MirrorSection,
  type MirrorSectionGroup,
} from './mirror-schema.ts';
import type { ScanResult } from './mirror-store.ts';
import { trimPageHtml } from './page-html.ts';
import type { PageContent } from './page-content.ts';

export type MirrorSource = 'mirror' | 'graph';

/**
 * The forcing argument, defined once so the seven tool descriptions cannot drift.
 *
 * It lives here rather than in either tool module because both need it and neither
 * should import the other. Named for what the caller wants rather than for the
 * implementation: a model asking for live data does not need to know a Firestore mirror
 * exists, only that there is a faster copy and a slower authority.
 */
export const USE_LIVE_DATA_PROPERTY = {
  type: 'boolean' as const,
  description:
    "Read directly from OneNote instead of this server's local copy. The local copy is " +
    'refreshed every few minutes, so pass true only when you need an edit made in the ' +
    'last few minutes. It is slower and spends the OneNote request budget.',
};

/**
 * An answer, and where it came from. Every covered tool puts both in its JSON.
 *
 * A discriminated union over two type parameters rather than one, because the two sides
 * do not always answer with the same shape: `search_pages` from Graph reports
 * `sectionsSearched` and `stoppedEarly` about a walk, and from the mirror reports
 * `notebooksSearched` and `scanTruncated` about a scan. Narrowing on `source` is what
 * lets a handler read the right one without a cast — and a cast is exactly how a field
 * that means nothing on one path ends up in the other path's result.
 *
 * Most tools use it with M and G the same, and infer both.
 */
export type Sourced<M, G = M> =
  | { readonly source: 'mirror'; readonly data: M; readonly mirroredAt?: string }
  | { readonly source: 'graph'; readonly data: G; readonly mirroredAt?: undefined };

/** Why a read fell through to Graph. A fixed set; these reach a log line. */
export type FallbackReason = 'forced' | 'miss' | 'stale' | 'oversize' | 'unavailable';

/** The slice of MirrorStore the read path uses. MirrorStore satisfies it. */
export interface MirrorReadStore {
  listNotebooks(): Promise<MirrorNotebook[]>;
  listSectionsUnder(parentId: string): Promise<MirrorSection[]>;
  listSectionGroupsUnder(parentId: string): Promise<MirrorSectionGroup[]>;
  listAllSections(): Promise<MirrorSection[]>;
  listAllSectionGroups(): Promise<MirrorSectionGroup[]>;
  getNotebook(notebookId: string): Promise<MirrorNotebook | null>;
  getSection(sectionId: string): Promise<MirrorSection | null>;
  getSectionGroup(groupId: string): Promise<MirrorSectionGroup | null>;
  getPage(pageId: string): Promise<MirrorPage | null>;
  getPageContent(pageId: string): Promise<MirrorPageContent | null>;
  listPagesInSection(sectionId: string, limit?: number): Promise<MirrorPage[]>;
  scanPages(scope?: { notebookId?: string; sectionId?: string }): Promise<ScanResult>;
}

/** One notebook as `list_notebooks` reports it from the mirror. */
export interface MirroredNotebook extends Notebook {
  /** Is this notebook's page content mirrored, or only its name? */
  readonly pagesMirrored: boolean;
}

/** Pages in a section, with the exact total rather than a `>= top` heuristic. */
export interface MirroredPages {
  readonly pages: PageSummary[];
  readonly total: number;
}

/** A title match from the mirror. */
export interface MirroredMatch {
  readonly pageId: string;
  readonly title: string;
  readonly lastModifiedDateTime: string;
  readonly sectionId: string;
  readonly sectionPath: string;
}

export interface MirroredSearch {
  readonly matches: MirroredMatch[];
  readonly totalMatches: number;
  readonly pagesScanned: number;
  /** True when SCAN_LIMIT was reached, so the answer is a sample and must say so. */
  readonly scanTruncated: boolean;
  readonly notebooksSearched: number;
  readonly notebooksInAccount: number;
}

export class MirrorReader {
  readonly #store: MirrorReadStore;
  readonly #blobs: MirrorBlobReader;
  readonly #maxInkBytes: number | undefined;

  constructor(store: MirrorReadStore, blobs: MirrorBlobReader, maxInkBytes?: number) {
    this.#store = store;
    this.#blobs = blobs;
    this.#maxInkBytes = maxInkBytes;
  }

  /**
   * Every notebook in the account, with whether its pages are mirrored.
   *
   * An empty collection is a miss: the sync has never run, and answering "no notebooks"
   * would be indistinguishable from an account with none.
   */
  async listNotebooks(): Promise<MirroredNotebook[] | null> {
    const notebooks = await this.#store.listNotebooks();
    if (notebooks.length === 0) return null;

    return notebooks.map((notebook) => ({
      id: notebook.id,
      displayName: notebook.displayName,
      pagesMirrored: notebook.mirrored,
    }));
  }

  /**
   * What sits directly inside one container.
   *
   * Two things are a miss. An unknown container id, obviously. And a section group whose
   * `childGroupsKnown` is false — `$expand` reaches one level of section group, so a
   * first-level group's nested groups are *absent from the tree response* rather than
   * known to be empty, and answering from that document would omit them and look
   * complete. The sweep sets the flag.
   */
  async listContainerChildren(
    kind: ContainerKind,
    containerId: string,
  ): Promise<ContainerChildren | null> {
    if (kind === 'notebooks') {
      if ((await this.#store.getNotebook(containerId)) === null) return null;
    } else {
      const group = await this.#store.getSectionGroup(containerId);
      if (group === null || !group.childGroupsKnown) return null;
    }

    const [sections, sectionGroups] = await Promise.all([
      this.#store.listSectionsUnder(containerId),
      this.#store.listSectionGroupsUnder(containerId),
    ]);

    return {
      sections: sections.map(toSection),
      sectionGroups: sectionGroups.map((group) => ({
        id: group.id,
        displayName: group.displayName,
      })),
    };
  }

  /**
   * Pages in one section, newest first.
   *
   * `total` is exact, because the mirror can count. The Graph path reports
   * `moreAvailable` from `pages.length >= top`, which is a heuristic that says "maybe" on
   * a section holding exactly `top` pages.
   */
  async listPagesInSection(sectionId: string, top: number): Promise<MirroredPages | null> {
    const section = await this.#store.getSection(sectionId);
    if (section === null || !section.mirrored) return null;

    // One more than asked for, so "there are exactly top" and "there are more" are
    // distinguishable without a second query.
    const pages = await this.#store.listPagesInSection(sectionId, top + 1);
    return { pages: pages.slice(0, top).map(toPageSummary), total: pages.length };
  }

  /**
   * One page's content: HTML trimmed at read time, ink from the bucket.
   *
   * `trimPageHtml` runs here rather than at sync time so the trimmer can change without
   * re-fetching every page from Graph. `fitInkToByteBudget` is re-applied for the same
   * reason — the stored PNG was fitted to whatever the budget was when it was rendered,
   * and it is a no-op unless that number has since shrunk.
   */
  async getPageContent(pageId: string): Promise<PageContent | null> {
    const page = await this.#store.getPage(pageId);
    if (page === null) return null;

    // `stale` is a write tool saying the stored copy is superseded; `missing` is a page
    // whose content could not be stored at all. Both are misses.
    if (page.contentState !== 'present') return null;

    const html =
      page.htmlLocation === 'gcs'
        ? await this.#blobs.getHtml(pageId)
        : (await this.#store.getPageContent(pageId))?.html ?? null;

    // The metadata says content is present and the content is not there. Rather than
    // answer with half a page, fall through and let Graph be right.
    if (html === null) return null;

    // A page with no ink recorded is `ink: null` and entirely normal — most pages are
    // typed. A page *with* ink recorded whose object has gone is a different thing, and
    // it has to be a miss for the whole page: answering `ink: null` there would say "this
    // page has no handwriting", which is a lie a model has no way to detect and which
    // silently drops the only copy of what the page says.
    let ink: InkImage | null = null;
    if (page.ink !== null) {
      ink = await this.#ink(pageId, page);
      if (ink === null) return null;
    }

    return { html: trimPageHtml(html), ink };
  }

  /**
   * Title substring search, filtered in this process.
   *
   * Firestore cannot match a substring. What makes doing it here affordable is that
   * `scanPages` projects server-side, so a scan of the mirrored pages transfers titles
   * rather than the HTML those documents would carry if `pages` and `pageContent` were
   * one collection.
   *
   * `matches` is the caller's job to bound; everything found is returned with the counts
   * beside it.
   */
  async searchTitles(
    matcher: (title: string) => boolean,
    scope: { notebookId?: string; sectionId?: string } = {},
  ): Promise<MirroredSearch | null> {
    if (scope.sectionId !== undefined) {
      const section = await this.#store.getSection(scope.sectionId);
      if (section === null || !section.mirrored) return null;
    }

    const notebooks = await this.#store.listNotebooks();
    if (notebooks.length === 0) return null;

    const scan = await this.#store.scanPages(scope);
    const matches = scan.pages.filter((page) => matcher(page.title)).map(toMatch);

    return {
      matches,
      totalMatches: matches.length,
      pagesScanned: scan.pages.length,
      scanTruncated: scan.truncated,
      notebooksSearched: notebooks.filter((notebook) => notebook.mirrored).length,
      notebooksInAccount: notebooks.length,
    };
  }

  /**
   * The stored structure in the shape `resolveSection` expects.
   *
   * Reassembled rather than stored that way, because the flat collections are what the
   * queries need and this is one read of a few hundred small documents.
   */
  async expandedTree(): Promise<ExpandedNotebook[] | null> {
    const [notebooks, groups, sections] = await Promise.all([
      this.#store.listNotebooks(),
      this.#store.listAllSectionGroups(),
      this.#store.listAllSections(),
    ]);
    if (notebooks.length === 0) return null;

    const byParent = new Map<string, MirrorSection[]>();
    for (const section of sections) {
      const list = byParent.get(section.parentId);
      if (list === undefined) byParent.set(section.parentId, [section]);
      else list.push(section);
    }

    return notebooks.map((notebook) => ({
      id: notebook.id,
      displayName: notebook.displayName,
      sections: (byParent.get(notebook.id) ?? []).map(toSection),
      sectionGroups: groups
        .filter((group) => group.parentId === notebook.id)
        .map((group) => ({
          id: group.id,
          displayName: group.displayName,
          sections: (byParent.get(group.id) ?? []).map(toSection),
        })),
    }));
  }

  /** Sections anywhere whose name contains the query, with their parents. */
  async findSectionsByName(displayName: string): Promise<SectionWithParents[] | null> {
    const [sections, groups, notebooks] = await Promise.all([
      this.#store.listAllSections(),
      this.#store.listAllSectionGroups(),
      this.#store.listNotebooks(),
    ]);
    if (sections.length === 0) return null;

    const wanted = displayName.toLowerCase();
    const notebookById = new Map(notebooks.map((n) => [n.id, n]));
    const groupById = new Map(groups.map((g) => [g.id, g]));

    return sections
      .filter((section) => section.displayName.toLowerCase().includes(wanted))
      .map((section) => {
        const notebook = notebookById.get(section.notebookId);
        const group = section.parentKind === 'sectionGroup' ? groupById.get(section.parentId) : undefined;
        return {
          id: section.id,
          displayName: section.displayName,
          parentNotebook:
            notebook === undefined ? null : { id: notebook.id, displayName: notebook.displayName },
          parentSectionGroup:
            group === undefined ? null : { id: group.id, displayName: group.displayName },
        };
      });
  }

  /** When the mirror last synced this page, for the result's `mirroredAt`. */
  async syncedAt(pageId: string): Promise<string | null> {
    const page = await this.#store.getPage(pageId);
    return page === null ? null : timestampToIso(page.contentSyncedAt);
  }

  /** The stored render, or null when the object is gone. Only called when ink is recorded. */
  async #ink(pageId: string, page: MirrorPage): Promise<InkImage | null> {
    if (page.ink === null) return null;

    const png = await this.#blobs.getInk(pageId);
    if (png === null) return null;

    const image: InkImage = {
      png,
      width: page.ink.width,
      height: page.ink.height,
      svg: '',
      strokeCount: page.ink.strokeCount,
    };

    return this.#maxInkBytes === undefined ? image : fitInkToByteBudget(image, this.#maxInkBytes);
  }
}

/**
 * The mirror behind `LookupStructure`, so `resolveSection` runs against it unchanged.
 *
 * Throws rather than answering empty when the mirror holds no structure, because
 * `resolveSection` would turn an empty tree into a `NameLookupError` listing no siblings
 * — which reads to a model as "that section does not exist" rather than "the mirror is
 * empty". The caller catches this and goes to Graph.
 */
export class MirrorStructureEmptyError extends Error {
  constructor() {
    super('The mirror holds no structure.');
    this.name = 'MirrorStructureEmptyError';
  }
}

export function mirrorLookupStructure(reader: MirrorReader): {
  getExpandedTree(): Promise<ExpandedNotebook[]>;
  findSectionsByName(displayName: string): Promise<SectionWithParents[]>;
} {
  return {
    getExpandedTree: async () => {
      const tree = await reader.expandedTree();
      if (tree === null) throw new MirrorStructureEmptyError();
      return tree;
    },
    findSectionsByName: async (displayName) => {
      const sections = await reader.findSectionsByName(displayName);
      if (sections === null) throw new MirrorStructureEmptyError();
      return sections;
    },
  };
}

/**
 * Try the mirror, fall back to Graph, and say which answered.
 *
 * The one place the branch is written. Every failure mode ends in the same fallback:
 * a caller that asked for live data, a miss, and a Firestore outage all reach Graph, and
 * only the reason in the log line distinguishes them. Refusing a tool call because a
 * cache is down would be strictly worse than the behaviour before the mirror existed.
 */
export async function readSourced<M, G = M>(
  mirror: MirrorReader | undefined,
  tool: string,
  forceGraph: boolean,
  fromMirror: (reader: MirrorReader) => Promise<M | null>,
  fromGraph: () => Promise<G>,
  mirroredAt?: (reader: MirrorReader) => Promise<string | null>,
): Promise<Sourced<M, G>> {
  if (mirror === undefined) return { data: await fromGraph(), source: 'graph' };

  if (forceGraph) {
    logEvent('mirror-read-fallback', { tool, reason: 'forced' });
    return { data: await fromGraph(), source: 'graph' };
  }

  let hit: M | null = null;
  try {
    hit = await fromMirror(mirror);
  } catch (err) {
    // Firestore unreachable, or the mirror holds no structure at all. Both are the same
    // answer: Graph is the ground truth and the mirror is an optimisation.
    logEvent('mirror-read-fallback', { tool, reason: reasonOf(err) });
    return { data: await fromGraph(), source: 'graph' };
  }

  if (hit === null) {
    logEvent('mirror-read-fallback', { tool, reason: 'miss' });
    return { data: await fromGraph(), source: 'graph' };
  }

  const stamp = mirroredAt === undefined ? null : await mirroredAt(mirror).catch(() => null);
  return { data: hit, source: 'mirror', ...(stamp === null ? {} : { mirroredAt: stamp }) };
}

function reasonOf(err: unknown): FallbackReason {
  return err instanceof MirrorStructureEmptyError ? 'miss' : 'unavailable';
}

function toSection(section: MirrorSection): Section {
  return { id: section.id, displayName: section.displayName };
}

function toPageSummary(page: MirrorPage): PageSummary {
  return {
    id: page.id,
    title: page.title,
    lastModifiedDateTime: page.lastModifiedDateTime,
  };
}

function toMatch(page: MirrorPage): MirroredMatch {
  return {
    pageId: page.id,
    title: page.title,
    lastModifiedDateTime: page.lastModifiedDateTime,
    sectionId: page.sectionId,
    sectionPath: page.sectionPath,
  };
}
