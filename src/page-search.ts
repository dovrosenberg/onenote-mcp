// Finding a page by title, across sections.
//
// There is no account-wide page list to search against. `GET /me/onenote/pages` fails
// with error 20266, "maximum sections exceeded", on an account with this many sections,
// so a title search that is not already scoped to one section has to walk the structure
// and list each section's pages. That is many requests, and the walk is therefore
// bounded twice: by how many sections it will visit, and by how long it will spend.
//
// A bound that is not reported is worse than no bound, because a truncated search that
// found nothing is indistinguishable from a complete search that found nothing. Every
// result below carries how many sections were searched, how many exist, and why the walk
// stopped, and `src/structure-tools.ts` puts that into the tool's own output.
//
// Matching is a case-insensitive substring of the page title. Page titles on this
// account are typed rather than handwritten, so `title` is a property worth matching on;
// nothing here reads page content.

import type {
  ContainerKind,
  NotebookTree,
  PageSummary,
  SectionGroupNode,
} from './graph-structure.ts';

/** How many sections one unscoped search will visit before it stops. */
export const MAX_SECTIONS_SEARCHED = 200;

/**
 * How long one unscoped search will spend listing pages, in milliseconds.
 *
 * The budget is checked before each section is fetched, not during, so a search can
 * overrun by one section's round trip. It exists so a call answers rather than sitting
 * on an MCP client's timeout with nothing to show.
 */
export const SEARCH_TIME_BUDGET_MS = 25_000;

/** How many pages are read out of each section while searching it. */
export const PAGES_PER_SECTION = 100;

/** How many sections are listed at once. Graph throttles a client that fans out wide. */
export const SECTION_CONCURRENCY = 4;

/** How many matches one search returns, most recently modified first. */
export const MAX_MATCHES = 100;

/** Why a search covered less than the whole account. */
export type SearchStop = 'section-limit' | 'time-budget';

/** One section, with the notebook and section groups it sits under. */
export interface SectionRef {
  readonly id: string;
  readonly displayName: string;
  /** `2026 / March / Daily todo` — the containers above it, then the section. */
  readonly path: string;
}

/** A page whose title matched. */
export interface PageMatch {
  readonly pageId: string;
  readonly title: string;
  readonly lastModifiedDateTime: string;
  readonly sectionId: string;
  /** Absent when the caller scoped the search to a section id, so no walk happened. */
  readonly sectionPath?: string;
}

export interface SearchResult {
  readonly matches: PageMatch[];
  /** Matches found before `MAX_MATCHES` was applied. */
  readonly totalMatches: number;
  readonly sectionsSearched: number;
  /** Sections the walk found. Equal to `sectionsSearched` on a complete search. */
  readonly sectionsFound: number;
  readonly stoppedEarly: boolean;
  readonly stoppedBecause: SearchStop | null;
}

/** The slice of `GraphStructure` a search calls; a test supplies a plain object. */
export interface SearchableStructure {
  getFullTree(): Promise<NotebookTree[]>;
  listPagesInSection(sectionId: string, top?: number): Promise<PageSummary[]>;
}

export interface SearchOptions {
  readonly maxSections?: number;
  readonly timeBudgetMs?: number;
  readonly pagesPerSection?: number;
  readonly concurrency?: number;
  readonly maxMatches?: number;
  /** Injected so the time budget is testable without waiting for it. */
  readonly now?: () => number;
}

/** `notebook` / `sectionGroup` as a caller says it, to the Graph relationship name. */
export function toContainerKind(containerType: string): ContainerKind | null {
  if (containerType === 'notebook') return 'notebooks';
  if (containerType === 'sectionGroup') return 'sectionGroups';
  return null;
}

/** Case-insensitive substring match on a page title. */
export function titleMatches(title: string, query: string): boolean {
  return title.toLocaleLowerCase().includes(query.trim().toLocaleLowerCase());
}

/**
 * Every section in the tree, depth first, with its path.
 *
 * Section groups nest, and this account nests them a month deep under each notebook, so
 * the recursion is the point: a flatten that stopped at a notebook's direct sections
 * would search almost nothing.
 */
export function flattenSections(trees: readonly NotebookTree[]): SectionRef[] {
  const refs: SectionRef[] = [];

  const walk = (
    sections: readonly { id: string; displayName: string }[],
    groups: readonly SectionGroupNode[],
    prefix: readonly string[],
  ): void => {
    for (const section of sections) {
      refs.push({
        id: section.id,
        displayName: section.displayName,
        path: [...prefix, section.displayName].join(' / '),
      });
    }
    for (const group of groups) {
      walk(group.sections, group.sectionGroups, [...prefix, group.displayName]);
    }
  };

  for (const tree of trees) {
    walk(tree.sections, tree.sectionGroups, [tree.displayName]);
  }

  return refs;
}

/**
 * Title search inside one section. Nothing is truncated silently: `PAGES_PER_SECTION`
 * bounds the read, and the caller is told how many pages were looked at.
 */
export async function searchOneSection(
  structure: SearchableStructure,
  sectionId: string,
  query: string,
  options: { pagesPerSection?: number; maxMatches?: number } = {},
): Promise<SearchResult> {
  const pagesPerSection = options.pagesPerSection ?? PAGES_PER_SECTION;
  const pages = await structure.listPagesInSection(sectionId, pagesPerSection);

  const matches = pages
    .filter((page) => titleMatches(page.title, query))
    .map((page) => toMatch(page, sectionId));

  return capped(matches, 1, 1, null, options.maxMatches ?? MAX_MATCHES);
}

/**
 * Title search across every section the account exposes, bounded and reported.
 *
 * A failure listing one section aborts the search rather than being swallowed: an
 * expired refresh token fails every section the same way, and a search that quietly
 * returned "no matches" for it would be read as an answer.
 */
export async function searchAllSections(
  structure: SearchableStructure,
  query: string,
  options: SearchOptions = {},
): Promise<SearchResult> {
  const maxSections = options.maxSections ?? MAX_SECTIONS_SEARCHED;
  const timeBudgetMs = options.timeBudgetMs ?? SEARCH_TIME_BUDGET_MS;
  const pagesPerSection = options.pagesPerSection ?? PAGES_PER_SECTION;
  const concurrency = options.concurrency ?? SECTION_CONCURRENCY;
  const maxMatches = options.maxMatches ?? MAX_MATCHES;
  const now = options.now ?? Date.now;

  const deadline = now() + timeBudgetMs;
  const sections = flattenSections(await structure.getFullTree());
  const planned = sections.slice(0, maxSections);

  const matches: PageMatch[] = [];
  let searched = 0;
  let stopped: SearchStop | null = planned.length < sections.length ? 'section-limit' : null;
  let cursor = 0;

  const worker = async (): Promise<void> => {
    for (;;) {
      const ref = planned[cursor];
      if (ref === undefined) return;
      // The budget is checked before the fetch is issued, so an overrun costs one
      // section's round trip rather than the whole remaining walk.
      if (now() >= deadline) {
        // Overwrites 'section-limit' on purpose: the walk stopped at the earlier of the
        // two bounds, and the time budget is what the caller can act on by scoping.
        stopped = 'time-budget';
        return;
      }
      cursor += 1;

      const pages = await structure.listPagesInSection(ref.id, pagesPerSection);
      for (const page of pages) {
        if (titleMatches(page.title, query)) matches.push(toMatch(page, ref.id, ref.path));
      }
      searched += 1;
    }
  };

  await Promise.all(
    Array.from({ length: Math.min(concurrency, planned.length) }, () => worker()),
  );

  return capped(matches, searched, sections.length, stopped, maxMatches);
}

function toMatch(page: PageSummary, sectionId: string, sectionPath?: string): PageMatch {
  return {
    pageId: page.id,
    title: page.title,
    lastModifiedDateTime: page.lastModifiedDateTime,
    sectionId,
    // Spread rather than assigned: exactOptionalPropertyTypes rejects an explicit
    // undefined for an optional member.
    ...(sectionPath === undefined ? {} : { sectionPath }),
  };
}

/** Sort newest first, then apply the match cap, keeping the counts honest. */
function capped(
  matches: PageMatch[],
  sectionsSearched: number,
  sectionsFound: number,
  stoppedBecause: SearchStop | null,
  maxMatches: number,
): SearchResult {
  const sorted = [...matches].sort((a, b) =>
    b.lastModifiedDateTime.localeCompare(a.lastModifiedDateTime),
  );

  return {
    matches: sorted.slice(0, maxMatches),
    totalMatches: sorted.length,
    sectionsSearched,
    sectionsFound,
    stoppedEarly: sectionsSearched < sectionsFound,
    stoppedBecause,
  };
}
