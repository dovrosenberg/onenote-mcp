// Finding a page by title, across sections.
//
// There is no account-wide page list to search against. `GET /me/onenote/pages` fails
// with error 20266, "maximum sections exceeded", on an account with this many sections,
// so a title search that is not already scoped to one section has to visit each section
// separately. That is one request per section, and the walk is therefore bounded twice:
// by how many sections it will visit, and by how long it will spend.
//
// Two things make it cheaper than it was. The section list comes from `getExpandedTree`,
// one request rather than the 195 the per-container walk cost. And the title comparison
// is done by Graph, with `contains(tolower(title), '…')`, so a section returns its
// matches rather than its first hundred pages — nothing is read and discarded, and no
// per-section page bound can hide a match.
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
  ExpandedNotebook,
  ExpandedSectionGroup,
  PageSummary,
} from './graph-structure.ts';

/**
 * How many sections one unscoped search will visit before it stops.
 *
 * Sized against the request budget rather than against the account. OneNote allows 120
 * requests a minute and 400 an hour; at the paced rate this service runs at, 60 sections
 * is what fits inside the time budget below, and it leaves most of the hourly budget for
 * everything else. The account holds 560 sections, so an unscoped search is a sample and
 * says so — `stoppedEarly` and the counts are in every result.
 */
export const MAX_SECTIONS_SEARCHED = 60;

/**
 * How long one unscoped search will spend listing pages, in milliseconds.
 *
 * The budget is checked before each section is fetched, not during, so a search can
 * overrun by one section's round trip. It exists so a call answers rather than sitting
 * on an MCP client's timeout with nothing to show.
 */
export const SEARCH_TIME_BUDGET_MS = 25_000;

/**
 * How many matches are read out of each section.
 *
 * This bounds the matches, not the pages: Graph applies the title filter, so a section
 * with 500 pages and two matching titles returns two.
 */
export const MATCHES_PER_SECTION = 100;

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
  getExpandedTree(): Promise<ExpandedNotebook[]>;
  findPagesMatchingTitle(sectionId: string, query: string): Promise<PageSummary[]>;
}

export interface SearchOptions {
  readonly maxSections?: number;
  readonly timeBudgetMs?: number;
  readonly matchesPerSection?: number;
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
 * Every section in the expanded tree, with the path that names it.
 *
 * The tree carries a notebook's own sections and its section groups' sections, which is
 * as deep as Graph's `$expand` reaches. A section nested below that is not searched and
 * not counted — `find_page_by_name` is the tool that reaches one, and this one says how
 * many sections it saw so a caller is not told "no such page" by a partial walk.
 */
export function flattenSections(trees: readonly ExpandedNotebook[]): SectionRef[] {
  const refs: SectionRef[] = [];

  const add = (
    section: { id: string; displayName: string },
    prefix: readonly string[],
  ): void => {
    refs.push({
      id: section.id,
      displayName: section.displayName,
      path: [...prefix, section.displayName].join(' / '),
    });
  };

  for (const tree of trees) {
    for (const section of tree.sections) add(section, [tree.displayName]);
    for (const group of tree.sectionGroups as readonly ExpandedSectionGroup[]) {
      for (const section of group.sections) add(section, [tree.displayName, group.displayName]);
    }
  }

  return refs;
}

/**
 * Title search inside one section: one request, and Graph does the matching.
 *
 * `titleMatches` is not applied again here. The filter Graph runs is the same rule —
 * a case-insensitive substring of the title — and re-checking it locally would only
 * disagree with the service over Unicode casing.
 */
export async function searchOneSection(
  structure: SearchableStructure,
  sectionId: string,
  query: string,
  options: { matchesPerSection?: number; maxMatches?: number } = {},
): Promise<SearchResult> {
  const pages = await structure.findPagesMatchingTitle(sectionId, query);
  const matches = pages
    .slice(0, options.matchesPerSection ?? MATCHES_PER_SECTION)
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
  const matchesPerSection = options.matchesPerSection ?? MATCHES_PER_SECTION;
  const concurrency = options.concurrency ?? SECTION_CONCURRENCY;
  const maxMatches = options.maxMatches ?? MAX_MATCHES;
  const now = options.now ?? Date.now;

  const deadline = now() + timeBudgetMs;
  // One request for every section in the account, rather than one per container.
  const sections = flattenSections(await structure.getExpandedTree());
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

      const pages = await structure.findPagesMatchingTitle(ref.id, query);
      for (const page of pages.slice(0, matchesPerSection)) {
        matches.push(toMatch(page, ref.id, ref.path));
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
