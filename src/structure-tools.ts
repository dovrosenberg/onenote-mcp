// The browsing tools: list_notebooks, list_sections, list_pages, search_pages, and the
// two name-based lookups, find_page_by_name and list_pages_by_name.
//
// These are what a calling model uses to get from "somewhere in this account" to a page
// id, which is the only thing the reading and writing tools take. The descriptions below
// are written for that model rather than for a human reading source: each one says what
// the tool returns, what id the next call needs, and what a bounded result means.
//
// Two shapes here are not arbitrary:
//
// `list_sections` returns sections and section groups in one result, each tagged with
// its type. Section groups are OneNote's "tab groups" and they nest — this account has a
// notebook per year and a section group per month — so a caller that got only sections
// back would have to guess that a second relationship exists to recurse into.
//
// `search_pages` never calls the account-wide page list; see src/page-search.ts for what
// it does instead and why the walk is bounded.
//
// The two `_by_name` tools exist because everything else here takes an id, so a caller
// that already knows the names pays three round trips to convert them. They resolve the
// whole path in one Graph request; the rules for matching are in src/name-lookup.ts.
//
// Every result is JSON in a single text block. Nothing here logs a notebook, section, or
// page name: the names are the answer and go to the caller, and src/logging.ts records
// only the tool name.

import type { CallToolResult, Tool } from '@modelcontextprotocol/sdk/types.js';

import type {
  ContainerChildren,
  ContainerKind,
  ExpandedNotebook,
  Notebook,
  NotebookTree,
  PageSummary,
  SectionWithParents,
} from './graph-structure.ts';
import {
  namesMatch,
  resolveSection,
  resolvedPayload,
  type NamePath,
  type ResolvedPath,
} from './name-lookup.ts';
import {
  MAX_MATCHES,
  MAX_SECTIONS_SEARCHED,
  SEARCH_TIME_BUDGET_MS,
  searchAllSections,
  searchOneSection,
  titleMatches,
  toContainerKind,
  type SearchOptions,
  type SearchResult,
} from './page-search.ts';
import {
  ToolInputError,
  optionalBoolean,
  optionalInteger,
  requiredString,
  optionalString,
  type ToolDefinition,
} from './mcp-tools.ts';
import {
  USE_LIVE_DATA_PROPERTY,
  mirrorLookupStructure,
  readSourced,
  type MirroredSearch,
  type MirrorReader,
  type Sourced,
} from './mirror-reader.ts';

/** Bounds on `list_pages`'s `top`. Graph decides its own page size regardless. */
const TOP_RANGE = { min: 1, max: 100 };

/** What `list_pages` returns when the caller gives no `top`. */
const DEFAULT_TOP = 50;

/** Every tool here reads; none of them writes. */
const READ_ONLY: Tool['annotations'] = { readOnlyHint: true, openWorldHint: true };

/** The slice of `GraphStructure` these tools call, so a test can pass a plain object. */
export interface StructureClient {
  listNotebooks(): Promise<Notebook[]>;
  listContainerChildren(kind: ContainerKind, containerId: string): Promise<ContainerChildren>;
  listPagesInSection(sectionId: string, top?: number): Promise<PageSummary[]>;
  getFullTree(): Promise<NotebookTree[]>;
  getExpandedTree(): Promise<ExpandedNotebook[]>;
  findSectionsByName(displayName: string): Promise<SectionWithParents[]>;
  findPagesByTitle(sectionId: string, title: string): Promise<PageSummary[]>;
  findPagesMatchingTitle(sectionId: string, query: string): Promise<PageSummary[]>;
}

/** The name arguments both `_by_name` tools share. */
const NAME_PATH_PROPERTIES = {
  notebookName: {
    type: 'string',
    description: 'The notebook name, matched in full and case-insensitively.',
  },
  sectionGroupName: {
    type: 'string',
    description:
      'The section group holding the section, if it is in one. Omit when the section ' +
      'sits directly in the notebook; omitting it does not search inside groups.',
  },
  sectionName: {
    type: 'string',
    description: 'The section name, matched in full and case-insensitively.',
  },
} as const;

/**
 * Build the browsing tools over one structure client.
 *
 * `searchOptions` exists so a test can shrink the search bounds; production passes
 * nothing and gets the constants in src/page-search.ts.
 */
export function createStructureTools(
  structure: StructureClient,
  searchOptions: SearchOptions = {},
  mirror?: MirrorReader,
): ToolDefinition[] {
  // The description quotes the bounds a caller will actually get, so a shrunk bound in a
  // test cannot leave the tool advertising the production numbers.
  const maxSections = searchOptions.maxSections ?? MAX_SECTIONS_SEARCHED;
  const budgetSeconds = Math.round((searchOptions.timeBudgetMs ?? SEARCH_TIME_BUDGET_MS) / 1000);
  const maxMatches = searchOptions.maxMatches ?? MAX_MATCHES;

  return [
    {
      name: 'list_notebooks',
      title: 'List notebooks',
      description:
        'List every OneNote notebook the signed-in account can see. Start here when you ' +
        'do not already hold a notebook, section group, or section id. Each notebook ' +
        'comes back with an id and a display name; pass the id to list_sections with ' +
        "containerType 'notebook' to see what is inside it.",
      inputSchema: {
        type: 'object',
        properties: { useLiveData: USE_LIVE_DATA_PROPERTY },
        additionalProperties: false,
      },
      annotations: READ_ONLY,
      handle: async (args) => {
        // The mirror stores every notebook in the account, not only the ones whose pages
        // are mirrored, precisely so this tool cannot answer partially. `pagesMirrored`
        // says which are which.
        const answer = await readSourced(
          mirror,
          'list_notebooks',
          optionalBoolean(args, 'useLiveData') ?? false,
          (reader) => reader.listNotebooks(),
          async () =>
            (await structure.listNotebooks()).map((notebook) => ({
              id: notebook.id,
              displayName: notebook.displayName,
              pagesMirrored: false,
            })),
        );

        return jsonResult({
          notebooks: answer.data.map((notebook) =>
            answer.source === 'mirror'
              ? notebook
              : { id: notebook.id, displayName: notebook.displayName },
          ),
          count: answer.data.length,
          ...sourcePayload(answer),
        });
      },
    },

    {
      name: 'list_sections',
      title: 'List sections and section groups',
      description:
        'List what sits directly inside one notebook or one section group: its sections ' +
        'and its section groups together, each tagged with its type. Section groups are ' +
        "OneNote's tab groups and they nest, so a returned item of type 'sectionGroup' " +
        "can be passed straight back in with containerType 'sectionGroup' to descend a " +
        'level. Only a section holds pages — take an item of type \'section\' to ' +
        'list_pages. This account is organised as a notebook per year with a section ' +
        'group per month, so expect to descend at least once.',
      inputSchema: {
        type: 'object',
        properties: {
          containerType: {
            type: 'string',
            enum: ['notebook', 'sectionGroup'],
            description: 'What containerId identifies.',
          },
          containerId: {
            type: 'string',
            description: 'A notebook id from list_notebooks, or a section group id from list_sections.',
          },
          useLiveData: USE_LIVE_DATA_PROPERTY,
        },
        required: ['containerType', 'containerId'],
        additionalProperties: false,
      },
      annotations: READ_ONLY,
      handle: async (args) => {
        const containerType = requiredString(args, 'containerType');
        const kind = toContainerKind(containerType);
        if (kind === null) {
          throw new ToolInputError('containerType', "must be 'notebook' or 'sectionGroup'");
        }
        const containerId = requiredString(args, 'containerId');

        const answer = await readSourced(
          mirror,
          'list_sections',
          optionalBoolean(args, 'useLiveData') ?? false,
          (reader) => reader.listContainerChildren(kind, containerId),
          () => structure.listContainerChildren(kind, containerId),
        );
        const children = answer.data;

        return jsonResult({
          containerType,
          containerId,
          // One list rather than two, so a caller can recurse over it without knowing
          // that Graph exposes the two child relationships separately.
          children: [
            ...children.sections.map((section) => ({
              type: 'section' as const,
              id: section.id,
              displayName: section.displayName,
            })),
            ...children.sectionGroups.map((group) => ({
              type: 'sectionGroup' as const,
              id: group.id,
              displayName: group.displayName,
            })),
          ],
          sectionCount: children.sections.length,
          sectionGroupCount: children.sectionGroups.length,
          ...sourcePayload(answer),
        });
      },
    },

    {
      name: 'list_pages',
      title: 'List pages in a section',
      description:
        'List the pages in one section, most recently modified first. sectionId must be ' +
        "an id of type 'section' from list_sections; a section group id will fail. top " +
        `bounds how many pages come back (${TOP_RANGE.min}-${TOP_RANGE.max}, default ` +
        `${DEFAULT_TOP}); when exactly top pages are returned, moreAvailable is true and ` +
        'older pages exist beyond them. Pass a page id to get_page_content to read one.',
      inputSchema: {
        type: 'object',
        properties: {
          sectionId: { type: 'string', description: 'A section id from list_sections.' },
          top: {
            type: 'integer',
            minimum: TOP_RANGE.min,
            maximum: TOP_RANGE.max,
            description: `How many pages to return. Defaults to ${DEFAULT_TOP}.`,
          },
          useLiveData: USE_LIVE_DATA_PROPERTY,
        },
        required: ['sectionId'],
        additionalProperties: false,
      },
      annotations: READ_ONLY,
      handle: async (args) => {
        const sectionId = requiredString(args, 'sectionId');
        const top = optionalInteger(args, 'top', TOP_RANGE) ?? DEFAULT_TOP;

        const answer = await readSourced(
          mirror,
          'list_pages',
          optionalBoolean(args, 'useLiveData') ?? false,
          (reader) => reader.listPagesInSection(sectionId, top),
          () => structure.listPagesInSection(sectionId, top),
        );

        // The two paths mean different things by "more available", and collapsing them
        // into one number gets one of them wrong. Graph was asked for `top` and can
        // never return more, so exactly `top` back means "at least one more" — a
        // heuristic that says maybe on a section holding exactly `top`. The mirror asked
        // for one extra and counted, so its answer is a fact.
        const pages = answer.source === 'mirror' ? answer.data.pages : answer.data;
        const moreAvailable =
          answer.source === 'mirror' ? answer.data.total > top : answer.data.length >= top;

        return jsonResult({
          sectionId,
          pages: pages.map((page) => ({
            id: page.id,
            title: page.title,
            lastModifiedDateTime: page.lastModifiedDateTime,
          })),
          count: pages.length,
          top,
          moreAvailable,
          ...sourcePayload(answer),
        });
      },
    },

    {
      name: 'search_pages',
      title: 'Search page titles',
      description:
        'Find pages whose title contains query, matched case-insensitively as a ' +
        'substring. Page bodies are not searched, only titles. Pass sectionId whenever ' +
        'you know which section to look in: without it the search walks every notebook ' +
        'and section group in the account and lists the pages of each section it finds, ' +
        `which costs many requests and is bounded at ${maxSections} sections and about ` +
        `${budgetSeconds} seconds. The result ` +
        'always reports sectionsSearched, sectionsFound, and stoppedEarly: when ' +
        'stoppedEarly is true, no match found is not the same as no such page, and the ' +
        'search should be narrowed with sectionId rather than repeated.',
      inputSchema: {
        type: 'object',
        properties: {
          query: {
            type: 'string',
            description: 'Text to look for in page titles, case-insensitive substring.',
          },
          sectionId: {
            type: 'string',
            description:
              'Restrict the search to one section. Strongly preferred when answering from OneNote directly; omitting it walks the account.',
          },
          useLiveData: USE_LIVE_DATA_PROPERTY,
        },
        required: ['query'],
        additionalProperties: false,
      },
      annotations: READ_ONLY,
      handle: async (args) => {
        const query = requiredString(args, 'query');
        const sectionId = optionalString(args, 'sectionId');

        const answer = await readSourced(
          mirror,
          'search_pages',
          optionalBoolean(args, 'useLiveData') ?? false,
          (reader) =>
            reader.searchTitles(
              (title) => titleMatches(title, query),
              sectionId === undefined ? {} : { sectionId },
            ),
          () =>
            sectionId === undefined
              ? searchAllSections(structure, query, searchOptions)
              : searchOneSection(structure, sectionId, query, searchOptions),
        );

        const common = {
          query,
          scope: sectionId === undefined ? 'account' : 'section',
          ...(sectionId === undefined ? {} : { sectionId }),
          ...sourcePayload(answer),
        };

        if (answer.source === 'graph') {
          const result = answer.data;
          return jsonResult({
            ...common,
            matches: result.matches,
            ...searchCounts(result),
            note: searchNote(result),
          });
        }

        // From the mirror there is no walk and no section bound, so `stoppedEarly` and
        // the section counts mean nothing. What *is* bounded is the scope: page content
        // is mirrored only for the selected notebooks, so a search that covered fewer
        // notebooks than the account holds has to say so or a model reads a partial
        // answer as a complete one.
        const found = answer.data;
        const matches = found.matches.slice(0, maxMatches);
        return jsonResult({
          ...common,
          matches,
          totalMatches: found.totalMatches,
          returned: matches.length,
          pagesScanned: found.pagesScanned,
          scanTruncated: found.scanTruncated,
          notebooksSearched: found.notebooksSearched,
          notebooksInAccount: found.notebooksInAccount,
          note: mirrorSearchNote(found, matches.length),
        });
      },
    },

    {
      name: 'find_page_by_name',
      title: 'Find a page by name',
      description:
        'Find a page when you already know where it lives: the notebook, the section ' +
        'group if it is in one, the section, and the page title. Saves the ' +
        'list_notebooks -> list_sections -> list_pages walk that reaching the same page ' +
        'by id would cost. Container names are matched in full and case-insensitively ' +
        'first, then against the name with any leading number stripped, so ' +
        "'February' finds a section group named '062 - February'; a substring is tried " +
        'last. matchedBy in the result says which of those matched. The page title is ' +
        'matched in full, ignoring case — use search_pages when you only remember part ' +
        'of one, or list_pages_by_name to see every title in the section. Omit ' +
        'sectionGroupName when the section sits directly in the notebook; it is not a ' +
        'wildcard. A name that matches nothing, or matches more than one thing, comes ' +
        'back as an error listing the candidates. Returns the page id to pass to ' +
        'get_page_content.',
      inputSchema: {
        type: 'object',
        properties: {
          ...NAME_PATH_PROPERTIES,
          pageTitle: {
            type: 'string',
            description: 'The page title, matched in full and case-insensitively.',
          },
          useLiveData: USE_LIVE_DATA_PROPERTY,
        },
        required: ['notebookName', 'sectionName', 'pageTitle'],
        additionalProperties: false,
      },
      annotations: READ_ONLY,
      handle: async (args) => {
        const path = namePath(args);
        const pageTitle = requiredString(args, 'pageTitle');

        const answer = await resolveAndList(
          mirror,
          structure,
          'find_page_by_name',
          optionalBoolean(args, 'useLiveData') ?? false,
          path,
          // Graph does the title comparison, case-insensitively, so no bound on how many
          // pages the section holds can hide a match. The mirror applies the same rule
          // itself, over every page it holds for the section.
          (sectionId) => structure.findPagesByTitle(sectionId, pageTitle),
          (reader, sectionId) =>
            reader.listPagesInSection(sectionId, Number.MAX_SAFE_INTEGER).then((held) =>
              held === null
                ? null
                : held.pages.filter((page) => namesMatch(page.title, pageTitle)),
            ),
        );
        return jsonResult({
          ...resolvedPayload(answer.resolved),
          pageTitle,
          matches: answer.pages.map(pagePayload),
          matchCount: answer.pages.length,
          note: findNote(answer.resolved, answer.pages.length),
          ...sourcePayload(answer.sourced),
        });
      },
    },

    {
      name: 'list_pages_by_name',
      title: 'List a section\'s pages by name',
      description:
        'List the pages in a section you can name, most recently modified first, as ' +
        'titles with their page ids. Use this when you know the section but not which ' +
        'page you want: read the titles, pick one, and pass its id straight to ' +
        'get_page_content — no further lookup is needed. Takes the notebook name, the ' +
        'section group name if the section is in one, and the section name. Names are ' +
        "matched in full ignoring case, then with any leading number stripped ('February' " +
        "finds '062 - February'), then as a substring. This is list_pages without the " +
        'two calls it would take to turn those names into a section id. top ' +
        `bounds the result (${TOP_RANGE.min}-${TOP_RANGE.max}, default ${DEFAULT_TOP}) ` +
        'and moreAvailable reports whether older pages exist beyond it.',
      inputSchema: {
        type: 'object',
        properties: {
          ...NAME_PATH_PROPERTIES,
          top: {
            type: 'integer',
            minimum: TOP_RANGE.min,
            maximum: TOP_RANGE.max,
            description: `How many pages to return. Defaults to ${DEFAULT_TOP}.`,
          },
          useLiveData: USE_LIVE_DATA_PROPERTY,
        },
        required: ['notebookName', 'sectionName'],
        additionalProperties: false,
      },
      annotations: READ_ONLY,
      handle: async (args) => {
        const path = namePath(args);
        const top = optionalInteger(args, 'top', TOP_RANGE) ?? DEFAULT_TOP;

        const answer = await resolveAndList(
          mirror,
          structure,
          'list_pages_by_name',
          optionalBoolean(args, 'useLiveData') ?? false,
          path,
          (sectionId) => structure.listPagesInSection(sectionId, top),
          (reader, sectionId) =>
            reader.listPagesInSection(sectionId, top).then((held) => held?.pages ?? null),
        );
        return jsonResult({
          ...resolvedPayload(answer.resolved),
          pages: answer.pages.map(pagePayload),
          count: answer.pages.length,
          top,
          moreAvailable: answer.pages.length >= top,
          ...sourcePayload(answer.sourced),
        });
      },
    },

  ];
}

function searchCounts(result: SearchResult): Record<string, unknown> {
  return {
    matchCount: result.matches.length,
    totalMatches: result.totalMatches,
    sectionsSearched: result.sectionsSearched,
    sectionsFound: result.sectionsFound,
    stoppedEarly: result.stoppedEarly,
    stoppedBecause: result.stoppedBecause,
  };
}

/**
 * One sentence saying what the numbers mean. The counts above are the contract; this is
 * there because a truncated search that found nothing reads as an answer without it.
 */
function searchNote(result: SearchResult): string {
  const parts: string[] = [];

  if (result.stoppedEarly) {
    const because =
      result.stoppedBecause === 'time-budget'
        ? 'the time budget ran out'
        : 'the section limit was reached';
    parts.push(
      `Searched ${result.sectionsSearched} of ${result.sectionsFound} sections and stopped because ${because}, so the sections not searched may hold matches. Narrow the search with sectionId.`,
    );
  } else {
    parts.push(`Searched all ${result.sectionsSearched} section(s) in scope.`);
  }

  if (result.totalMatches > result.matches.length) {
    parts.push(
      `${result.totalMatches} titles matched and the ${result.matches.length} most recently modified are returned.`,
    );
  }

  return parts.join(' ');
}

/** The three name arguments, read the same way by both `_by_name` tools. */
function namePath(args: Readonly<Record<string, unknown>>): {
  notebookName: string;
  sectionGroupName: string | undefined;
  sectionName: string;
} {
  return {
    notebookName: requiredString(args, 'notebookName'),
    sectionGroupName: optionalString(args, 'sectionGroupName'),
    sectionName: requiredString(args, 'sectionName'),
  };
}

function pagePayload(page: PageSummary): Record<string, unknown> {
  return { id: page.id, title: page.title, lastModifiedDateTime: page.lastModifiedDateTime };
}

/** What the counts mean, and only when they need saying. */
function findNote(resolved: ResolvedPath, matchCount: number): string {
  const parts: string[] = [];

  if (matchCount === 0) {
    // Graph matched the title across the whole section, so this is a complete answer
    // rather than a bounded one: no page in that section carries that title.
    parts.push(
      'No page in that section has that title. Titles are matched in full, ignoring case — use search_pages to match part of one.',
    );
  }

  if (resolved.deepSearchUsed) {
    parts.push(
      'The section sits below the section group named, and was found by an account-wide search on its name.',
    );
  }

  return parts.join(' ');
}

/** Every tool here answers with one JSON text block. */
function jsonResult(payload: unknown): CallToolResult {
  return { content: [{ type: 'text', text: JSON.stringify(payload, null, 2) }] };
}

/**
 * What a mirror-answered search has to say about its own scope.
 *
 * The Graph path's `stoppedEarly` warns that a walk was cut short. Here nothing is cut
 * short — but page content is mirrored only for the selected notebooks, so the *scope* is
 * narrower than the account even when the search itself was complete. A model that read
 * "no matches" without knowing that would conclude the page does not exist.
 */
function mirrorSearchNote(found: MirroredSearch, returned: number): string {
  const parts: string[] = [];

  if (found.notebooksSearched < found.notebooksInAccount) {
    const unsearched = found.notebooksInAccount - found.notebooksSearched;
    parts.push(
      `Only ${found.notebooksSearched} of ${found.notebooksInAccount} notebooks have their ` +
        `pages held locally, so pages in the other ${unsearched} were not searched. Pass ` +
        'useLiveData true to search OneNote directly.',
    );
  } else {
    parts.push('Every notebook in the account was searched.');
  }

  if (found.scanTruncated) {
    parts.push(
      `The scan stopped after ${found.pagesScanned} pages, so no match found is not the ` +
        'same as no such page. Narrow it with sectionId.',
    );
  }

  if (returned < found.totalMatches) {
    parts.push(`${found.totalMatches} titles matched; the ${returned} most recent are shown.`);
  }

  return parts.join(' ');
}

/** `source` and `mirroredAt`, spread into every covered tool's result. */
function sourcePayload(answer: Sourced<unknown, unknown>): Record<string, unknown> {
  return {
    source: answer.source,
    ...(answer.mirroredAt === undefined ? {} : { mirroredAt: answer.mirroredAt }),
  };
}

/**
 * The shared body of the two `_by_name` reading tools: resolve, then list.
 *
 * **A `NameLookupError` from the mirror is retried against Graph rather than reported.**
 * The mirror's structure equals Graph's expanded tree by construction, but it is only as
 * fresh as the last sync — a section created ten minutes ago is absent, and
 * `resolveSection` would turn that into "sectionName matched nothing" listing the wrong
 * siblings. That reads to a model as "no such section", which is a worse answer than a
 * slower one. The retry costs one Graph request and only on a failure.
 *
 * The whole resolve-and-list pair goes to one side or the other. Resolving against the
 * mirror and then listing against Graph would work, but it would report a `source` that
 * described half the answer.
 */
async function resolveAndList(
  mirror: MirrorReader | undefined,
  structure: StructureClient,
  tool: string,
  forceGraph: boolean,
  path: NamePath,
  listFromGraph: (sectionId: string) => Promise<PageSummary[]>,
  listFromMirror: (reader: MirrorReader, sectionId: string) => Promise<PageSummary[] | null>,
): Promise<{ resolved: ResolvedPath; pages: PageSummary[]; sourced: Sourced<unknown, unknown> }> {
  const viaGraph = async (): Promise<{ resolved: ResolvedPath; pages: PageSummary[] }> => {
    const resolved = await resolveSection(structure, path);
    return { resolved, pages: await listFromGraph(resolved.section.id) };
  };

  const sourced = await readSourced(
    mirror,
    tool,
    forceGraph,
    async (reader) => {
      // A NameLookupError propagates out of readSourced's try, which turns it into a
      // fallback rather than an answer. That is the retry.
      const resolved = await resolveSection(mirrorLookupStructure(reader), path);
      const pages = await listFromMirror(reader, resolved.section.id);
      return pages === null ? null : { resolved, pages };
    },
    viaGraph,
  );

  return { resolved: sourced.data.resolved, pages: sourced.data.pages, sourced };
}
