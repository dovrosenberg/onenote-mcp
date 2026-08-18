// The four browsing tools: list_notebooks, list_sections, list_pages, search_pages.
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
// Every result is JSON in a single text block. Nothing here logs a notebook, section, or
// page name: the names are the answer and go to the caller, and src/logging.ts records
// only the tool name.

import type { CallToolResult, Tool } from '@modelcontextprotocol/sdk/types.js';

import type {
  ContainerChildren,
  ContainerKind,
  Notebook,
  NotebookTree,
  PageSummary,
} from './graph-structure.ts';
import {
  MAX_SECTIONS_SEARCHED,
  SEARCH_TIME_BUDGET_MS,
  searchAllSections,
  searchOneSection,
  toContainerKind,
  type SearchOptions,
  type SearchResult,
} from './page-search.ts';
import {
  ToolInputError,
  optionalInteger,
  requiredString,
  optionalString,
  type ToolDefinition,
} from './mcp-tools.ts';

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
}

/**
 * Build the browsing tools over one structure client.
 *
 * `searchOptions` exists so a test can shrink the search bounds; production passes
 * nothing and gets the constants in src/page-search.ts.
 */
export function createStructureTools(
  structure: StructureClient,
  searchOptions: SearchOptions = {},
): ToolDefinition[] {
  // The description quotes the bounds a caller will actually get, so a shrunk bound in a
  // test cannot leave the tool advertising the production numbers.
  const maxSections = searchOptions.maxSections ?? MAX_SECTIONS_SEARCHED;
  const budgetSeconds = Math.round((searchOptions.timeBudgetMs ?? SEARCH_TIME_BUDGET_MS) / 1000);

  return [
    {
      name: 'list_notebooks',
      title: 'List notebooks',
      description:
        'List every OneNote notebook the signed-in account can see. Start here when you ' +
        'do not already hold a notebook, section group, or section id. Each notebook ' +
        'comes back with an id and a display name; pass the id to list_sections with ' +
        "containerType 'notebook' to see what is inside it.",
      inputSchema: { type: 'object', properties: {}, additionalProperties: false },
      annotations: READ_ONLY,
      handle: async () => {
        const notebooks = await structure.listNotebooks();
        return jsonResult({
          notebooks: notebooks.map((notebook) => ({
            id: notebook.id,
            displayName: notebook.displayName,
          })),
          count: notebooks.length,
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

        const children = await structure.listContainerChildren(kind, containerId);
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
        },
        required: ['sectionId'],
        additionalProperties: false,
      },
      annotations: READ_ONLY,
      handle: async (args) => {
        const sectionId = requiredString(args, 'sectionId');
        const top = optionalInteger(args, 'top', TOP_RANGE) ?? DEFAULT_TOP;

        const pages = await structure.listPagesInSection(sectionId, top);
        return jsonResult({
          sectionId,
          pages: pages.map((page) => ({
            id: page.id,
            title: page.title,
            lastModifiedDateTime: page.lastModifiedDateTime,
          })),
          count: pages.length,
          top,
          // Graph was asked for `top` and gave exactly that many, so the section holds
          // at least one more. Saying so is cheaper than a second call to find out.
          moreAvailable: pages.length >= top,
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
              'Restrict the search to one section. Strongly preferred; omitting it walks the account.',
          },
        },
        required: ['query'],
        additionalProperties: false,
      },
      annotations: READ_ONLY,
      handle: async (args) => {
        const query = requiredString(args, 'query');
        const sectionId = optionalString(args, 'sectionId');

        const result =
          sectionId === undefined
            ? await searchAllSections(structure, query, searchOptions)
            : await searchOneSection(structure, sectionId, query, searchOptions);

        return jsonResult({
          query,
          scope: sectionId === undefined ? 'account' : 'section',
          ...(sectionId === undefined ? {} : { sectionId }),
          matches: result.matches,
          ...searchCounts(result),
          note: searchNote(result),
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

/** Every tool here answers with one JSON text block. */
function jsonResult(payload: unknown): CallToolResult {
  return { content: [{ type: 'text', text: JSON.stringify(payload, null, 2) }] };
}
