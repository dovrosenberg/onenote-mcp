// Turning names into ids: notebook, optional section group, section.
//
// The browsing tools in ./structure-tools.ts take ids, so a caller that already knows
// the names spends three or four round trips getting to a page. This resolves the whole
// path in one Graph request, using the expanded tree described in the `Graph request
// budget` section of CLAUDE.md.
//
// Three rules the tools built on this depend on:
//
// Matching is exact and case-insensitive, on every name. `"monthly log"` finds
// `"Monthly Log"` and `"Monthly"` finds nothing. A caller that only half-remembers a
// name has `search_pages`, which matches substrings; a lookup that silently accepted a
// prefix would return a different section than the caller named.
//
// Ambiguity is reported, never resolved. Two sections with the same name in one notebook
// are a real possibility, and picking the first would give a confidently wrong answer.
//
// `sectionGroupName` omitted means the section is a direct child of the notebook, not
// "look anywhere". A caller that does not know where the section sits should browse.

import type { ExpandedNotebook, Section, SectionGroup } from './graph-structure.ts';

/** The slice of `GraphStructure` this module calls, so a test can pass a plain object. */
export interface LookupStructure {
  getExpandedTree(): Promise<ExpandedNotebook[]>;
  listContainerChildren(
    kind: 'notebooks' | 'sectionGroups',
    containerId: string,
  ): Promise<{ sections: Section[]; sectionGroups: SectionGroup[] }>;
}

/** What the caller named. `sectionGroupName` absent means a notebook-level section. */
export interface NamePath {
  readonly notebookName: string;
  readonly sectionGroupName?: string | undefined;
  readonly sectionName: string;
}

/** One resolved container in the path, with the name as Graph spells it. */
export interface ResolvedNode {
  readonly id: string;
  readonly displayName: string;
}

/** A path resolved all the way to a section. */
export interface ResolvedPath {
  readonly notebook: ResolvedNode;
  readonly sectionGroup: ResolvedNode | null;
  readonly section: ResolvedNode;
  /** True when the section was found only by walking past the expanded tree. */
  readonly deepSearchUsed: boolean;
}

/** Which name failed, and what was there instead. */
export type LookupFailureKind = 'not-found' | 'ambiguous';

/**
 * A name that matched nothing, or matched more than once.
 *
 * This is an error rather than an empty result on purpose: a caller that named a section
 * and got back an empty page list cannot tell "that section is empty" from "there is no
 * such section", and the second is the one it has to act on.
 */
export class NameLookupError extends Error {
  readonly kind: LookupFailureKind;
  /** Which argument failed: `notebookName`, `sectionGroupName`, or `sectionName`. */
  readonly argument: string;
  /** The names that were there, or the ambiguous candidates. */
  readonly candidates: readonly ResolvedNode[];

  constructor(
    kind: LookupFailureKind,
    argument: string,
    wanted: string,
    candidates: readonly ResolvedNode[],
  ) {
    super(buildMessage(kind, argument, wanted, candidates));
    this.name = 'NameLookupError';
    this.kind = kind;
    this.argument = argument;
    this.candidates = candidates;
  }
}

/** How many sibling names an error message lists before it stops. */
const MAX_NAMES_LISTED = 25;

/**
 * Bounds on the fallback walk below a named section group.
 *
 * The expanded tree stops at a group's own sections, so a section inside a group nested
 * under it is absent from that response. The walk that finds it costs one request per
 * container, and the OneNote budget is 400 requests an hour, so it is bounded twice: by
 * depth and by total requests. The real account has no nesting at this level, so neither
 * bound is reached on the common path.
 */
const MAX_DEEP_DEPTH = 5;
const MAX_DEEP_REQUESTS = 20;

/** Just the two fields, dropping whatever children the node arrived with. */
function plain(node: ResolvedNode): ResolvedNode {
  return { id: node.id, displayName: node.displayName };
}

/** Case-insensitive, whole-string, and insensitive to surrounding whitespace. */
export function namesMatch(a: string, b: string): boolean {
  return a.trim().toLowerCase() === b.trim().toLowerCase();
}

/**
 * The one match for `wanted` among `nodes`.
 *
 * @throws {NameLookupError} when nothing matches or more than one does.
 */
export function matchOne(
  nodes: readonly ResolvedNode[],
  wanted: string,
  argument: string,
): ResolvedNode {
  const matches = nodes.filter((node) => namesMatch(node.displayName, wanted));
  // Narrowed to id and display name: the nodes coming in are expanded notebooks and
  // section groups carrying their children, and a ResolvedPath that leaked those would
  // put the whole subtree into every tool result.
  if (matches.length === 1) return plain(matches[0] as ResolvedNode);
  if (matches.length === 0) {
    throw new NameLookupError('not-found', argument, wanted, nodes.map(plain));
  }
  throw new NameLookupError('ambiguous', argument, wanted, matches.map(plain));
}

/**
 * Resolve notebook → section group → section, in one Graph request where possible.
 *
 * The expanded tree covers a notebook's own sections and one level of section group. A
 * section that is not there may still exist deeper, so the notebook that matched — and
 * only that notebook — is walked before the lookup is called a failure. The real account
 * has no such nesting, so the walk does not run on the common path.
 *
 * @throws {NameLookupError} when a name matches nothing or matches more than once.
 */
export async function resolveSection(
  structure: LookupStructure,
  path: NamePath,
): Promise<ResolvedPath> {
  const tree = await structure.getExpandedTree();
  const notebook = matchOne(tree, path.notebookName, 'notebookName');
  const expanded = tree.find((candidate) => candidate.id === notebook.id) as ExpandedNotebook;

  if (path.sectionGroupName === undefined) {
    const direct = expanded.sections.filter((section) =>
      namesMatch(section.displayName, path.sectionName),
    );
    if (direct.length === 1) {
      return {
        notebook,
        sectionGroup: null,
        section: plain(direct[0] as ResolvedNode),
        deepSearchUsed: false,
      };
    }
    if (direct.length > 1) {
      throw new NameLookupError('ambiguous', 'sectionName', path.sectionName, direct);
    }
    // Not among the notebook's own sections. It may sit inside a section group the
    // caller did not name, which is worth saying rather than reporting a bare miss.
    throw new NameLookupError(
      'not-found',
      'sectionName',
      path.sectionName,
      expanded.sections,
    );
  }

  const group = matchOne(expanded.sectionGroups, path.sectionGroupName, 'sectionGroupName');
  const expandedGroup = expanded.sectionGroups.find((candidate) => candidate.id === group.id);
  const sections = expandedGroup?.sections ?? [];

  const matches = sections.filter((section) => namesMatch(section.displayName, path.sectionName));
  if (matches.length === 1) {
    return {
      notebook,
      sectionGroup: group,
      section: plain(matches[0] as ResolvedNode),
      deepSearchUsed: false,
    };
  }
  if (matches.length > 1) {
    throw new NameLookupError('ambiguous', 'sectionName', path.sectionName, matches);
  }

  // The expanded response stops at this group's sections, so a section inside a group
  // nested under it is absent from the response rather than known to be absent. Only
  // that subtree is walked, and only when the cheap answer came back empty.
  const deep = await findSectionBelow(structure, group.id, path.sectionName);
  if (deep.matches.length === 1) {
    return {
      notebook,
      sectionGroup: group,
      section: plain(deep.matches[0] as ResolvedNode),
      deepSearchUsed: true,
    };
  }
  if (deep.matches.length > 1) {
    throw new NameLookupError('ambiguous', 'sectionName', path.sectionName, deep.matches);
  }
  // `deep.seen` re-reads this group's own sections on its way to the nested ones, so it
  // already contains `sections`; using both would list every name twice.
  throw new NameLookupError(
    'not-found',
    'sectionName',
    path.sectionName,
    deep.seen.length > 0 ? deep.seen : sections,
  );
}

/**
 * Every section named `wanted` in the groups nested under `groupId`.
 *
 * Breadth-first and sequential: sequential because the concurrency limit is 5 and a
 * fan-out here would be the same mistake `getFullTree` makes, breadth-first so the
 * request bound is spent on the levels nearest the group the caller named.
 *
 * @returns the matches, and every section name seen along the way for the error message.
 */
async function findSectionBelow(
  structure: LookupStructure,
  groupId: string,
  wanted: string,
): Promise<{ matches: ResolvedNode[]; seen: ResolvedNode[] }> {
  const matches: ResolvedNode[] = [];
  const seen: ResolvedNode[] = [];

  let frontier: string[] = [groupId];
  let requests = 0;

  for (let depth = 0; depth < MAX_DEEP_DEPTH && frontier.length > 0; depth += 1) {
    const next: string[] = [];

    for (const id of frontier) {
      if (requests >= MAX_DEEP_REQUESTS) return { matches, seen };
      requests += 1;

      const children = await structure.listContainerChildren('sectionGroups', id);
      for (const section of children.sections) {
        seen.push(section);
        if (namesMatch(section.displayName, wanted)) matches.push(section);
      }
      next.push(...children.sectionGroups.map((group) => group.id));
    }

    frontier = next;
  }

  return { matches, seen };
}

/**
 * The message an `isError` tool result carries.
 *
 * It lists the names that were actually there. Those are the caller's own notebook and
 * section names, which is content this repository keeps out of logs — src/logging.ts
 * records the tool name and nothing else — but they are the answer to the question the
 * caller asked, so they go to the caller.
 */
function buildMessage(
  kind: LookupFailureKind,
  argument: string,
  wanted: string,
  candidates: readonly ResolvedNode[],
): string {
  const names = candidates.map((node) => node.displayName);
  const shown = names.slice(0, MAX_NAMES_LISTED).join(', ');
  const more = names.length > MAX_NAMES_LISTED ? ` (and ${names.length - MAX_NAMES_LISTED} more)` : '';

  if (kind === 'ambiguous') {
    return (
      `${argument} '${wanted}' matched ${String(candidates.length)} of them, so nothing was ` +
      `chosen. Use the ids from list_notebooks, list_sections or list_pages instead.`
    );
  }
  if (names.length === 0) {
    return `${argument} '${wanted}' matched nothing, and there was nothing there to match.`;
  }
  return (
    `${argument} '${wanted}' matched nothing. Names are matched in full, ignoring case. ` +
    `What was there: ${shown}${more}.`
  );
}
