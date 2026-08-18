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

import type { ExpandedNotebook, SectionWithParents } from './graph-structure.ts';

/** The slice of `GraphStructure` this module calls, so a test can pass a plain object. */
export interface LookupStructure {
  getExpandedTree(): Promise<ExpandedNotebook[]>;
  findSectionsByName(displayName: string): Promise<SectionWithParents[]>;
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
 * The fallback is one filtered request, not a walk.
 *
 * `findSectionsByName` returns sections at any depth with their parents, so a section
 * nested below what the expanded tree reaches costs one request rather than one per
 * container. It matches on a substring, which is the widest thing that endpoint will
 * answer case-insensitively, so the full-name comparison is applied here.
 */

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

  // The expanded response stops at this group's sections, so a section nested below it
  // is absent from that response rather than known to be absent. One filtered request
  // settles it at any depth, and only runs when the cheap answer came back empty.
  const deep = (await structure.findSectionsByName(path.sectionName)).filter(
    (section) =>
      // Graph matched a substring; the full-name rule is this module's, so it is applied
      // here rather than left to the service.
      namesMatch(section.displayName, path.sectionName) &&
      section.parentSectionGroup !== null &&
      namesMatch(section.parentSectionGroup.displayName, path.sectionGroupName as string) &&
      // Two notebooks can hold a section group of the same name, so the notebook is
      // checked too rather than assumed from the group.
      section.parentNotebook !== null &&
      namesMatch(section.parentNotebook.displayName, notebook.displayName),
  );

  if (deep.length === 1) {
    return {
      notebook,
      sectionGroup: group,
      section: plain(deep[0] as ResolvedNode),
      deepSearchUsed: true,
    };
  }
  if (deep.length > 1) {
    throw new NameLookupError('ambiguous', 'sectionName', path.sectionName, deep.map(plain));
  }
  throw new NameLookupError('not-found', 'sectionName', path.sectionName, sections);
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
