// Turning names into ids: notebook, optional section group, section.
//
// The browsing tools in ./structure-tools.ts take ids, so a caller that already knows
// the names spends three or four round trips getting to a page. This resolves the whole
// path in one Graph request, using the expanded tree described in the `Graph request
// budget` section of CLAUDE.md.
//
// Three rules the tools built on this depend on:
//
// Container names are matched by a ladder, and the result says which rung matched. Exact
// and case-insensitive first. Then the same comparison against the candidate with an
// ordering prefix removed, because this account names its section groups `062 - February`
// and a caller knows the month, not the number. Then a case-insensitive substring. Each
// rung is tried only when the one above it found nothing, so a name that matches exactly
// can never be beaten by a looser match on something else.
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

/** Which rung of the ladder produced a match. */
export type MatchRule = 'exact' | 'without-prefix' | 'substring';

/** How each name in the path was matched. */
export interface MatchRules {
  readonly notebook: MatchRule;
  readonly sectionGroup: MatchRule | null;
  readonly section: MatchRule;
}

/** A path resolved all the way to a section. */
export interface ResolvedPath {
  readonly notebook: ResolvedNode;
  readonly sectionGroup: ResolvedNode | null;
  readonly section: ResolvedNode;
  /** True when the section was found only by walking past the expanded tree. */
  readonly deepSearchUsed: boolean;
  /** Which rung matched each name, so a caller can see what it actually got. */
  readonly matchedBy: MatchRules;
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
 * A leading ordering prefix, as OneNote users write one to force a sort order.
 *
 * Covers `062 - February`, `02. February`, `2) February`, `03 February`. The digits and
 * the separator go; a name that is only digits is left alone, because removing the whole
 * name would make every such candidate match everything.
 */
const ORDERING_PREFIX = /^\s*\d+\s*(?:[-–—.):]\s*|\s+)(?=\S)/;

/** The candidate name with any ordering prefix removed. */
export function withoutOrderingPrefix(name: string): string {
  return name.replace(ORDERING_PREFIX, '').trim();
}

/** The rungs, in the order they are tried. */
const RULES: readonly { rule: MatchRule; test: (candidate: string, wanted: string) => boolean }[] = [
  { rule: 'exact', test: namesMatch },
  {
    rule: 'without-prefix',
    // `062 - February` matched by `February`. The prefix comes off the stored name, not
    // off what the caller typed: the caller is the one who does not know the number.
    test: (candidate, wanted) => namesMatch(withoutOrderingPrefix(candidate), wanted),
  },
  {
    rule: 'substring',
    test: (candidate, wanted) =>
      wanted.trim() !== '' && candidate.trim().toLowerCase().includes(wanted.trim().toLowerCase()),
  },
];

/**
 * Everything that matches `wanted`, by the strictest rung that matches anything.
 *
 * @returns the matches and the rung, or an empty list and a null rung.
 */
export function matchNodes(
  nodes: readonly ResolvedNode[],
  wanted: string,
): { matches: ResolvedNode[]; rule: MatchRule | null } {
  for (const { rule, test } of RULES) {
    const matches = nodes.filter((node) => test(node.displayName, wanted));
    if (matches.length > 0) return { matches, rule };
  }
  return { matches: [], rule: null };
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
): { node: ResolvedNode; rule: MatchRule } {
  const { matches, rule } = matchNodes(nodes, wanted);

  // Narrowed to id and display name: the nodes coming in are expanded notebooks and
  // section groups carrying their children, and a ResolvedPath that leaked those would
  // put the whole subtree into every tool result.
  if (matches.length === 1 && rule !== null) {
    return { node: plain(matches[0] as ResolvedNode), rule };
  }
  if (matches.length === 0) {
    throw new NameLookupError('not-found', argument, wanted, nodes.map(plain));
  }
  throw new NameLookupError('ambiguous', argument, wanted, matches.map(plain));
}

/**
 * Resolve notebook → section group → section, in one Graph request where possible.
 *
 * Each name goes through the ladder in `matchNodes`, and `matchedBy` in the result says
 * which rung answered. That is what lets `February` find `062 - February` without ever
 * letting a loose match win over an exact one.
 *
 * The expanded tree covers a notebook's own sections and one level of section group. A
 * section that is not there may still exist deeper, so one filtered account-wide request
 * settles that case before the lookup is called a failure.
 *
 * @throws {NameLookupError} when a name matches nothing or matches more than once.
 */
export async function resolveSection(
  structure: LookupStructure,
  path: NamePath,
): Promise<ResolvedPath> {
  const tree = await structure.getExpandedTree();
  const notebookMatch = matchOne(tree, path.notebookName, 'notebookName');
  const notebook = notebookMatch.node;
  const expanded = tree.find((candidate) => candidate.id === notebook.id) as ExpandedNotebook;

  if (path.sectionGroupName === undefined) {
    const direct = matchNodes(expanded.sections, path.sectionName);
    if (direct.matches.length === 1 && direct.rule !== null) {
      return {
        notebook,
        sectionGroup: null,
        section: plain(direct.matches[0] as ResolvedNode),
        deepSearchUsed: false,
        matchedBy: { notebook: notebookMatch.rule, sectionGroup: null, section: direct.rule },
      };
    }
    if (direct.matches.length > 1) {
      throw new NameLookupError('ambiguous', 'sectionName', path.sectionName, direct.matches);
    }
    // Not among the notebook's own sections. It may sit inside a section group the
    // caller did not name, which is worth saying rather than reporting a bare miss.
    throw new NameLookupError('not-found', 'sectionName', path.sectionName, expanded.sections);
  }

  const groupMatch = matchOne(expanded.sectionGroups, path.sectionGroupName, 'sectionGroupName');
  const group = groupMatch.node;
  const expandedGroup = expanded.sectionGroups.find((candidate) => candidate.id === group.id);
  const sections = expandedGroup?.sections ?? [];

  const inGroup = matchNodes(sections, path.sectionName);
  if (inGroup.matches.length === 1 && inGroup.rule !== null) {
    return {
      notebook,
      sectionGroup: group,
      section: plain(inGroup.matches[0] as ResolvedNode),
      deepSearchUsed: false,
      matchedBy: {
        notebook: notebookMatch.rule,
        sectionGroup: groupMatch.rule,
        section: inGroup.rule,
      },
    };
  }
  if (inGroup.matches.length > 1) {
    throw new NameLookupError('ambiguous', 'sectionName', path.sectionName, inGroup.matches);
  }

  // The expanded response stops at this group's sections, so a section nested below it
  // is absent from that response rather than known to be absent. One filtered request
  // settles it at any depth, and only runs when the cheap answer came back empty.
  const candidates = (await structure.findSectionsByName(path.sectionName)).filter(
    (section) =>
      section.parentSectionGroup !== null &&
      namesMatch(section.parentSectionGroup.displayName, group.displayName) &&
      // Two notebooks can hold a section group of the same name, so the notebook is
      // checked too rather than assumed from the group.
      section.parentNotebook !== null &&
      namesMatch(section.parentNotebook.displayName, notebook.displayName),
  );

  // Graph matched a substring of its own; the ladder is this module's rule, applied here.
  const deep = matchNodes(candidates, path.sectionName);
  if (deep.matches.length === 1 && deep.rule !== null) {
    return {
      notebook,
      sectionGroup: group,
      section: plain(deep.matches[0] as ResolvedNode),
      deepSearchUsed: true,
      matchedBy: {
        notebook: notebookMatch.rule,
        sectionGroup: groupMatch.rule,
        section: deep.rule,
      },
    };
  }
  if (deep.matches.length > 1) {
    throw new NameLookupError('ambiguous', 'sectionName', path.sectionName, deep.matches);
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
    `${argument} '${wanted}' matched nothing. A name is matched in full ignoring case, ` +
    `then against the name with any leading number removed ('February' finds ` +
    `'062 - February'), then as a substring. What was there: ${shown}${more}.`
  );
}
