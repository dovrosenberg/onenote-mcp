// Turning a Graph response body into the shapes this repository uses.
//
// Split out of ./graph-structure.ts, which was doing three jobs at once: building URLs,
// running requests under the gate, and decoding what came back. Only the third is here,
// and none of it touches the network — every function takes an already-parsed value and
// the URL it came from, and either returns a typed object or throws
// `GraphResponseError`. That is what makes it testable without a fake `fetch`.
//
// Two rules hold throughout, and both are load-bearing:
//
// **A `url` is threaded through every function purely so a failure can name the request.**
// It is never inspected. `GraphResponseError` carries it because a decode failure with no
// request attached tells an operator nothing about which of the fifty calls in a walk
// produced it.
//
// **No thrown message quotes the value that failed.** These bodies carry notebook,
// section and page names, which are the user's own writing, and this repository's output
// can reach a public log. A message says which key was unusable and never what was in it.
//
// The tolerances are deliberate rather than lazy. A null `displayName` becomes an empty
// string, because a caller that only formats the name should not have to handle
// `undefined`. An absent expanded relationship becomes an empty array, because Graph
// omits a relationship that holds nothing and a notebook with no section groups is
// ordinary. A relationship present but not an array is a fault, because that is the
// service returning a shape nothing here can use.

import { GraphResponseError } from './graph-structure.ts';
import type {
  ExpandedNotebook,
  PageSummary,
  SectionWithParents,
} from './graph-structure.ts';

/** A notebook, section group, or section: the two fields every node has. */
export function toNode(item: unknown, url: string): { id: string; displayName: string } {
  const record = asRecord(item, url);
  return {
    id: requireString(record, 'id', url),
    // A notebook or section can come back with a null displayName; an empty string is
    // preferable to `undefined` reaching a caller that only formats it.
    displayName: optionalString(record, 'displayName') ?? '',
  };
}

/**
 * One notebook out of the expanded response.
 *
 * An absent `sections` or `sectionGroups` is read as empty rather than raised on: Graph
 * omits an expanded relationship that holds nothing, and a notebook with no section
 * groups is ordinary.
 */
export function toExpandedNotebook(item: unknown, url: string): ExpandedNotebook {
  const record = asRecord(item, url);
  return {
    ...toNode(record, url),
    sections: toNodeArray(record['sections'], url).map((node) => toNode(node, url)),
    sectionGroups: toNodeArray(record['sectionGroups'], url).map((group) => {
      const groupRecord = asRecord(group, url);
      return {
        ...toNode(groupRecord, url),
        sections: toNodeArray(groupRecord['sections'], url).map((node) => toNode(node, url)),
      };
    }),
  };
}

/** An expanded relationship: absent means empty, anything but an array is a fault. */
export function toNodeArray(value: unknown, url: string): unknown[] {
  if (value === undefined || value === null) return [];
  if (!Array.isArray(value)) {
    throw new GraphResponseError(
      `GET ${url} returned an expanded relationship that is not an array.`,
      url,
    );
  }
  return value;
}

export function toSectionWithParents(item: unknown, url: string): SectionWithParents {
  const record = asRecord(item, url);
  const notebook = record['parentNotebook'];
  const group = record['parentSectionGroup'];
  return {
    ...toNode(record, url),
    // A section directly under a notebook has no parent section group, and Graph returns
    // null rather than omitting it. Both readings end as null here.
    parentNotebook: isRecord(notebook) ? toNode(notebook, url) : null,
    parentSectionGroup: isRecord(group) ? toNode(group, url) : null,
  };
}

export function toPageSummary(item: unknown, url: string): PageSummary {
  const record = asRecord(item, url);
  return {
    id: requireString(record, 'id', url),
    title: optionalString(record, 'title') ?? '',
    lastModifiedDateTime: optionalString(record, 'lastModifiedDateTime') ?? '',
  };
}

export function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

export function asRecord(item: unknown, url: string): Record<string, unknown> {
  if (typeof item !== 'object' || item === null || Array.isArray(item)) {
    throw new GraphResponseError(`GET ${url} returned a non-object inside "value".`, url);
  }
  return item as Record<string, unknown>;
}

export function requireString(
  record: Record<string, unknown>,
  key: string,
  url: string,
): string {
  const value = record[key];
  if (typeof value !== 'string' || value === '') {
    // The offending value is not printed: these bodies carry user content, and the
    // repository's hygiene rules keep page and notebook names out of anything loggable.
    throw new GraphResponseError(`GET ${url} returned an item with no usable "${key}".`, url);
  }
  return value;
}

export function optionalString(
  record: Record<string, unknown>,
  key: string,
): string | undefined {
  const value = record[key];
  return typeof value === 'string' ? value : undefined;
}

/** A single-quoted OData string literal; an embedded quote is doubled. */
export function quoteOData(value: string): string {
  return `'${value.replaceAll("'", "''")}'`;
}

/**
 * `Promise.all` with a ceiling on how many run at once, preserving input order.
 *
 * Written out rather than pulled in as a dependency because it is fifteen lines and the
 * thing it prevents — an unbounded fan-out against an API that allows five concurrent
 * requests — is the most expensive mistake this repository has made.
 */
export async function mapWithLimit<T, R>(
  items: readonly T[],
  limit: number,
  run: (item: T) => Promise<R>,
): Promise<R[]> {
  const results: R[] = new Array<R>(items.length);
  let cursor = 0;

  const worker = async (): Promise<void> => {
    for (;;) {
      const index = cursor;
      if (index >= items.length) return;
      cursor += 1;
      results[index] = await run(items[index] as T);
    }
  };

  await Promise.all(Array.from({ length: Math.min(limit, items.length) }, () => worker()));
  return results;
}

/** An error body that cannot be read must not mask the status that caused the throw. */
export async function safeText(response: Response): Promise<string> {
  try {
    return await response.text();
  } catch {
    return '';
  }
}

export function truncate(text: string, max: number): string {
  return text.length <= max ? text : `${text.slice(0, max)}… (${text.length} chars)`;
}

export function describeError(err: unknown): string {
  if (err instanceof Error) return `${err.name}: ${err.message}`;
  return String(err);
}
