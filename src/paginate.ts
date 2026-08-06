/**
 * `@odata.nextLink` pagination (docs/01-read-endpoints.md §5).
 *
 * The nextLink is a complete, pre-encoded absolute URL — GraphClient's
 * buildUrl already special-cases an absolute `path` to pass it straight
 * through unmodified (no base-URL re-prefixing, no query re-merge). That's
 * what makes following it "verbatim" here rather than reconstructing it.
 */

import type { GraphClient } from './client.js';
import type { GraphCollection } from './types.js';

/** Flat item stream across every page reachable by following nextLink. */
export async function* pageItems<T>(
  client: GraphClient,
  path: string,
  query?: Record<string, unknown>,
): AsyncGenerator<T> {
  const seen = new Set<string>();
  let url: string | undefined = path;
  let urlQuery: Record<string, unknown> | undefined = query;

  while (url !== undefined) {
    if (seen.has(url)) {
      // Non-advancing or self-referential nextLink — stop rather than spin.
      return;
    }
    seen.add(url);

    // Route through buildSpec+request (not the narrower `get<T>(path, query)`
    // overload) since this function's query type is the wider
    // Record<string, unknown> mandated by the paginate.ts contract.
    const spec = client.buildSpec('GET', url, undefined, urlQuery);
    const page: GraphCollection<T> = await client.request<GraphCollection<T>>(spec);
    for (const item of page.value) {
      yield item;
    }

    url = page['@odata.nextLink'];
    // nextLink already carries every original query param re-encoded; never
    // re-append the caller's original query on subsequent pages.
    urlQuery = undefined;
  }
}

export interface CollectItemsOptions {
  query?: Record<string, unknown>;
  /** Stop as soon as this many items are collected — truncates mid-page. */
  limit?: number;
  /** Follow every nextLink to exhaustion. Ignored when `limit` is set. */
  all?: boolean;
}

export async function collectItems<T>(
  client: GraphClient,
  path: string,
  opts?: CollectItemsOptions,
): Promise<T[]> {
  const { query, limit, all } = opts ?? {};

  if (limit === undefined && !all) {
    // Default: exactly one page, no nextLink follow at all.
    const spec = client.buildSpec('GET', path, undefined, query);
    const page = await client.request<GraphCollection<T>>(spec);
    return [...page.value];
  }

  const results: T[] = [];
  for await (const item of pageItems<T>(client, path, query)) {
    results.push(item);
    if (limit !== undefined && results.length >= limit) break;
  }
  return results;
}
