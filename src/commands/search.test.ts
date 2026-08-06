/**
 * Tests for `outlook search`. Mirrors folders.test.ts's stubbing idiom:
 * collectItems/pageItems call `client.request` (never `client.get`), so the
 * network boundary to stub is `client.request` — stubbing `get` would leave
 * `request` live and trip testUtils' deliberately-throwing fetchImpl.
 */

import { describe, expect, it, vi } from 'vitest';
import type { GraphClient } from '../client.js';
import { UsageError } from '../errors.js';
import { GRAPH_BASE, type GraphCollection, type Message, type RequestSpec } from '../types.js';
import { searchCommand } from './search.js';
import { makeCtx, text } from './testUtils.js';

function stubPages(client: GraphClient, ...pages: GraphCollection<Message>[]) {
  const spy = vi.spyOn(client, 'request');
  for (const page of pages) spy.mockResolvedValueOnce(page as never);
  return spy;
}

function specOf(spy: { mock: { calls: unknown[][] } }, call = 0): RequestSpec {
  const spec = spy.mock.calls[call]?.[0] as RequestSpec | undefined;
  if (spec === undefined) throw new Error('client.request was not called');
  return spec;
}

/** Relative Graph path of a captured RequestSpec (GRAPH_BASE + query stripped). */
function pathOf(spy: { mock: { calls: unknown[][] } }, call = 0): string {
  const url = specOf(spy, call).url;
  const stripped = url.startsWith(GRAPH_BASE) ? url.slice(GRAPH_BASE.length) : url;
  return stripped.split('?')[0] ?? '';
}

/** Query string of a captured RequestSpec, parsed for individual-param assertions. */
function queryOf(spy: { mock: { calls: unknown[][] } }, call = 0): URLSearchParams {
  const url = specOf(spy, call).url;
  return new URLSearchParams(url.split('?')[1] ?? '');
}

function sampleMessage(overrides: Partial<Message> = {}): Message {
  return {
    id: 'AAMkADhMGAAA=',
    subject: 'Pizza night?',
    from: { emailAddress: { address: 'frannis@contoso.com', name: 'Frannis' } },
    receivedDateTime: '2024-06-01T12:00:00Z',
    importance: 'normal',
    isRead: true,
    hasAttachments: false,
    ...overrides,
  };
}

describe('search', () => {
  it('searches /me/messages with the quoted query and renders results', async () => {
    const { ctx, client, out } = makeCtx(['pizza']);
    const req = stubPages(client, { value: [sampleMessage()] });

    await searchCommand.handler(ctx);

    expect(req).toHaveBeenCalledTimes(1);
    expect(pathOf(req)).toBe('/me/messages');
    expect(queryOf(req).get('$search')).toBe('"pizza"');
    expect(text(out)).toContain('Pizza night?');
  });

  it('requires a query and throws a clean UsageError when absent', async () => {
    const { ctx, client } = makeCtx([]);
    const req = vi.spyOn(client, 'request');

    await expect(searchCommand.handler(ctx)).rejects.toBeInstanceOf(UsageError);
    expect(req).toHaveBeenCalledTimes(0);
  });

  it('wraps the query in one pair of double quotes, backslash-escaping embedded quotes/backslashes', async () => {
    const { ctx, client } = makeCtx(['say "hi" \\ bye']);
    const req = stubPages(client, { value: [] });

    await searchCommand.handler(ctx);

    expect(queryOf(req).get('$search')).toBe('"say \\"hi\\" \\\\ bye"');
  });

  it('KQL property-scoped clauses (e.g. from:) pass through unmodified inside the outer quotes', async () => {
    const { ctx, client } = makeCtx(['from:randiw']);
    const req = stubPages(client, { value: [] });

    await searchCommand.handler(ctx);

    expect(queryOf(req).get('$search')).toBe('"from:randiw"');
  });

  it('--limit sets $top and is forwarded to pagination', async () => {
    const { ctx, client } = makeCtx(['pizza', '--limit', '5']);
    const req = stubPages(client, { value: [sampleMessage()] });

    await searchCommand.handler(ctx);

    expect(queryOf(req).get('$top')).toBe('5');
  });

  it('rejects --limit above the documented 1,000-result ceiling for $search', async () => {
    const { ctx, client } = makeCtx(['pizza', '--limit', '5000']);
    const req = vi.spyOn(client, 'request');

    await expect(searchCommand.handler(ctx)).rejects.toBeInstanceOf(UsageError);
    expect(req).toHaveBeenCalledTimes(0);
  });

  it('--folder with a well-known name scopes the search and lowercases the name', async () => {
    const { ctx, client } = makeCtx(['pizza', '--folder', 'Inbox']);
    const req = stubPages(client, { value: [] });

    await searchCommand.handler(ctx);

    expect(pathOf(req)).toBe('/me/mailFolders/inbox/messages');
  });

  it('--folder with a raw (non-well-known) id is percent-encoded, not lowercased', async () => {
    const rawId = 'AAMk/odd=folder';
    const { ctx, client } = makeCtx(['pizza', '--folder', rawId]);
    const req = stubPages(client, { value: [] });

    await searchCommand.handler(ctx);

    expect(pathOf(req)).toBe(`/me/mailFolders/${encodeURIComponent(rawId)}/messages`);
  });

  it('never sends $orderby or $filter — $search has its own documented built-in sort', async () => {
    const { ctx, client } = makeCtx(['pizza']);
    const req = stubPages(client, { value: [sampleMessage()] });

    await searchCommand.handler(ctx);

    expect(queryOf(req).get('$orderby')).toBeNull();
    expect(queryOf(req).get('$filter')).toBeNull();
  });

  it('--json emits the raw message array, unmodified', async () => {
    const { ctx, client, out } = makeCtx(['pizza', '--json']);
    const messages = [sampleMessage()];
    stubPages(client, { value: messages });

    await searchCommand.handler(ctx);

    expect(JSON.parse(text(out))).toEqual(messages);
  });

  it('zero results render as an empty JSON array, not an error', async () => {
    const { ctx, client, out } = makeCtx(['pizza', '--json']);
    stubPages(client, { value: [] });

    await searchCommand.handler(ctx);

    expect(JSON.parse(text(out))).toEqual([]);
  });

  it('--all paginates through every page', async () => {
    const { ctx, client } = makeCtx(['pizza', '--all']);
    const req = stubPages(
      client,
      {
        value: [sampleMessage({ id: 'p1' })],
        '@odata.nextLink': `${GRAPH_BASE}/me/messages?$skip=1`,
      },
      { value: [sampleMessage({ id: 'p2' })] },
    );

    await searchCommand.handler(ctx);

    expect(req).toHaveBeenCalledTimes(2);
  });

  it('never touches the network directly — only through the stubbed client method', async () => {
    const { ctx, client, fetchImpl } = makeCtx(['pizza']);
    stubPages(client, { value: [sampleMessage()] });

    await searchCommand.handler(ctx);

    expect(fetchImpl).toHaveBeenCalledTimes(0);
  });
});
