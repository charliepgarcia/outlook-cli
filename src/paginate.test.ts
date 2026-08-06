import { describe, it, expect, vi } from 'vitest';
import { GraphClient } from './client.js';
import { pageItems, collectItems } from './paginate.js';
import { GRAPH_BASE } from './types.js';

function jsonResponse(body: unknown): Response {
  return new Response(JSON.stringify(body), { status: 200 });
}

function clientWithFetch(fetchImpl: ReturnType<typeof vi.fn>): GraphClient {
  return new GraphClient({ getToken: async () => 'tok', fetchImpl });
}

describe('pageItems / collectItems', () => {
  it('follows @odata.nextLink across 3 pages and yields every item', async () => {
    const page1Next = `${GRAPH_BASE}/me/messages?$skip=2`;
    const page2Next = `${GRAPH_BASE}/me/messages?$skip=4`;
    const fetchImpl = vi
      .fn()
      .mockResolvedValueOnce(
        jsonResponse({ value: [{ id: '1' }, { id: '2' }], '@odata.nextLink': page1Next }),
      )
      .mockResolvedValueOnce(
        jsonResponse({ value: [{ id: '3' }, { id: '4' }], '@odata.nextLink': page2Next }),
      )
      .mockResolvedValueOnce(jsonResponse({ value: [{ id: '5' }] }));
    const client = clientWithFetch(fetchImpl);

    const items = await collectItems<{ id: string }>(client, '/me/messages', { all: true });

    expect(items.map((i) => i.id)).toEqual(['1', '2', '3', '4', '5']);
    expect(fetchImpl).toHaveBeenCalledTimes(3);
    // Page 2/3 requests must hit the nextLink URL directly, not a re-prefixed one.
    expect(fetchImpl.mock.calls[1]?.[0]).toBe(page1Next);
    expect(fetchImpl.mock.calls[2]?.[0]).toBe(page2Next);
  });

  it('default (no limit, no all) returns just the first page', async () => {
    const fetchImpl = vi
      .fn()
      .mockResolvedValueOnce(
        jsonResponse({ value: [{ id: '1' }, { id: '2' }], '@odata.nextLink': `${GRAPH_BASE}/me/messages?$skip=2` }),
      );
    const client = clientWithFetch(fetchImpl);

    const items = await collectItems<{ id: string }>(client, '/me/messages');

    expect(items.map((i) => i.id)).toEqual(['1', '2']);
    expect(fetchImpl).toHaveBeenCalledTimes(1);
  });

  it('limit truncates mid-page without fetching further pages than necessary', async () => {
    const fetchImpl = vi
      .fn()
      .mockResolvedValueOnce(
        jsonResponse({
          value: [{ id: '1' }, { id: '2' }, { id: '3' }],
          '@odata.nextLink': `${GRAPH_BASE}/me/messages?$skip=3`,
        }),
      );
    const client = clientWithFetch(fetchImpl);

    const items = await collectItems<{ id: string }>(client, '/me/messages', { limit: 2 });

    expect(items.map((i) => i.id)).toEqual(['1', '2']);
    expect(fetchImpl).toHaveBeenCalledTimes(1);
  });

  it('a self-referential nextLink terminates instead of looping forever', async () => {
    const selfLink = `${GRAPH_BASE}/me/messages?$skip=0`;
    const fetchImpl = vi.fn().mockResolvedValue(jsonResponse({ value: [{ id: '1' }], '@odata.nextLink': selfLink }));
    const client = clientWithFetch(fetchImpl);

    const items: { id: string }[] = [];
    for await (const item of pageItems<{ id: string }>(client, selfLink)) {
      items.push(item);
    }

    expect(items).toEqual([{ id: '1' }]);
    expect(fetchImpl).toHaveBeenCalledTimes(1);
  });

  it('a non-advancing nextLink (identical to the request path) also terminates', async () => {
    const path = '/me/messages';
    const fetchImpl = vi.fn().mockResolvedValue(jsonResponse({ value: [{ id: 'x' }], '@odata.nextLink': path }));
    const client = clientWithFetch(fetchImpl);

    const items: { id: string }[] = [];
    for await (const item of pageItems<{ id: string }>(client, path)) {
      items.push(item);
    }

    expect(items).toEqual([{ id: 'x' }]);
    expect(fetchImpl).toHaveBeenCalledTimes(1);
  });
});
