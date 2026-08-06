import { describe, it, expect, vi } from 'vitest';
import { GraphClient } from './client.js';
import { GraphError } from './errors.js';
import { GRAPH_BASE } from './types.js';

function jsonResponse(body: unknown, status = 200, headers: Record<string, string> = {}): Response {
  return new Response(JSON.stringify(body), { status, headers });
}

function noSleep() {
  return vi.fn(async (_ms: number) => {});
}

describe('GraphClient — retry policy', () => {
  it('429 with Retry-After: 3 waits ~3s then succeeds', async () => {
    const sleep = noSleep();
    const fetchImpl = vi
      .fn()
      .mockResolvedValueOnce(new Response('', { status: 429, headers: { 'Retry-After': '3' } }))
      .mockResolvedValueOnce(jsonResponse({ ok: true }));
    const client = new GraphClient({ getToken: async () => 'tok', fetchImpl, sleep });

    const result = await client.get('/me');

    expect(result).toEqual({ ok: true });
    expect(sleep).toHaveBeenCalledTimes(1);
    expect(sleep).toHaveBeenCalledWith(3000);
    expect(fetchImpl).toHaveBeenCalledTimes(2);
  });

  it('429 without Retry-After falls back to backoff and still succeeds', async () => {
    const sleep = noSleep();
    const fetchImpl = vi
      .fn()
      .mockResolvedValueOnce(new Response('', { status: 429 }))
      .mockResolvedValueOnce(jsonResponse({ value: [] }));
    const client = new GraphClient({ getToken: async () => 'tok', fetchImpl, sleep });

    const result = await client.get('/me/messages');

    expect(result).toEqual({ value: [] });
    expect(sleep).toHaveBeenCalledTimes(1);
    // No Retry-After header present, so it must not have taken the fixed-seconds path.
    const delayArg = sleep.mock.calls[0]?.[0];
    expect(typeof delayArg).toBe('number');
    expect(delayArg).toBeGreaterThanOrEqual(0);
    expect(delayArg).toBeLessThanOrEqual(30_000);
  });

  it('does NOT retry POST on 429 — fetch called exactly once, error surfaced', async () => {
    const sleep = noSleep();
    const fetchImpl = vi.fn().mockResolvedValue(
      new Response(
        JSON.stringify({ error: { code: 'TooManyRequests', message: 'Please retry again later.' } }),
        { status: 429, headers: { 'Retry-After': '5' } },
      ),
    );
    const client = new GraphClient({ getToken: async () => 'tok', fetchImpl, sleep });

    await expect(client.post('/me/sendMail', { message: { subject: 'hi' } })).rejects.toThrow(GraphError);

    expect(fetchImpl).toHaveBeenCalledTimes(1);
    expect(sleep).not.toHaveBeenCalled();
  });

  it('GET IS retried on 429/503/504', async () => {
    const sleep = noSleep();
    const fetchImpl = vi
      .fn()
      .mockResolvedValueOnce(new Response('', { status: 503 }))
      .mockResolvedValueOnce(new Response('', { status: 504 }))
      .mockResolvedValueOnce(jsonResponse({ id: '1' }));
    const client = new GraphClient({ getToken: async () => 'tok', fetchImpl, sleep });

    const result = await client.get('/me/messages/1');

    expect(result).toEqual({ id: '1' });
    expect(fetchImpl).toHaveBeenCalledTimes(3);
    expect(sleep).toHaveBeenCalledTimes(2);
  });

  it('retries PATCH and DELETE (idempotent verbs) on a transient 503', async () => {
    const sleep = noSleep();
    const patchFetch = vi
      .fn()
      .mockResolvedValueOnce(new Response('', { status: 503 }))
      .mockResolvedValueOnce(jsonResponse({ isRead: true }));
    const patchClient = new GraphClient({ getToken: async () => 'tok', fetchImpl: patchFetch, sleep });
    await patchClient.patch('/me/messages/1', { isRead: true });
    expect(patchFetch).toHaveBeenCalledTimes(2);

    const delFetch = vi
      .fn()
      .mockResolvedValueOnce(new Response('', { status: 503 }))
      .mockResolvedValueOnce(new Response(null, { status: 204 }));
    const delClient = new GraphClient({ getToken: async () => 'tok', fetchImpl: delFetch, sleep });
    await delClient.del('/me/messages/1');
    expect(delFetch).toHaveBeenCalledTimes(2);
  });

  it('maxRetries exhaustion throws GraphError', async () => {
    const sleep = noSleep();
    const fetchImpl = vi.fn().mockResolvedValue(
      new Response(JSON.stringify({ error: { code: 'TooManyRequests', message: 'Please retry again later.' } }), {
        status: 429,
      }),
    );
    const client = new GraphClient({ getToken: async () => 'tok', fetchImpl, sleep, maxRetries: 2 });

    await expect(client.get('/me/messages')).rejects.toThrow(GraphError);
    // Initial attempt + 2 retries = 3 total.
    expect(fetchImpl).toHaveBeenCalledTimes(3);
  });
});

describe('GraphClient — response body handling', () => {
  it('204 No Content resolves to undefined', async () => {
    const fetchImpl = vi.fn().mockResolvedValue(new Response(null, { status: 204 }));
    const client = new GraphClient({ getToken: async () => 'tok', fetchImpl });

    const result = await client.del('/me/messages/1');

    expect(result).toBeUndefined();
  });

  it('non-JSON error body does not crash — still throws a GraphError', async () => {
    const fetchImpl = vi.fn().mockResolvedValue(new Response('<html>Bad Gateway</html>', { status: 502 }));
    const client = new GraphClient({ getToken: async () => 'tok', fetchImpl, sleep: noSleep() });

    await expect(client.get('/me')).rejects.toThrow(GraphError);
    expect(fetchImpl).toHaveBeenCalledTimes(1);
  });
});

describe('GraphClient — headers', () => {
  it('attaches Prefer: IdType="ImmutableId" on GET requests', async () => {
    const fetchImpl = vi.fn().mockResolvedValue(jsonResponse({ value: [] }));
    const client = new GraphClient({ getToken: async () => 'tok', fetchImpl });

    await client.get('/me/messages');

    const init = fetchImpl.mock.calls[0]?.[1] as RequestInit;
    const headers = init.headers as Record<string, string>;
    expect(headers['Prefer']).toBe('IdType="ImmutableId"');
    expect(headers['Authorization']).toBe('Bearer tok');
  });

  it('attaches Content-Type: application/json on bodied requests', async () => {
    const fetchImpl = vi.fn().mockResolvedValue(jsonResponse({ id: '1' }, 201));
    const client = new GraphClient({ getToken: async () => 'tok', fetchImpl });

    await client.post('/me/messages', { subject: 'hi' });

    const init = fetchImpl.mock.calls[0]?.[1] as RequestInit;
    const headers = init.headers as Record<string, string>;
    expect(headers['Content-Type']).toBe('application/json');
  });
});

describe('GraphClient — buildSpec purity', () => {
  it('buildSpec never calls getToken() and never touches the network', () => {
    const getToken = vi.fn(async () => 'tok');
    const fetchImpl = vi.fn();
    const client = new GraphClient({ getToken, fetchImpl });

    const spec = client.buildSpec('GET', '/me');

    expect(spec.method).toBe('GET');
    expect(spec.url).toBe(`${GRAPH_BASE}/me`);
    expect(getToken).not.toHaveBeenCalled();
    expect(fetchImpl).not.toHaveBeenCalled();
  });

  it('skips undefined query values and does not over-encode OData $ params', () => {
    const client = new GraphClient({ getToken: async () => 'tok' });

    const spec = client.buildSpec('GET', '/me/messages', undefined, {
      $select: 'id,subject',
      $filter: "isRead eq false",
      $top: undefined,
    });

    expect(spec.url).toContain('$select=id%2Csubject');
    expect(spec.url).toContain('$filter=isRead%20eq%20false');
    expect(spec.url).not.toContain('$top');
    expect(spec.url).not.toContain('undefined');
  });

  it('passes an absolute nextLink-style URL straight through without re-prefixing the base', () => {
    const client = new GraphClient({ getToken: async () => 'tok' });
    const nextLink = `${GRAPH_BASE}/me/messages?$skip=10&$top=10`;

    const spec = client.buildSpec('GET', nextLink);

    expect(spec.url).toBe(nextLink);
    expect(spec.url.indexOf(GRAPH_BASE)).toBe(0);
    expect(spec.url.match(new RegExp(GRAPH_BASE.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g'))?.length).toBe(1);
  });
});
