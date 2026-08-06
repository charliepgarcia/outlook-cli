/**
 * Tests for `outlook list`. Mirrors folders.test.ts's idiom: real GraphClient
 * from makeCtx, network boundary stubbed via vi.spyOn(client, 'request')
 * (collectItems/pageItems build a RequestSpec via client.buildSpec and hand
 * it straight to client.request — NOT client.get — so that's the seam to
 * stub; stubbing `get` leaves `request` live and testUtils' throwing
 * fetchImpl trips with "NETWORK CALLED").
 *
 * Endpoint/$orderby/$filter/$select verified against docs/01-read-endpoints.md:
 *   - §2 "List messages in a specific folder": GET /me/mailFolders/{id}/messages
 *   - §3.2: `$filter=isRead eq false` is a confirmed working example
 *   - §3.3/§3.4: newest-first is `$orderby=receivedDateTime desc`; keeping
 *     `isRead` out of $orderby (and receivedDateTime out of $filter) avoids
 *     the documented InefficientFilter constraint
 *   - §2 well-known folder names (inbox, drafts, sentitems, deleteditems, …)
 *     are lowercase tokens usable in place of `{id}`
 */

import { describe, expect, it, vi } from 'vitest';
import type { GraphClient } from '../client.js';
import { UsageError } from '../errors.js';
import { GRAPH_BASE, type GraphCollection, type Message, type RequestSpec } from '../types.js';
import { listCommand } from './list.js';
import { makeCtx, text } from './testUtils.js';

function stubPages(client: GraphClient, ...pages: GraphCollection<Message>[]) {
  const spy = vi.spyOn(client, 'request');
  for (const page of pages) spy.mockResolvedValueOnce(page as never);
  return spy;
}

/** Relative Graph path of a captured RequestSpec (spec.url is absolute). */
function pathOf(spy: { mock: { calls: unknown[][] } }, call = 0): string {
  const spec = spy.mock.calls[call]?.[0] as RequestSpec | undefined;
  if (spec === undefined) return '';
  const url = spec.url.startsWith(GRAPH_BASE) ? spec.url.slice(GRAPH_BASE.length) : spec.url;
  return url.split('?')[0] ?? '';
}

function queryOf(spy: { mock: { calls: unknown[][] } }, call = 0): URLSearchParams {
  const spec = spy.mock.calls[call]?.[0] as RequestSpec | undefined;
  return new URL(spec?.url ?? GRAPH_BASE).searchParams;
}

function sampleMessage(overrides: Partial<Message> = {}): Message {
  return {
    id: 'AAMk-msg-0001',
    subject: 'Quarterly numbers',
    from: { emailAddress: { name: 'Alice', address: 'alice@example.com' } },
    receivedDateTime: '2024-06-01T12:00:00Z',
    isRead: false,
    hasAttachments: false,
    bodyPreview: 'Please see attached...',
    ...overrides,
  } as unknown as Message;
}

describe('list', () => {
  it('defaults: folder=inbox, limit=25, single page, no $filter', async () => {
    const { ctx, client } = makeCtx([]);
    const req = stubPages(client, { value: [sampleMessage()] });

    await listCommand.handler(ctx);

    expect(req).toHaveBeenCalledTimes(1);
    expect(pathOf(req)).toBe('/me/mailFolders/inbox/messages');
    expect(queryOf(req).get('$top')).toBe('25');
    expect(queryOf(req).has('$filter')).toBe(false);
  });

  it('renders the default (non-JSON) output via renderMessageList', async () => {
    const { ctx, client, out } = makeCtx([]);
    stubPages(client, { value: [sampleMessage({ subject: 'Renders me' })] });

    await listCommand.handler(ctx);

    expect(text(out)).toContain('Renders me');
    expect(() => JSON.parse(text(out))).toThrow();
  });

  it('--json emits the raw message array, unmodified', async () => {
    const { ctx, client, out } = makeCtx(['--json']);
    const messages = [sampleMessage()];
    stubPages(client, { value: messages });

    await listCommand.handler(ctx);

    expect(JSON.parse(text(out))).toEqual(messages);
  });

  it('--limit <n> reaches collectItems as both $top and the page-size cap', async () => {
    const { ctx, client } = makeCtx(['--limit', '5']);
    const req = stubPages(client, { value: [sampleMessage()] });

    await listCommand.handler(ctx);

    expect(req).toHaveBeenCalledTimes(1);
    expect(queryOf(req).get('$top')).toBe('5');
  });

  it('rejects --limit outside the documented 1-1000 $top range', async () => {
    const { ctx: ctxZero, client: clientZero } = makeCtx(['--limit', '0']);
    const reqZero = vi.spyOn(clientZero, 'request');
    await expect(listCommand.handler(ctxZero)).rejects.toBeInstanceOf(UsageError);
    expect(reqZero).toHaveBeenCalledTimes(0);

    const { ctx: ctxBig, client: clientBig } = makeCtx(['--limit', '1001']);
    const reqBig = vi.spyOn(clientBig, 'request');
    await expect(listCommand.handler(ctxBig)).rejects.toBeInstanceOf(UsageError);
    expect(reqBig).toHaveBeenCalledTimes(0);
  });

  it('--all paginates through every page via @odata.nextLink', async () => {
    const { ctx, client } = makeCtx(['--all']);
    const req = stubPages(
      client,
      {
        value: [sampleMessage({ id: 'p1' })],
        '@odata.nextLink': `${GRAPH_BASE}/me/mailFolders/inbox/messages?$skip=1`,
      },
      { value: [sampleMessage({ id: 'p2' })] },
    );

    await listCommand.handler(ctx);

    expect(req).toHaveBeenCalledTimes(2);
  });

  it('--unread adds $filter=isRead eq false as its own param, $orderby unchanged', async () => {
    const { ctx, client } = makeCtx(['--unread']);
    const req = stubPages(client, { value: [sampleMessage()] });

    await listCommand.handler(ctx);

    expect(queryOf(req).get('$filter')).toBe('isRead eq false');
    expect(queryOf(req).get('$orderby')).toBe('receivedDateTime desc');
  });

  it('$orderby is newest-first per docs/01 §3.3', async () => {
    const { ctx, client } = makeCtx([]);
    const req = stubPages(client, { value: [] });

    await listCommand.handler(ctx);

    expect(queryOf(req).get('$orderby')).toBe('receivedDateTime desc');
  });

  it('$select is present and narrow (only the fields renderMessageList needs)', async () => {
    const { ctx, client } = makeCtx([]);
    const req = stubPages(client, { value: [] });

    await listCommand.handler(ctx);

    expect(queryOf(req).get('$select')).toBe(
      'subject,from,receivedDateTime,isRead,hasAttachments,bodyPreview',
    );
  });

  it('--folder <well-known-name> resolves to /me/mailFolders/<name>/messages', async () => {
    const { ctx, client } = makeCtx(['--folder', 'drafts']);
    const req = stubPages(client, { value: [] });

    await listCommand.handler(ctx);

    expect(pathOf(req)).toBe('/me/mailFolders/drafts/messages');
  });

  it('--folder is resolved case-insensitively to the canonical lower-case well-known path', async () => {
    const { ctx, client } = makeCtx(['--folder', 'Inbox']);
    const req = stubPages(client, { value: [] });

    await listCommand.handler(ctx);

    expect(pathOf(req)).toBe('/me/mailFolders/inbox/messages');
  });

  it('--folder <raw-id> (not well-known) is passed through percent-encoded', async () => {
    const rawId = 'AAMk/odd=folder';
    const { ctx, client } = makeCtx(['--folder', rawId]);
    const req = stubPages(client, { value: [] });

    await listCommand.handler(ctx);

    expect(pathOf(req)).toBe(`/me/mailFolders/${encodeURIComponent(rawId)}/messages`);
  });

  it('never touches the network directly — only through the stubbed client method', async () => {
    const { ctx, client, fetchImpl } = makeCtx([]);
    stubPages(client, { value: [sampleMessage()] });

    await listCommand.handler(ctx);

    expect(fetchImpl).toHaveBeenCalledTimes(0);
  });
});
