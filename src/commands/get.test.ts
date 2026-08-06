/**
 * Tests for `outlook get`. Mirrors send.test.ts's shape: a real GraphClient
 * from makeCtx (so buildSpec/URL-joining behave exactly as in production),
 * with the network boundary (`client.get`) stubbed via vi.spyOn. makeCtx's
 * injected fetch THROWS on purpose, so any leak past the stub fails loudly.
 */

import { describe, expect, it, vi } from 'vitest';
import { UsageError } from '../errors.js';
import type { Message } from '../types.js';
import { getCommand } from './get.js';
import { makeCtx, text } from './testUtils.js';

function sampleMessage(overrides: Partial<Message> = {}): Message {
  return {
    id: 'AAMkADhMGAAA=',
    subject: 'Meet for lunch?',
    from: { emailAddress: { address: 'frannis@contoso.com', name: 'Frannis' } },
    toRecipients: [{ emailAddress: { address: 'me@contoso.com' } }],
    receivedDateTime: '2024-06-01T12:00:00Z',
    importance: 'normal',
    isRead: true,
    hasAttachments: false,
    body: { contentType: 'text', content: 'The new cafeteria is open.' },
    ...overrides,
  };
}

describe('get', () => {
  it('fetches the message by id and renders it as human-readable text', async () => {
    const { ctx, client, out } = makeCtx(['AAMkADhMGAAA=']);
    const message = sampleMessage();
    const get = vi.spyOn(client, 'get').mockResolvedValue(message);

    await getCommand.handler(ctx);

    expect(get).toHaveBeenCalledTimes(1);
    expect(text(out)).toContain('Subject: Meet for lunch?');
    expect(text(out)).toContain('Frannis <frannis@contoso.com>');
  });

  it('requires a message id and throws a clean UsageError when absent', async () => {
    const { ctx, client } = makeCtx([]);
    const get = vi.spyOn(client, 'get');

    await expect(getCommand.handler(ctx)).rejects.toBeInstanceOf(UsageError);
    expect(get).toHaveBeenCalledTimes(0);
  });

  it('builds the exact documented get-message endpoint, percent-encoding the id', async () => {
    const rawId = 'AAMk/AGVmMDEzMTM4=+odd';
    const { ctx, client } = makeCtx([rawId]);
    const get = vi.spyOn(client, 'get').mockResolvedValue(sampleMessage({ id: rawId }));

    await getCommand.handler(ctx);

    expect(get.mock.calls[0]?.[0]).toBe(`/me/messages/${encodeURIComponent(rawId)}`);
    // No $select/$expand invented for attachment metadata — docs/01 documents
    // attachments only as a separate collection endpoint (§7), never inlined
    // into message-get via query params.
    expect(get.mock.calls[0]?.[1]).toBeUndefined();
  });

  it('--html renders the raw HTML body instead of converting it to text', async () => {
    const { ctx, client, out } = makeCtx(['msg-html', '--html']);
    const message = sampleMessage({
      body: { contentType: 'html', content: '<p>Hello <b>world</b></p>' },
    });
    vi.spyOn(client, 'get').mockResolvedValue(message);

    await getCommand.handler(ctx);

    expect(text(out)).toContain('<p>Hello <b>world</b></p>');
  });

  it('without --html, an HTML body is converted to plain text', async () => {
    // No `from`/`to` here (unlike sampleMessage()) so the rendered header
    // block contains no incidental "<addr>" angle brackets of its own —
    // keeps the "no HTML tags leaked" assertion below unambiguous.
    const { ctx, client, out } = makeCtx(['msg-html']);
    const message: Message = {
      id: 'msg-html',
      subject: 'HTML body',
      body: { contentType: 'html', content: '<p>Hello <b>world</b></p>' },
    };
    vi.spyOn(client, 'get').mockResolvedValue(message);

    await getCommand.handler(ctx);

    expect(text(out)).not.toMatch(/<[a-zA-Z/][^>]*>/);
    expect(text(out)).toContain('Hello');
    expect(text(out)).toContain('world');
  });

  it('--json emits the raw message object, unmodified', async () => {
    const { ctx, client, out } = makeCtx(['msg-1', '--json']);
    const message = sampleMessage();
    vi.spyOn(client, 'get').mockResolvedValue(message);

    await getCommand.handler(ctx);

    expect(JSON.parse(text(out))).toEqual(message);
  });

  it('without --json, output is human text, not a JSON blob', async () => {
    const { ctx, client, out } = makeCtx(['msg-1']);
    vi.spyOn(client, 'get').mockResolvedValue(sampleMessage());

    await getCommand.handler(ctx);

    expect(() => JSON.parse(text(out))).toThrow();
  });

  it('never touches the network directly — only through the stubbed client method', async () => {
    const { ctx, client, fetchImpl } = makeCtx(['msg-1']);
    vi.spyOn(client, 'get').mockResolvedValue(sampleMessage());

    await getCommand.handler(ctx);

    expect(fetchImpl).toHaveBeenCalledTimes(0);
  });
});
