import { describe, expect, it, vi } from 'vitest';
import { UsageError } from '../errors.js';
import type { Message } from '../types.js';
import { draftCommand } from './draft.js';
import { makeCtx, text } from './testUtils.js';

const DRAFT: Message = { id: 'DRAFT1', subject: 'Hello', isDraft: true };

describe('draft', () => {
  it('list reads the drafts folder', async () => {
    const { ctx, client } = makeCtx(['list']);
    const request = vi
      .spyOn(client, 'request')
      .mockResolvedValue({ value: [DRAFT] } as never);

    await draftCommand.handler(ctx);

    const spec = request.mock.calls[0]?.[0] as { url: string };
    expect(spec.url).toContain('/me/mailFolders/drafts/messages');
  });

  it('create posts a composed message to /me/messages', async () => {
    const { ctx, client, out } = makeCtx([
      'create', '--subject', 'Draft subject', '--to', 'a@example.com', '--body', 'B',
    ]);
    const post = vi.spyOn(client, 'post').mockResolvedValue(DRAFT as never);

    await draftCommand.handler(ctx);

    expect(post.mock.calls[0]?.[0]).toBe('/me/messages');
    expect(post.mock.calls[0]?.[1]).toEqual({
      subject: 'Draft subject',
      body: { contentType: 'text', content: 'B' },
      toRecipients: [{ emailAddress: { address: 'a@example.com' } }],
    });
    expect(text(out)).toContain('DRAFT1');
  });

  it('create with no content at all is a usage error', async () => {
    const { ctx } = makeCtx(['create']);
    await expect(draftCommand.handler(ctx)).rejects.toBeInstanceOf(UsageError);
  });

  it('update patches the named draft', async () => {
    const { ctx, client } = makeCtx(['update', 'DRAFT1', '--subject', 'New']);
    const patch = vi.spyOn(client, 'patch').mockResolvedValue(DRAFT as never);

    await draftCommand.handler(ctx);

    expect(patch.mock.calls[0]?.[0]).toBe('/me/messages/DRAFT1');
    expect(patch.mock.calls[0]?.[1]).toEqual({ subject: 'New' });
  });

  it('update without an id is a usage error', async () => {
    const { ctx } = makeCtx(['update', '--subject', 'New']);
    await expect(draftCommand.handler(ctx)).rejects.toBeInstanceOf(UsageError);
  });

  it('send posts to the draft send action', async () => {
    const { ctx, client, out } = makeCtx(['send', 'DRAFT1']);
    const post = vi.spyOn(client, 'post').mockResolvedValue(undefined as never);

    await draftCommand.handler(ctx);

    expect(post.mock.calls[0]?.[0]).toBe('/me/messages/DRAFT1/send');
    expect(text(out)).toContain('sent');
  });

  it('send --dry-run does not send the draft', async () => {
    const { ctx, client, fetchImpl, out } = makeCtx(['send', 'DRAFT1', '--dry-run']);
    const post = vi.spyOn(client, 'post');

    await draftCommand.handler(ctx);

    expect(post).toHaveBeenCalledTimes(0);
    expect(fetchImpl).toHaveBeenCalledTimes(0);
    expect(text(out)).toContain('DRY RUN');
    expect(text(out)).toContain('/me/messages/DRAFT1/send');
  });

  it('create --dry-run does not create', async () => {
    const { ctx, client } = makeCtx(['create', '--subject', 'S', '--dry-run']);
    const post = vi.spyOn(client, 'post');

    await draftCommand.handler(ctx);

    expect(post).toHaveBeenCalledTimes(0);
  });

  it('an unknown subcommand lists the valid ones', async () => {
    const { ctx } = makeCtx(['frobnicate']);
    await expect(draftCommand.handler(ctx)).rejects.toThrow(/list, create, update, send/);
  });

  it('no subcommand at all is a usage error', async () => {
    const { ctx } = makeCtx([]);
    await expect(draftCommand.handler(ctx)).rejects.toBeInstanceOf(UsageError);
  });
});
