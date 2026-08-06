import { describe, expect, it, vi } from 'vitest';
import { UsageError } from '../errors.js';
import { forwardCommand, replyAllCommand, replyCommand } from './reply.js';
import { makeCtx, text } from './testUtils.js';

describe('reply / reply-all / forward', () => {
  it('reply posts to the /reply action with a comment body', async () => {
    const { ctx, client } = makeCtx(['MSG123', '--body', 'Thanks!']);
    const post = vi.spyOn(client, 'post').mockResolvedValue(undefined as never);

    await replyCommand.handler(ctx);

    expect(post).toHaveBeenCalledTimes(1);
    expect(post.mock.calls[0]?.[0]).toBe('/me/messages/MSG123/reply');
    expect(post.mock.calls[0]?.[1]).toEqual({ comment: 'Thanks!' });
  });

  it('reply-all posts to the /replyAll action', async () => {
    const { ctx, client } = makeCtx(['MSG123', '--body', 'All of you']);
    const post = vi.spyOn(client, 'post').mockResolvedValue(undefined as never);

    await replyAllCommand.handler(ctx);

    expect(post.mock.calls[0]?.[0]).toBe('/me/messages/MSG123/replyAll');
    expect(post.mock.calls[0]?.[1]).toEqual({ comment: 'All of you' });
  });

  it('forward posts comment plus toRecipients', async () => {
    const { ctx, client } = makeCtx([
      'MSG123', '--to', 'x@example.com,y@example.com', '--body', 'FYI',
    ]);
    const post = vi.spyOn(client, 'post').mockResolvedValue(undefined as never);

    await forwardCommand.handler(ctx);

    expect(post.mock.calls[0]?.[0]).toBe('/me/messages/MSG123/forward');
    expect(post.mock.calls[0]?.[1]).toEqual({
      comment: 'FYI',
      toRecipients: [
        { emailAddress: { address: 'x@example.com' } },
        { emailAddress: { address: 'y@example.com' } },
      ],
    });
  });

  it('forward without --to is a usage error, not an empty send', async () => {
    const { ctx } = makeCtx(['MSG123', '--body', 'FYI']);
    await expect(forwardCommand.handler(ctx)).rejects.toBeInstanceOf(UsageError);
  });

  it('a missing messageId is a usage error', async () => {
    const { ctx } = makeCtx(['--body', 'hi']);
    await expect(replyCommand.handler(ctx)).rejects.toBeInstanceOf(UsageError);
  });

  it('defaults comment to empty string when --body is omitted', async () => {
    const { ctx, client } = makeCtx(['MSG123']);
    const post = vi.spyOn(client, 'post').mockResolvedValue(undefined as never);

    await replyCommand.handler(ctx);

    expect(post.mock.calls[0]?.[1]).toEqual({ comment: '' });
  });

  it('url-encodes an id containing unsafe characters', async () => {
    const { ctx, client } = makeCtx(['AAMk/id+with=chars', '--body', 'x']);
    const post = vi.spyOn(client, 'post').mockResolvedValue(undefined as never);

    await replyCommand.handler(ctx);

    expect(post.mock.calls[0]?.[0]).toBe(
      `/me/messages/${encodeURIComponent('AAMk/id+with=chars')}/reply`,
    );
  });

  describe('--dry-run leaves the network untouched', () => {
    for (const [label, command, argv] of [
      ['reply', replyCommand, ['MSG1', '--body', 'b', '--dry-run']],
      ['reply-all', replyAllCommand, ['MSG1', '--body', 'b', '--dry-run']],
      [
        'forward',
        forwardCommand,
        ['MSG1', '--to', 'z@example.com', '--body', 'b', '--dry-run'],
      ],
    ] as const) {
      it(label, async () => {
        const { ctx, client, fetchImpl, out } = makeCtx([...argv]);
        const post = vi.spyOn(client, 'post');

        await command.handler(ctx);

        expect(post).toHaveBeenCalledTimes(0);
        expect(fetchImpl).toHaveBeenCalledTimes(0);
        expect(text(out)).toContain('DRY RUN');
      });
    }
  });
});
