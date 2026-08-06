import { describe, expect, it, vi } from 'vitest';
import { UsageError } from '../errors.js';
import { flagCommand, markReadCommand, markUnreadCommand, unflagCommand } from './flags.js';
import { makeCtx, text } from './testUtils.js';

describe('mark-read / mark-unread (PATCH isRead)', () => {
  it('mark-read issues PATCH /me/messages/{id} with body { isRead: true }', async () => {
    const { ctx, client } = makeCtx(['MSG1']);
    const patch = vi.spyOn(client, 'patch').mockResolvedValue({ id: 'MSG1', isRead: true } as never);
    const post = vi.spyOn(client, 'post');
    const del = vi.spyOn(client, 'del');

    await markReadCommand.handler(ctx);

    expect(patch).toHaveBeenCalledTimes(1);
    expect(patch.mock.calls[0]?.[0]).toBe('/me/messages/MSG1');
    expect(patch.mock.calls[0]?.[1]).toEqual({ isRead: true });
    expect(post).toHaveBeenCalledTimes(0);
    expect(del).toHaveBeenCalledTimes(0);
  });

  it('mark-unread issues PATCH /me/messages/{id} with body { isRead: false }', async () => {
    const { ctx, client } = makeCtx(['MSG1']);
    const patch = vi.spyOn(client, 'patch').mockResolvedValue({ id: 'MSG1', isRead: false } as never);

    await markUnreadCommand.handler(ctx);

    expect(patch).toHaveBeenCalledTimes(1);
    expect(patch.mock.calls[0]?.[0]).toBe('/me/messages/MSG1');
    expect(patch.mock.calls[0]?.[1]).toEqual({ isRead: false });
  });

  it('mark-read requires a message id', async () => {
    const { ctx } = makeCtx([]);
    await expect(markReadCommand.handler(ctx)).rejects.toBeInstanceOf(UsageError);
  });

  it('mark-read --dry-run sends nothing and reaches no network', async () => {
    const { ctx, client, fetchImpl, out } = makeCtx(['MSG1', '--dry-run']);
    const patch = vi.spyOn(client, 'patch');

    await markReadCommand.handler(ctx);

    expect(patch).toHaveBeenCalledTimes(0);
    expect(fetchImpl).toHaveBeenCalledTimes(0);
    expect(text(out)).toContain('DRY RUN');
    expect(text(out)).toContain('/me/messages/MSG1');
  });

  it('mark-unread --dry-run sends nothing and reaches no network', async () => {
    const { ctx, client, fetchImpl, out } = makeCtx(['MSG1', '--dry-run']);
    const patch = vi.spyOn(client, 'patch');

    await markUnreadCommand.handler(ctx);

    expect(patch).toHaveBeenCalledTimes(0);
    expect(fetchImpl).toHaveBeenCalledTimes(0);
    expect(text(out)).toContain('DRY RUN');
    expect(text(out)).toContain('/me/messages/MSG1');
  });

  it('--json emits parseable JSON with id and the updated message', async () => {
    const { ctx, client, out } = makeCtx(['MSG1', '--json']);
    vi.spyOn(client, 'patch').mockResolvedValue({ id: 'MSG1', isRead: true } as never);

    await markReadCommand.handler(ctx);

    const parsed = JSON.parse(text(out));
    expect(parsed.id).toBe('MSG1');
    expect(parsed.message).toEqual({ id: 'MSG1', isRead: true });
  });
});

describe('flag / unflag (PATCH flag.flagStatus)', () => {
  it('flag issues PATCH with body { flag: { flagStatus: "flagged" } }', async () => {
    const { ctx, client } = makeCtx(['MSG1']);
    const patch = vi.spyOn(client, 'patch').mockResolvedValue({ id: 'MSG1' } as never);
    const post = vi.spyOn(client, 'post');
    const del = vi.spyOn(client, 'del');

    await flagCommand.handler(ctx);

    expect(patch).toHaveBeenCalledTimes(1);
    expect(patch.mock.calls[0]?.[0]).toBe('/me/messages/MSG1');
    expect(patch.mock.calls[0]?.[1]).toEqual({ flag: { flagStatus: 'flagged' } });
    expect(post).toHaveBeenCalledTimes(0);
    expect(del).toHaveBeenCalledTimes(0);
  });

  it('unflag issues PATCH with body { flag: { flagStatus: "notFlagged" } }', async () => {
    const { ctx, client } = makeCtx(['MSG1']);
    const patch = vi.spyOn(client, 'patch').mockResolvedValue({ id: 'MSG1' } as never);

    await unflagCommand.handler(ctx);

    expect(patch).toHaveBeenCalledTimes(1);
    expect(patch.mock.calls[0]?.[0]).toBe('/me/messages/MSG1');
    expect(patch.mock.calls[0]?.[1]).toEqual({ flag: { flagStatus: 'notFlagged' } });
  });

  it('flag requires a message id', async () => {
    const { ctx } = makeCtx([]);
    await expect(flagCommand.handler(ctx)).rejects.toBeInstanceOf(UsageError);
  });

  it('unflag requires a message id', async () => {
    const { ctx } = makeCtx([]);
    await expect(unflagCommand.handler(ctx)).rejects.toBeInstanceOf(UsageError);
  });

  it('flag --dry-run sends nothing and reaches no network', async () => {
    const { ctx, client, fetchImpl, out } = makeCtx(['MSG1', '--dry-run']);
    const patch = vi.spyOn(client, 'patch');

    await flagCommand.handler(ctx);

    expect(patch).toHaveBeenCalledTimes(0);
    expect(fetchImpl).toHaveBeenCalledTimes(0);
    expect(text(out)).toContain('DRY RUN');
    expect(text(out)).toContain('/me/messages/MSG1');
  });

  it('unflag --dry-run sends nothing and reaches no network', async () => {
    const { ctx, client, fetchImpl, out } = makeCtx(['MSG1', '--dry-run']);
    const patch = vi.spyOn(client, 'patch');

    await unflagCommand.handler(ctx);

    expect(patch).toHaveBeenCalledTimes(0);
    expect(fetchImpl).toHaveBeenCalledTimes(0);
    expect(text(out)).toContain('DRY RUN');
    expect(text(out)).toContain('/me/messages/MSG1');
  });
});

describe('message-id percent-encoding', () => {
  it('encodes special characters in the message id path segment (mark-read)', async () => {
    const rawId = 'AAMk/id+with=chars';
    const { ctx, client } = makeCtx([rawId]);
    const patch = vi.spyOn(client, 'patch').mockResolvedValue({ id: rawId } as never);

    await markReadCommand.handler(ctx);

    expect(patch.mock.calls[0]?.[0]).toBe(`/me/messages/${encodeURIComponent(rawId)}`);
  });

  it('encodes special characters in the message id path segment (flag)', async () => {
    const rawId = 'AAMk/id+with=chars';
    const { ctx, client } = makeCtx([rawId]);
    const patch = vi.spyOn(client, 'patch').mockResolvedValue({ id: rawId } as never);

    await flagCommand.handler(ctx);

    expect(patch.mock.calls[0]?.[0]).toBe(`/me/messages/${encodeURIComponent(rawId)}`);
  });
});

describe('command metadata', () => {
  it('all four commands are marked mutating with distinct names', () => {
    expect(markReadCommand.mutating).toBe(true);
    expect(markUnreadCommand.mutating).toBe(true);
    expect(flagCommand.mutating).toBe(true);
    expect(unflagCommand.mutating).toBe(true);
    const names = [markReadCommand.name, markUnreadCommand.name, flagCommand.name, unflagCommand.name];
    expect(new Set(names).size).toBe(4);
  });
});
