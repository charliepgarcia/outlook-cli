import { describe, expect, it, vi } from 'vitest';
import { UsageError } from '../errors.js';
import { deleteCommand, moveCommand, purgeCommand } from './move.js';
import { makeCtx, text } from './testUtils.js';

describe('delete (safe default — recoverable move)', () => {
  it('performs POST /me/messages/{id}/move with destinationId "deleteditems" — never a raw DELETE', async () => {
    const { ctx, client } = makeCtx(['MSG1']);
    const post = vi.spyOn(client, 'post').mockResolvedValue({ id: 'NEWID1' } as never);
    const del = vi.spyOn(client, 'del');

    await deleteCommand.handler(ctx);

    expect(post).toHaveBeenCalledTimes(1);
    expect(post.mock.calls[0]?.[0]).toBe('/me/messages/MSG1/move');
    expect(post.mock.calls[0]?.[1]).toEqual({ destinationId: 'deleteditems' });
    expect(del).toHaveBeenCalledTimes(0);
  });

  it('--dry-run sends nothing and reaches no network', async () => {
    const { ctx, client, fetchImpl, out } = makeCtx(['MSG1', '--dry-run']);
    const post = vi.spyOn(client, 'post');
    const del = vi.spyOn(client, 'del');

    await deleteCommand.handler(ctx);

    expect(post).toHaveBeenCalledTimes(0);
    expect(del).toHaveBeenCalledTimes(0);
    expect(fetchImpl).toHaveBeenCalledTimes(0);
    expect(text(out)).toContain('DRY RUN');
    expect(text(out)).toContain('/me/messages/MSG1/move');
    expect(text(out)).toContain('deleteditems');
  });

  it('requires a message id', async () => {
    const { ctx } = makeCtx([]);
    await expect(deleteCommand.handler(ctx)).rejects.toBeInstanceOf(UsageError);
  });
});

describe('move (arbitrary destination folder)', () => {
  it('normalizes a well-known folder name to lowercase', async () => {
    const { ctx, client } = makeCtx(['MSG1', '--folder', 'Archive']);
    const post = vi.spyOn(client, 'post').mockResolvedValue({ id: 'NEWID2' } as never);

    await moveCommand.handler(ctx);

    expect(post.mock.calls[0]?.[0]).toBe('/me/messages/MSG1/move');
    expect(post.mock.calls[0]?.[1]).toEqual({ destinationId: 'archive' });
  });

  it('preserves a raw (non-well-known) folder id verbatim, case included', async () => {
    const { ctx, client } = makeCtx(['MSG1', '--folder', 'AAMkAGI2TG93AAA=']);
    const post = vi.spyOn(client, 'post').mockResolvedValue({ id: 'NEWID3' } as never);

    await moveCommand.handler(ctx);

    expect(post.mock.calls[0]?.[1]).toEqual({ destinationId: 'AAMkAGI2TG93AAA=' });
  });

  it('requires --folder', async () => {
    const { ctx } = makeCtx(['MSG1']);
    await expect(moveCommand.handler(ctx)).rejects.toBeInstanceOf(UsageError);
  });

  it('requires a message id', async () => {
    const { ctx } = makeCtx(['--folder', 'inbox']);
    await expect(moveCommand.handler(ctx)).rejects.toBeInstanceOf(UsageError);
  });

  it('--dry-run sends nothing and reaches no network', async () => {
    const { ctx, client, fetchImpl, out } = makeCtx(['MSG1', '--folder', 'inbox', '--dry-run']);
    const post = vi.spyOn(client, 'post');

    await moveCommand.handler(ctx);

    expect(post).toHaveBeenCalledTimes(0);
    expect(fetchImpl).toHaveBeenCalledTimes(0);
    expect(text(out)).toContain('DRY RUN');
    expect(text(out)).toContain('/me/messages/MSG1/move');
    expect(text(out)).toContain('destinationId');
    expect(text(out)).toContain('inbox');
  });

  it('--json emits parseable JSON describing the move', async () => {
    const { ctx, client, out } = makeCtx(['MSG1', '--folder', 'inbox', '--json']);
    vi.spyOn(client, 'post').mockResolvedValue({ id: 'NEWID4' } as never);

    await moveCommand.handler(ctx);

    const parsed = JSON.parse(text(out));
    expect(parsed.movedFrom).toBe('MSG1');
    expect(parsed.destinationId).toBe('inbox');
  });
});

describe('purge (raw DELETE, requires --yes)', () => {
  it('refuses to proceed without --yes', async () => {
    const { ctx, client } = makeCtx(['MSG1']);
    const del = vi.spyOn(client, 'del');

    await expect(purgeCommand.handler(ctx)).rejects.toBeInstanceOf(UsageError);
    expect(del).toHaveBeenCalledTimes(0);
  });

  it('the --yes gate applies even under --dry-run (no silent preview bypass)', async () => {
    const { ctx, fetchImpl } = makeCtx(['MSG1', '--dry-run']);

    await expect(purgeCommand.handler(ctx)).rejects.toBeInstanceOf(UsageError);
    expect(fetchImpl).toHaveBeenCalledTimes(0);
  });

  it('requires a message id', async () => {
    const { ctx } = makeCtx(['--yes']);
    await expect(purgeCommand.handler(ctx)).rejects.toBeInstanceOf(UsageError);
  });

  it('--yes issues a raw DELETE /me/messages/{id} — never a move', async () => {
    const { ctx, client } = makeCtx(['MSG1', '--yes']);
    const del = vi.spyOn(client, 'del').mockResolvedValue(undefined as never);
    const post = vi.spyOn(client, 'post');

    await purgeCommand.handler(ctx);

    expect(del).toHaveBeenCalledTimes(1);
    expect(del.mock.calls[0]?.[0]).toBe('/me/messages/MSG1');
    expect(post).toHaveBeenCalledTimes(0);
  });

  it('--yes --dry-run sends nothing and reaches no network', async () => {
    const { ctx, client, fetchImpl, out } = makeCtx(['MSG1', '--yes', '--dry-run']);
    const del = vi.spyOn(client, 'del');

    await purgeCommand.handler(ctx);

    expect(del).toHaveBeenCalledTimes(0);
    expect(fetchImpl).toHaveBeenCalledTimes(0);
    expect(text(out)).toContain('DRY RUN');
    expect(text(out)).toContain('DELETE');
    expect(text(out)).toContain('/me/messages/MSG1');
  });

  it('is a distinct command from delete — delete needs no --yes at all', async () => {
    expect(purgeCommand.name).toBe('purge');
    expect(deleteCommand.name).toBe('delete');
    expect(purgeCommand.name).not.toBe(deleteCommand.name);

    const { ctx, client } = makeCtx(['MSG1']);
    vi.spyOn(client, 'post').mockResolvedValue({ id: 'NEWID5' } as never);
    await expect(deleteCommand.handler(ctx)).resolves.toBeUndefined();
  });
});
