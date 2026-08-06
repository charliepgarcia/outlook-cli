import { mkdtemp, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { describe, expect, it, vi } from 'vitest';
import { UsageError } from '../errors.js';
import { sendCommand } from './send.js';
import { makeCtx, text } from './testUtils.js';

describe('send', () => {
  it('builds the exact documented sendMail envelope', async () => {
    const { ctx, client } = makeCtx([
      '--to', 'a@example.com',
      '--subject', 'Hello',
      '--body', 'Body text',
    ]);
    const post = vi.spyOn(client, 'post').mockResolvedValue(undefined as never);

    await sendCommand.handler(ctx);

    expect(post).toHaveBeenCalledTimes(1);
    expect(post.mock.calls[0]?.[0]).toBe('/me/sendMail');
    expect(post.mock.calls[0]?.[1]).toEqual({
      message: {
        subject: 'Hello',
        body: { contentType: 'text', content: 'Body text' },
        toRecipients: [{ emailAddress: { address: 'a@example.com' } }],
      },
      saveToSentItems: true,
    });
  });

  it('accepts repeated --to and comma-separated --to together', async () => {
    const { ctx, client } = makeCtx([
      '--to', 'a@example.com,b@example.com',
      '--to', 'c@example.com',
      '--subject', 'S',
    ]);
    const post = vi.spyOn(client, 'post').mockResolvedValue(undefined as never);

    await sendCommand.handler(ctx);

    const payload = post.mock.calls[0]?.[1] as {
      message: { toRecipients: { emailAddress: { address: string } }[] };
    };
    expect(payload.message.toRecipients.map((r) => r.emailAddress.address)).toEqual([
      'a@example.com',
      'b@example.com',
      'c@example.com',
    ]);
  });

  it('--html sets contentType html', async () => {
    const { ctx, client } = makeCtx([
      '--to', 'a@example.com', '--subject', 'S', '--body', '<b>hi</b>', '--html',
    ]);
    const post = vi.spyOn(client, 'post').mockResolvedValue(undefined as never);

    await sendCommand.handler(ctx);

    const payload = post.mock.calls[0]?.[1] as { message: { body: { contentType: string } } };
    expect(payload.message.body.contentType).toBe('html');
  });

  it('--no-save-sent sets saveToSentItems false', async () => {
    const { ctx, client } = makeCtx([
      '--to', 'a@example.com', '--subject', 'S', '--no-save-sent',
    ]);
    const post = vi.spyOn(client, 'post').mockResolvedValue(undefined as never);

    await sendCommand.handler(ctx);

    const payload = post.mock.calls[0]?.[1] as { saveToSentItems: boolean };
    expect(payload.saveToSentItems).toBe(false);
  });

  it('reads --body-file from disk and attaches files', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'outlook-send-'));
    const bodyPath = join(dir, 'body.txt');
    const attachPath = join(dir, 'note.txt');
    await writeFile(bodyPath, 'from a file');
    await writeFile(attachPath, 'attached bytes');

    const { ctx, client } = makeCtx([
      '--to', 'a@example.com', '--subject', 'S',
      '--body-file', bodyPath, '--attach', attachPath,
    ]);
    const post = vi.spyOn(client, 'post').mockResolvedValue(undefined as never);

    await sendCommand.handler(ctx);

    const payload = post.mock.calls[0]?.[1] as {
      message: {
        body: { content: string };
        attachments: { name: string; contentBytes: string; '@odata.type': string }[];
      };
    };
    expect(payload.message.body.content).toBe('from a file');
    expect(payload.message.attachments).toHaveLength(1);
    expect(payload.message.attachments[0]?.name).toBe('note.txt');
    expect(payload.message.attachments[0]?.['@odata.type']).toBe(
      '#microsoft.graph.fileAttachment',
    );
    expect(
      Buffer.from(payload.message.attachments[0]?.contentBytes ?? '', 'base64').toString(),
    ).toBe('attached bytes');
  });

  it('--dry-run sends NOTHING and touches neither post nor the network', async () => {
    const { ctx, client, fetchImpl, out } = makeCtx([
      '--to', 'a@example.com', '--subject', 'S', '--body', 'B', '--dry-run',
    ]);
    const post = vi.spyOn(client, 'post');

    await sendCommand.handler(ctx);

    expect(post).toHaveBeenCalledTimes(0);
    expect(fetchImpl).toHaveBeenCalledTimes(0);
    expect(text(out)).toContain('DRY RUN');
    expect(text(out)).toContain('/me/sendMail');
  });

  it('requires at least one recipient', async () => {
    const { ctx } = makeCtx(['--subject', 'S', '--body', 'B']);
    await expect(sendCommand.handler(ctx)).rejects.toBeInstanceOf(UsageError);
  });

  it('requires a subject', async () => {
    const { ctx } = makeCtx(['--to', 'a@example.com', '--body', 'B']);
    await expect(sendCommand.handler(ctx)).rejects.toBeInstanceOf(UsageError);
  });

  it('rejects a malformed recipient rather than sending to it', async () => {
    const { ctx } = makeCtx(['--to', 'not-an-address', '--subject', 'S']);
    await expect(sendCommand.handler(ctx)).rejects.toBeInstanceOf(UsageError);
  });

  it('emits parseable JSON under --json', async () => {
    const { ctx, client, out } = makeCtx([
      '--to', 'a@example.com', '--subject', 'S', '--json',
    ]);
    vi.spyOn(client, 'post').mockResolvedValue(undefined as never);

    await sendCommand.handler(ctx);

    expect(() => JSON.parse(text(out))).not.toThrow();
  });
});
