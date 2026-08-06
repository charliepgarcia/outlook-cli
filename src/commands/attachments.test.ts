/**
 * Tests for `outlook attachments`. Mirrors get.test.ts's real-GraphClient
 * shape (makeCtx), but the network boundary here is `client.request`, NOT
 * `client.get`: collectItems/pageItems (paginate.ts) build a RequestSpec via
 * client.buildSpec and hand it to client.request directly. Stubbing `get`
 * would leave `request` live and trip testUtils' deliberately-throwing
 * fetchImpl with "NETWORK CALLED" — this cost a sibling agent 8 failing
 * tests, per the task brief. `stubPages`/`pathOf` below are copied from the
 * working idiom in folders.test.ts.
 *
 * Disk tests write ONLY inside an OS temp dir (node:os tmpdir + mkdtemp),
 * tracked and removed in afterEach so a failing assertion never leaks a
 * directory.
 */

import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterEach, describe, expect, it, vi } from 'vitest';
import type { GraphClient } from '../client.js';
import { UsageError } from '../errors.js';
import { GRAPH_BASE, type Attachment, type GraphCollection, type RequestSpec } from '../types.js';
import { attachmentsCommand } from './attachments.js';
import { makeCtx, text } from './testUtils.js';

function stubPages(client: GraphClient, ...pages: GraphCollection<Attachment>[]) {
  const spy = vi.spyOn(client, 'request');
  for (const page of pages) spy.mockResolvedValueOnce(page as never);
  return spy;
}

function pathOf(spy: { mock: { calls: unknown[][] } }, call = 0): string {
  const spec = spy.mock.calls[call]?.[0] as RequestSpec | undefined;
  if (spec === undefined) return '';
  const url = spec.url.startsWith(GRAPH_BASE) ? spec.url.slice(GRAPH_BASE.length) : spec.url;
  return url.split('?')[0] ?? '';
}

function sampleAttachment(overrides: Partial<Attachment> = {}): Attachment {
  return {
    id: 'AAMkAttachment1',
    name: 'notes.txt',
    contentType: 'text/plain',
    size: 11,
    isInline: false,
    '@odata.type': '#microsoft.graph.fileAttachment',
    contentBytes: Buffer.from('hello world').toString('base64'),
    ...overrides,
  };
}

const tempDirs: string[] = [];

async function makeTempDir(): Promise<string> {
  const dir = await mkdtemp(join(tmpdir(), 'outlook-attachments-'));
  tempDirs.push(dir);
  return dir;
}

afterEach(async () => {
  await Promise.all(tempDirs.splice(0).map((d) => rm(d, { recursive: true, force: true })));
});

describe('attachments list', () => {
  it('GETs the documented attachments endpoint and renders human-readable output', async () => {
    const { ctx, client, out } = makeCtx(['list', 'msg-1']);
    const attachments = [sampleAttachment(), sampleAttachment({ id: 'a2', name: 'photo.png' })];
    const req = stubPages(client, { value: attachments });

    await attachmentsCommand.handler(ctx);

    expect(req).toHaveBeenCalledTimes(1);
    expect(pathOf(req)).toBe('/me/messages/msg-1/attachments');
    expect(text(out)).toContain('notes.txt');
    expect(text(out)).toContain('photo.png');
  });

  it('defaults to list when no subcommand keyword is given (bare message id)', async () => {
    const { ctx, client, out } = makeCtx(['msg-1']);
    const req = stubPages(client, { value: [sampleAttachment()] });

    await attachmentsCommand.handler(ctx);

    expect(pathOf(req)).toBe('/me/messages/msg-1/attachments');
    expect(text(out)).toContain('notes.txt');
  });

  it('--json emits the raw attachment array, unmodified', async () => {
    const { ctx, client, out } = makeCtx(['list', 'msg-1', '--json']);
    const attachments = [sampleAttachment()];
    stubPages(client, { value: attachments });

    await attachmentsCommand.handler(ctx);

    expect(JSON.parse(text(out))).toEqual(attachments);
  });

  it('missing message id throws a clean UsageError via requirePositional', async () => {
    const { ctx, client } = makeCtx(['list']);
    const req = stubPages(client, { value: [] });

    await expect(attachmentsCommand.handler(ctx)).rejects.toBeInstanceOf(UsageError);
    expect(req).toHaveBeenCalledTimes(0);
  });

  it('never touches the network directly — only through the stubbed client method', async () => {
    const { ctx, client, fetchImpl } = makeCtx(['list', 'msg-1']);
    stubPages(client, { value: [sampleAttachment()] });

    await attachmentsCommand.handler(ctx);

    expect(fetchImpl).toHaveBeenCalledTimes(0);
  });
});

describe('attachments download', () => {
  it('decodes base64 contentBytes and writes the correct bytes to disk', async () => {
    const dir = await makeTempDir();
    const attachment = sampleAttachment({ contentBytes: Buffer.from('hello world').toString('base64') });
    const { ctx, client, out } = makeCtx(['download', 'msg-1', attachment.id, '--out', dir]);
    stubPages(client, { value: [attachment] });

    await attachmentsCommand.handler(ctx);

    const written = await readFile(join(dir, attachment.name), 'utf8');
    expect(written).toBe('hello world');
    expect(text(out)).toContain(attachment.name);
  });

  it('refuses to silently overwrite an existing file without --force', async () => {
    const dir = await makeTempDir();
    const attachment = sampleAttachment();
    await writeFile(join(dir, attachment.name), 'pre-existing content');
    const { ctx, client } = makeCtx(['download', 'msg-1', attachment.id, '--out', dir]);
    stubPages(client, { value: [attachment] });

    await expect(attachmentsCommand.handler(ctx)).rejects.toBeInstanceOf(UsageError);
    expect(await readFile(join(dir, attachment.name), 'utf8')).toBe('pre-existing content');
  });

  it('--force overwrites an existing file', async () => {
    const dir = await makeTempDir();
    const attachment = sampleAttachment({ contentBytes: Buffer.from('new content').toString('base64') });
    await writeFile(join(dir, attachment.name), 'pre-existing content');
    const { ctx, client } = makeCtx(['download', 'msg-1', attachment.id, '--out', dir, '--force']);
    stubPages(client, { value: [attachment] });

    await attachmentsCommand.handler(ctx);

    expect(await readFile(join(dir, attachment.name), 'utf8')).toBe('new content');
  });

  it('sanitizes a path-traversal attachment name so the write stays inside --out', async () => {
    const dir = await makeTempDir();
    const attachment = sampleAttachment({
      name: '../../evil.txt',
      contentBytes: Buffer.from('payload').toString('base64'),
    });
    const { ctx, client } = makeCtx(['download', 'msg-1', attachment.id, '--out', dir]);
    stubPages(client, { value: [attachment] });

    await attachmentsCommand.handler(ctx);

    // Lands inside dir under a basenamed filename, never above it.
    expect(await readFile(join(dir, 'evil.txt'), 'utf8')).toBe('payload');
    await expect(readFile(join(dir, '..', '..', 'evil.txt'), 'utf8')).rejects.toThrow();
  });

  it('rejects itemAttachment with a clear UsageError naming the type', async () => {
    const dir = await makeTempDir();
    const attachment = sampleAttachment({
      '@odata.type': '#microsoft.graph.itemAttachment',
      contentBytes: undefined,
    });
    const { ctx, client } = makeCtx(['download', 'msg-1', attachment.id, '--out', dir]);
    stubPages(client, { value: [attachment] });

    await expect(attachmentsCommand.handler(ctx)).rejects.toThrow(/itemAttachment|item has no downloadable/);
  });

  it('--all downloads every attachment on the message', async () => {
    const dir = await makeTempDir();
    const a1 = sampleAttachment({ id: 'a1', name: 'one.txt', contentBytes: Buffer.from('one').toString('base64') });
    const a2 = sampleAttachment({ id: 'a2', name: 'two.txt', contentBytes: Buffer.from('two').toString('base64') });
    const { ctx, client } = makeCtx(['download', 'msg-1', '--all', '--out', dir]);
    stubPages(client, { value: [a1, a2] });

    await attachmentsCommand.handler(ctx);

    expect(await readFile(join(dir, 'one.txt'), 'utf8')).toBe('one');
    expect(await readFile(join(dir, 'two.txt'), 'utf8')).toBe('two');
  });

  it('without an attachment id or --all, throws a clean UsageError before any network call', async () => {
    const dir = await makeTempDir();
    const { ctx, client } = makeCtx(['download', 'msg-1', '--out', dir]);
    const req = stubPages(client, { value: [sampleAttachment()] });

    await expect(attachmentsCommand.handler(ctx)).rejects.toBeInstanceOf(UsageError);
    expect(req).toHaveBeenCalledTimes(0);
  });

  it('missing message id throws a clean UsageError via requirePositional', async () => {
    const { ctx, client } = makeCtx(['download']);
    const req = stubPages(client, { value: [] });

    await expect(attachmentsCommand.handler(ctx)).rejects.toBeInstanceOf(UsageError);
    expect(req).toHaveBeenCalledTimes(0);
  });

  it('never touches the network directly — only through the stubbed client method', async () => {
    const dir = await makeTempDir();
    const attachment = sampleAttachment();
    const { ctx, client, fetchImpl } = makeCtx(['download', 'msg-1', attachment.id, '--out', dir]);
    stubPages(client, { value: [attachment] });

    await attachmentsCommand.handler(ctx);

    expect(fetchImpl).toHaveBeenCalledTimes(0);
  });
});
