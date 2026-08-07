/**
 * `outlook attachments download` — batch destination allocation.
 *
 * Split out of attachments.download.test.ts to stay under the 350-LOC ceiling.
 * Scaffolding (stubPages / sampleAttachment / makeTempDir) mirrors that file;
 * see its header for why the network boundary stubbed here is `client.request`
 * rather than `client.get`.
 */

import { mkdtemp, readdir, readFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterEach, describe, expect, it, vi } from 'vitest';
import type { GraphClient } from '../client.js';
import type { Attachment, GraphCollection } from '../types.js';
import { attachmentsCommand } from './attachments.js';
import { makeCtx, text } from './testUtils.js';

function stubPages(client: GraphClient, ...pages: GraphCollection<Attachment>[]) {
  const spy = vi.spyOn(client, 'request');
  for (const page of pages) spy.mockResolvedValueOnce(page as never);
  return spy;
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
  const dir = await mkdtemp(join(tmpdir(), 'outlook-attach-dest-'));
  tempDirs.push(dir);
  return dir;
}

afterEach(async () => {
  await Promise.all(tempDirs.splice(0).map((d) => rm(d, { recursive: true, force: true })));
});

/**
 * Batch destination allocation (review finding F-A2).
 *
 * sanitizeFilename is deliberately many-to-one — `a:b.txt` and `a?b.txt` both
 * land on `a_b.txt` — and Outlook routinely puts two attachments named
 * `image001.png` on one message. Resolving each attachment independently
 * therefore aimed two downloads at one path: with --force the second write
 * clobbered the first while the results array still reported both as written,
 * and without it the batch aborted halfway through.
 */
describe('attachments download — destination collisions', () => {
  it('THE CRITICAL ONE: two attachments with the same name both survive, and neither is silently lost', async () => {
    const dir = await makeTempDir();
    const a1 = sampleAttachment({ id: 'a1', name: 'image001.png', contentBytes: Buffer.from('first').toString('base64') });
    const a2 = sampleAttachment({ id: 'a2', name: 'image001.png', contentBytes: Buffer.from('second').toString('base64') });
    const { ctx, client } = makeCtx(['download', 'msg-1', '--all', '--out', dir]);
    stubPages(client, { value: [a1, a2] });

    await attachmentsCommand.handler(ctx);

    expect(await readFile(join(dir, 'image001.png'), 'utf8')).toBe('first');
    expect(await readFile(join(dir, 'image001 (2).png'), 'utf8')).toBe('second');
    expect((await readdir(dir)).sort()).toEqual(['image001 (2).png', 'image001.png']);
  });

  it('--force no longer lets one attachment overwrite a sibling in the same batch', async () => {
    const dir = await makeTempDir();
    const a1 = sampleAttachment({ id: 'a1', name: 'dup.txt', contentBytes: Buffer.from('first').toString('base64') });
    const a2 = sampleAttachment({ id: 'a2', name: 'dup.txt', contentBytes: Buffer.from('second').toString('base64') });
    const { ctx, client } = makeCtx(['download', 'msg-1', '--all', '--force', '--out', dir]);
    stubPages(client, { value: [a1, a2] });

    await attachmentsCommand.handler(ctx);

    expect(await readFile(join(dir, 'dup.txt'), 'utf8')).toBe('first');
    expect(await readFile(join(dir, 'dup (2).txt'), 'utf8')).toBe('second');
  });

  it('two names that SANITIZE onto one path are separated (a:b.txt vs a?b.txt)', async () => {
    const dir = await makeTempDir();
    const a1 = sampleAttachment({ id: 'a1', name: 'a:b.txt', contentBytes: Buffer.from('colon').toString('base64') });
    const a2 = sampleAttachment({ id: 'a2', name: 'a?b.txt', contentBytes: Buffer.from('question').toString('base64') });
    const { ctx, client } = makeCtx(['download', 'msg-1', '--all', '--out', dir]);
    stubPages(client, { value: [a1, a2] });

    await attachmentsCommand.handler(ctx);

    expect(await readFile(join(dir, 'a_b.txt'), 'utf8')).toBe('colon');
    expect(await readFile(join(dir, 'a_b (2).txt'), 'utf8')).toBe('question');
  });

  it('reports each attachment at the path it actually got, not a shared one', async () => {
    const dir = await makeTempDir();
    const a1 = sampleAttachment({ id: 'a1', name: 'dup.txt' });
    const a2 = sampleAttachment({ id: 'a2', name: 'dup.txt' });
    const { ctx, client, out } = makeCtx(['download', 'msg-1', '--all', '--out', dir, '--json']);
    stubPages(client, { value: [a1, a2] });

    await attachmentsCommand.handler(ctx);

    const results = JSON.parse(text(out)) as { id: string; path: string }[];
    expect(results.map((r) => r.path)).toEqual([join(dir, 'dup.txt'), join(dir, 'dup (2).txt')]);
    expect(new Set(results.map((r) => r.path)).size).toBe(2);
  });

  it('the dry-run plan names exactly the paths a real run writes', async () => {
    const dir = await makeTempDir();
    const attachments = [
      sampleAttachment({ id: 'a1', name: 'dup.txt' }),
      sampleAttachment({ id: 'a2', name: 'dup.txt' }),
      sampleAttachment({ id: 'a3', name: 'a:b.txt' }),
      sampleAttachment({ id: 'a4', name: 'a?b.txt' }),
    ];

    const dry = makeCtx(['download', 'msg-1', '--all', '--dry-run', '--out', dir, '--json']);
    stubPages(dry.client, { value: attachments });
    await attachmentsCommand.handler(dry.ctx);
    const planned = (JSON.parse(text(dry.out)) as { wouldWrite: { path: string }[] }).wouldWrite;

    const real = makeCtx(['download', 'msg-1', '--all', '--out', dir, '--json']);
    stubPages(real.client, { value: attachments });
    await attachmentsCommand.handler(real.ctx);
    const written = JSON.parse(text(real.out)) as { path: string }[];

    expect(planned.map((p) => p.path)).toEqual(written.map((w) => w.path));
  });

  it('disambiguation keeps the extension, so the file stays openable', async () => {
    const dir = await makeTempDir();
    const attachments = [
      sampleAttachment({ id: 'a1', name: 'report.pdf' }),
      sampleAttachment({ id: 'a2', name: 'report.pdf' }),
      sampleAttachment({ id: 'a3', name: 'report.pdf' }),
    ];
    const { ctx, client } = makeCtx(['download', 'msg-1', '--all', '--out', dir]);
    stubPages(client, { value: attachments });

    await attachmentsCommand.handler(ctx);

    expect((await readdir(dir)).sort()).toEqual(['report (2).pdf', 'report (3).pdf', 'report.pdf']);
  });
});

/**
 * Containment check (review finding F-A6). `rel.startsWith('..')` also rejected
 * a legitimate `..gitignore`, whose relative path merely begins with those two
 * characters. Escape means the whole first segment is `..`.
 */
describe('attachments download — containment check precision', () => {
  it('writes a legitimate name that merely begins with ".." instead of refusing it', async () => {
    const dir = await makeTempDir();
    const attachment = sampleAttachment({
      id: 'a1',
      name: '..gitignore',
      contentBytes: Buffer.from('dotfile').toString('base64'),
    });
    const { ctx, client } = makeCtx(['download', 'msg-1', attachment.id, '--out', dir]);
    stubPages(client, { value: [attachment] });

    await attachmentsCommand.handler(ctx);

    expect(await readFile(join(dir, '..gitignore'), 'utf8')).toBe('dotfile');
  });

  it('regression: deep traversal is still contained', async () => {
    const dir = await makeTempDir();
    const evil = sampleAttachment({
      id: 'a1',
      name: '../../etc/passwd',
      contentBytes: Buffer.from('payload').toString('base64'),
    });
    const { ctx, client } = makeCtx(['download', 'msg-1', evil.id, '--out', dir]);
    stubPages(client, { value: [evil] });

    await attachmentsCommand.handler(ctx);

    expect(await readFile(join(dir, 'passwd'), 'utf8')).toBe('payload');
    expect(await readdir(dir)).toEqual(['passwd']);
  });
});
