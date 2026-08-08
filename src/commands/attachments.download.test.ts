/**
 * `outlook attachments download` — the two security-fix concerns:
 * --dry-run must touch nothing on disk, and a server-supplied attachment name
 * must be sanitized before it becomes a path.
 *
 * Split out of attachments.test.ts purely to stay under the 350-LOC ceiling.
 */
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

import { mkdtemp, readdir, readFile, rm, writeFile } from 'node:fs/promises';
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
/**
 * Security fix regression coverage (docs/security-review-2026-08-07.md #6b):
 * `--dry-run` used to decode base64 and call `writeFile` for real. The fix
 * moves the dry-run check BEFORE `mkdir`/`writeFile`, so a dry run must leave
 * the filesystem completely untouched — not even the --out directory gets
 * created if it did not already exist.
 */
describe('attachments download --dry-run', () => {
  it('THE CRITICAL ONE: --all --dry-run writes no files, and never even creates --out if absent', async () => {
    const dir = await makeTempDir();
    // Deliberately never created — the dry-run branch returns before mkdir.
    const outDir = join(dir, 'nested', 'out');
    const a1 = sampleAttachment({ id: 'a1', name: 'one.txt' });
    const a2 = sampleAttachment({ id: 'a2', name: 'two.txt' });
    const { ctx, client } = makeCtx(['download', 'msg-1', '--all', '--dry-run', '--out', outDir]);
    stubPages(client, { value: [a1, a2] });

    await attachmentsCommand.handler(ctx);

    // The out dir may not exist at all — assert accordingly rather than
    // assuming readdir resolves.
    await expect(readdir(outDir)).rejects.toMatchObject({ code: 'ENOENT' });
    // And nothing leaked into the parent temp dir either.
    expect(await readdir(dir)).toEqual([]);
  });

  it('--all --dry-run writes no files when --out already exists', async () => {
    const dir = await makeTempDir();
    const a1 = sampleAttachment({ id: 'a1', name: 'one.txt' });
    const a2 = sampleAttachment({ id: 'a2', name: 'two.txt' });
    const { ctx, client } = makeCtx(['download', 'msg-1', '--all', '--dry-run', '--out', dir]);
    stubPages(client, { value: [a1, a2] });

    await attachmentsCommand.handler(ctx);

    expect(await readdir(dir)).toEqual([]);
  });

  it('reports "DRY RUN" and the destination filename in human-readable output', async () => {
    const dir = await makeTempDir();
    const attachment = sampleAttachment({ id: 'a1', name: 'notes.txt' });
    const { ctx, client, out } = makeCtx(['download', 'msg-1', '--all', '--dry-run', '--out', dir]);
    stubPages(client, { value: [attachment] });

    await attachmentsCommand.handler(ctx);

    expect(text(out)).toContain('DRY RUN');
    expect(text(out)).toContain('notes.txt');
    expect(await readdir(dir)).toEqual([]);
  });

  it('--json reports dryRun: true and a wouldWrite array of planned destinations', async () => {
    const dir = await makeTempDir();
    const attachment = sampleAttachment({ id: 'a1', name: 'notes.txt' });
    const { ctx, client, out } = makeCtx(['download', 'msg-1', '--all', '--dry-run', '--out', dir, '--json']);
    stubPages(client, { value: [attachment] });

    await attachmentsCommand.handler(ctx);

    const emitted = JSON.parse(text(out)) as { dryRun: boolean; wouldWrite: unknown[] };
    expect(emitted.dryRun).toBe(true);
    expect(emitted.wouldWrite).toEqual([{ id: 'a1', name: 'notes.txt', path: join(dir, 'notes.txt') }]);
    expect(await readdir(dir)).toEqual([]);
  });

  it('regression: a real (non-dry-run) --all download still writes the files, unaffected by the dry-run branch', async () => {
    const dir = await makeTempDir();
    const a1 = sampleAttachment({ id: 'a1', name: 'one.txt', contentBytes: Buffer.from('one').toString('base64') });
    const a2 = sampleAttachment({ id: 'a2', name: 'two.txt', contentBytes: Buffer.from('two').toString('base64') });
    const { ctx, client } = makeCtx(['download', 'msg-1', '--all', '--out', dir]);
    stubPages(client, { value: [a1, a2] });

    await attachmentsCommand.handler(ctx);

    expect(await readFile(join(dir, 'one.txt'), 'utf8')).toBe('one');
    expect(await readFile(join(dir, 'two.txt'), 'utf8')).toBe('two');
  });
});

/**
 * safeFilename.ts sanitization, exercised through the real download path
 * (docs/security-review-2026-08-07.md #7). See safeFilename.ts for the exact
 * transformation: illegal Win32 chars (incl. the NTFS ADS colon) become '_',
 * and reserved DOS device names are prefixed with '_' rather than replaced.
 */
describe('attachments download — filename sanitization', () => {
  it('prefixes a Windows-reserved device name (CON) with an underscore', async () => {
    const dir = await makeTempDir();
    const attachment = sampleAttachment({
      id: 'a1',
      name: 'CON',
      contentBytes: Buffer.from('reserved').toString('base64'),
    });
    const { ctx, client } = makeCtx(['download', 'msg-1', attachment.id, '--out', dir]);
    stubPages(client, { value: [attachment] });

    await attachmentsCommand.handler(ctx);

    expect(await readFile(join(dir, '_CON'), 'utf8')).toBe('reserved');
    expect(await readdir(dir)).toEqual(['_CON']);
  });

  it('replaces a colon (NTFS alternate-data-stream hazard) with an underscore', async () => {
    const dir = await makeTempDir();
    const attachment = sampleAttachment({
      id: 'a1',
      name: 'a:b.txt',
      contentBytes: Buffer.from('payload').toString('base64'),
    });
    const { ctx, client } = makeCtx(['download', 'msg-1', attachment.id, '--out', dir]);
    stubPages(client, { value: [attachment] });

    await attachmentsCommand.handler(ctx);

    expect(await readFile(join(dir, 'a_b.txt'), 'utf8')).toBe('payload');
  });

  it('regression: path-traversal protection still holds under --all, in both dry-run and real downloads', async () => {
    const dir = await makeTempDir();
    const evil = sampleAttachment({
      id: 'a1',
      name: '../evil.txt',
      contentBytes: Buffer.from('payload').toString('base64'),
    });

    const dry = makeCtx(['download', 'msg-1', '--all', '--dry-run', '--out', dir]);
    stubPages(dry.client, { value: [evil] });
    await attachmentsCommand.handler(dry.ctx);
    expect(await readdir(dir)).toEqual([]); // dry-run: still nothing written

    const real = makeCtx(['download', 'msg-1', '--all', '--out', dir]);
    stubPages(real.client, { value: [evil] });
    await attachmentsCommand.handler(real.ctx);

    expect(await readFile(join(dir, 'evil.txt'), 'utf8')).toBe('payload');
    await expect(readFile(join(dir, '..', 'evil.txt'), 'utf8')).rejects.toThrow();
  });
});
