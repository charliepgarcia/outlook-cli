/**
 * Tests for `outlook folders`. Mirrors get.test.ts (real GraphClient from
 * makeCtx, network boundary stubbed via vi.spyOn) and send.test.ts's
 * --dry-run assertion (fetchImpl must never be called).
 */

import { describe, expect, it, vi } from 'vitest';
import type { GraphClient } from '../client.js';
import { UsageError } from '../errors.js';
import { GRAPH_BASE, type GraphCollection, type MailFolder, type RequestSpec } from '../types.js';
import { foldersCommand } from './folders.js';
import { makeCtx, text } from './testUtils.js';

/**
 * The network boundary for a PAGINATED read is `client.request`, NOT
 * `client.get`: collectItems/pageItems build a RequestSpec through
 * client.buildSpec and hand it straight to client.request (paginate.ts).
 * Stubbing `get` leaves `request` live, the real fetch fires, and testUtils'
 * deliberately-throwing fetchImpl trips with "NETWORK CALLED". Stub here.
 */
function stubPages(
  client: GraphClient,
  ...pages: GraphCollection<MailFolder>[]
) {
  const spy = vi.spyOn(client, 'request');
  for (const page of pages) spy.mockResolvedValueOnce(page as never);
  return spy;
}

/**
 * Relative Graph path of a captured RequestSpec. `spec.url` is absolute
 * (GRAPH_BASE + path + query), so strip the base and drop the query string.
 */
function pathOf(spy: { mock: { calls: unknown[][] } }, call = 0): string {
  const spec = spy.mock.calls[call]?.[0] as RequestSpec | undefined;
  if (spec === undefined) return '';
  const url = spec.url.startsWith(GRAPH_BASE)
    ? spec.url.slice(GRAPH_BASE.length)
    : spec.url;
  return url.split('?')[0] ?? '';
}

function sampleFolder(overrides: Partial<MailFolder> = {}): MailFolder {
  return {
    id: 'AAMkADhMGAAA=',
    displayName: 'Inbox',
    parentFolderId: 'root',
    childFolderCount: 2,
    unreadItemCount: 3,
    totalItemCount: 10,
    ...overrides,
  };
}

describe('folders list', () => {
  it('GETs /me/mailFolders and renders human-readable output', async () => {
    const { ctx, client, out } = makeCtx(['list']);
    const folders = [sampleFolder(), sampleFolder({ id: 'f2', displayName: 'Archive' })];
    const req = stubPages(client, { value: folders });

    await foldersCommand.handler(ctx);

    expect(req).toHaveBeenCalledTimes(1);
    expect(pathOf(req)).toBe('/me/mailFolders');
    expect(text(out)).toContain('Inbox');
    expect(text(out)).toContain('Archive');
  });

  it('defaults to "list" when no subcommand is given', async () => {
    const { ctx, client } = makeCtx([]);
    const req = stubPages(client, { value: [sampleFolder()] });

    await foldersCommand.handler(ctx);

    expect(req).toHaveBeenCalledTimes(1);
    expect(pathOf(req)).toBe('/me/mailFolders');
  });

  it('--json emits the raw folder array, unmodified', async () => {
    const { ctx, client, out } = makeCtx(['list', '--json']);
    const folders = [sampleFolder()];
    stubPages(client, { value: folders });

    await foldersCommand.handler(ctx);

    expect(JSON.parse(text(out))).toEqual(folders);
  });

  it('without --json, output is human text, not a JSON blob', async () => {
    const { ctx, client, out } = makeCtx(['list']);
    stubPages(client, { value: [sampleFolder()] });

    await foldersCommand.handler(ctx);

    expect(() => JSON.parse(text(out))).toThrow();
  });

  it('--all paginates through every page', async () => {
    const { ctx, client } = makeCtx(['list', '--all']);
    const req = stubPages(
      client,
      {
        value: [sampleFolder({ id: 'p1' })],
        '@odata.nextLink': `${GRAPH_BASE}/me/mailFolders?$skip=1`,
      },
      { value: [sampleFolder({ id: 'p2' })] },
    );

    await foldersCommand.handler(ctx);

    expect(req).toHaveBeenCalledTimes(2);
  });

  it('--parent <name> resolves to the documented childFolders path, lowercasing well-known names', async () => {
    const { ctx, client } = makeCtx(['list', '--parent', 'Inbox']);
    const req = stubPages(client, { value: [sampleFolder()] });

    await foldersCommand.handler(ctx);

    expect(pathOf(req)).toBe('/me/mailFolders/inbox/childFolders');
  });

  it('--parent <id> (not a well-known name) is passed through unchanged and percent-encoded', async () => {
    const rawId = 'AAMk/odd=folder';
    const { ctx, client } = makeCtx(['list', '--parent', rawId]);
    const req = stubPages(client, { value: [] });

    await foldersCommand.handler(ctx);

    expect(pathOf(req)).toBe(`/me/mailFolders/${encodeURIComponent(rawId)}/childFolders`);
  });

  it('never touches the network directly — only through the stubbed client method', async () => {
    const { ctx, client, fetchImpl } = makeCtx(['list']);
    stubPages(client, { value: [sampleFolder()] });

    await foldersCommand.handler(ctx);

    expect(fetchImpl).toHaveBeenCalledTimes(0);
  });
});

describe('folders create', () => {
  it('builds the correct POST path and body', async () => {
    const { ctx, client } = makeCtx(['create', '--name', 'Clutter']);
    const post = vi.spyOn(client, 'post').mockResolvedValue(
      sampleFolder({ displayName: 'Clutter' }),
    );

    await foldersCommand.handler(ctx);

    expect(post).toHaveBeenCalledTimes(1);
    expect(post.mock.calls[0]?.[0]).toBe('/me/mailFolders');
    expect(post.mock.calls[0]?.[1]).toEqual({ displayName: 'Clutter' });
  });

  it('emits the created folder on success', async () => {
    const { ctx, client, out } = makeCtx(['create', '--name', 'Clutter']);
    vi.spyOn(client, 'post').mockResolvedValue(sampleFolder({ displayName: 'Clutter' }));

    await foldersCommand.handler(ctx);

    expect(text(out)).toContain('Clutter');
  });

  it('--dry-run performs NO network call and prints the dry-run report', async () => {
    const { ctx, client, fetchImpl, out } = makeCtx(['create', '--name', 'Clutter', '--dry-run']);
    const post = vi.spyOn(client, 'post');

    await foldersCommand.handler(ctx);

    expect(post).toHaveBeenCalledTimes(0);
    expect(fetchImpl).toHaveBeenCalledTimes(0);
    expect(text(out)).toContain('DRY RUN');
    expect(text(out)).toContain('/me/mailFolders');
  });

  it('missing --name throws a clean UsageError', async () => {
    const { ctx, client } = makeCtx(['create']);
    const post = vi.spyOn(client, 'post');

    await expect(foldersCommand.handler(ctx)).rejects.toBeInstanceOf(UsageError);
    expect(post).toHaveBeenCalledTimes(0);
  });

  it('rejects --parent as an invented, unsupported flag on create', async () => {
    const { ctx, client } = makeCtx(['create', '--name', 'Clutter', '--parent', 'inbox']);
    const post = vi.spyOn(client, 'post');

    await expect(foldersCommand.handler(ctx)).rejects.toBeInstanceOf(UsageError);
    expect(post).toHaveBeenCalledTimes(0);
  });
});

describe('folders — unknown subcommand', () => {
  it('throws a clean UsageError for an unrecognized subcommand', async () => {
    const { ctx } = makeCtx(['bogus']);

    await expect(foldersCommand.handler(ctx)).rejects.toBeInstanceOf(UsageError);
  });
});
