/**
 * `outlook attachments` — list and download message attachments
 * (docs/01-read-endpoints.md §7).
 *
 * Read-only: `mutating: false`. Both subcommands only ever GET from Graph —
 * `list` renders metadata, `download` writes bytes to LOCAL disk. Writing a
 * file to the caller's own filesystem is not a mutation of remote mailbox
 * state, so there is no `guardMutation`/`--dry-run` wrapper here (that
 * chokepoint in command.ts exists for calls that change what Graph stores).
 *
 * Only one Graph endpoint is used, for both subcommands:
 *   GET /me/messages/{id}/attachments                          (docs/01 §7.1)
 * docs/01 §7.3's own example response embeds `contentBytes` directly in the
 * list payload (no `$expand` needed), and §7.1 additionally documents that
 * `$filter` on `/attachments` is silently ignored — "filter client-side after
 * listing" is the docs' own guidance (§7.1). So `download <id> [attachmentId]`
 * does not invent a separate single-attachment GET path; it fetches the same
 * list `collectItems` fetches for `list` and selects client-side. Pagination
 * is always fully collected (`{ all: true }`) since docs/01 documents no
 * `$top` ceiling for this collection (§9, open question 3) — a partial page
 * would silently under-list or under-download attachments.
 *
 * Subcommand dispatch on `positionals[0]`, but unlike folders.ts's `list`
 * (which takes no id of its own), `attachments list` DOES consume a
 * positional (`<messageId>`) — so an unrecognized first token can't be
 * treated as an invalid-subcommand error the way folders.ts does; it must
 * be treated as the id itself under the default `list` behavior. Only the
 * two known keywords shift the id one slot to the right:
 *   `download ...`      -> download subcommand, id at positionals[1]
 *   `list <messageId>`  -> list subcommand,     id at positionals[1]
 *   anything else / absent -> list subcommand,  id at positionals[0]
 * (Graph message ids are opaque generated strings and never collide with
 * the literal keywords "list"/"download" in practice.)
 */

import { access, mkdir, writeFile } from 'node:fs/promises';
import { basename, isAbsolute, relative, resolve } from 'node:path';
import { emit, type CommandContext, type CommandSpec } from '../command.js';
import { UsageError } from '../errors.js';
import { renderAttachments } from '../output.js';
import { collectItems } from '../paginate.js';
import type { Attachment } from '../types.js';

const FILE_TYPE = 'microsoft.graph.fileAttachment';

/** Strips an optional leading '#' (docs/01 §7.3: S9's own example omits it). */
function normalizedType(a: Attachment): string {
  return (a['@odata.type'] ?? '').replace(/^#/, '');
}

function isFileAttachment(a: Attachment): boolean {
  return normalizedType(a) === FILE_TYPE;
}

/** Human label for error messages. output.ts's equivalent `attachmentKind`
 * helper is not exported, so this is a small, deliberately local duplicate. */
function kindLabel(a: Attachment): string {
  const t = normalizedType(a);
  return t.startsWith('microsoft.graph.') ? t.slice('microsoft.graph.'.length) : t || 'unknown';
}

async function fetchAttachments(ctx: CommandContext, messageId: string): Promise<Attachment[]> {
  const path = `/me/messages/${encodeURIComponent(messageId)}/attachments`;
  return collectItems<Attachment>(ctx.client, path, { all: true });
}

async function listHandler(ctx: CommandContext, idIndex: number): Promise<void> {
  const messageId = ctx.args.requirePositional(idIndex, '<messageId>');
  const attachments = await fetchAttachments(ctx, messageId);
  emit(ctx, attachments, () => renderAttachments(attachments));
}

/**
 * Basenames the server-supplied name (strips any `../` or absolute prefix)
 * and then double-checks the resolved path is still inside `outDir` — the
 * two required layers of the path-traversal guard (docs' `name` is
 * attacker-controlled server data, not to be trusted for local disk writes).
 */
function safeDestination(outDir: string, attachment: Attachment): string {
  const raw = attachment.name && attachment.name.trim() ? attachment.name : `attachment-${attachment.id}`;
  const base = basename(raw);
  const name = base === '' || base === '.' || base === '..' ? `attachment-${attachment.id}` : base;
  const dest = resolve(outDir, name);
  const rel = relative(outDir, dest);
  if (rel.startsWith('..') || isAbsolute(rel)) {
    throw new UsageError(`Refusing to write outside --out directory: "${attachment.name}"`);
  }
  return dest;
}

async function exists(p: string): Promise<boolean> {
  try {
    await access(p);
    return true;
  } catch {
    return false;
  }
}

interface DownloadResult {
  id: string;
  name: string;
  path: string;
  size: number;
}

async function downloadOne(dest: string, attachment: Attachment, force: boolean): Promise<DownloadResult> {
  if (!isFileAttachment(attachment)) {
    throw new UsageError(
      `Cannot download "${attachment.name}" — ${kindLabel(attachment)} has no downloadable bytes (docs/01 §7.4)`,
    );
  }
  if (!force && (await exists(dest))) {
    throw new UsageError(`Refusing to overwrite existing file: ${dest} (pass --force)`);
  }
  const buf = Buffer.from(attachment.contentBytes ?? '', 'base64');
  await writeFile(dest, buf);
  return { id: attachment.id, name: attachment.name, path: dest, size: buf.length };
}

async function downloadHandler(ctx: CommandContext): Promise<void> {
  const messageId = ctx.args.requirePositional(1, '<messageId>');
  const attachmentId = ctx.args.positionals[2];
  const all = ctx.args.bool('all');
  const force = ctx.args.bool('force');
  const outDir = resolve(ctx.args.str('out', process.cwd()) ?? process.cwd());

  if (!all && !attachmentId) {
    throw new UsageError('outlook attachments download requires [attachmentId] or --all');
  }

  const attachments = await fetchAttachments(ctx, messageId);
  const targets = all ? attachments : attachments.filter((a) => a.id === attachmentId);
  if (!all && targets.length === 0) {
    throw new UsageError(`No attachment with id "${attachmentId}" on message ${messageId}`);
  }

  await mkdir(outDir, { recursive: true });

  const results: DownloadResult[] = [];
  for (const attachment of targets) {
    const dest = safeDestination(outDir, attachment);
    results.push(await downloadOne(dest, attachment, force));
  }

  emit(ctx, results, () =>
    results.length === 0
      ? '(nothing to download)'
      : results.map((r) => `${r.name} -> ${r.path} (${r.size} bytes)`).join('\n'),
  );
}

async function attachmentsHandler(ctx: CommandContext): Promise<void> {
  const sub = ctx.args.positionals[0];
  if (sub === 'download') return downloadHandler(ctx);
  if (sub === 'list') return listHandler(ctx, 1);
  return listHandler(ctx, 0); // no keyword: sub, if present, IS the messageId
}

export const attachmentsCommand: CommandSpec = {
  name: 'attachments',
  summary: 'List or download the attachments on a message',
  usage: 'outlook attachments [list|download] <messageId> [attachmentId] [--all] [--out dir] [--force] [--json]',
  mutating: false,
  handler: attachmentsHandler,
};
