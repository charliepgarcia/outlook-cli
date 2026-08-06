/**
 * `outlook folders` — list and create mail folders.
 *
 * Subcommand dispatch on positional 0 (`list` is the default when absent):
 *   outlook folders [list] [--parent <name|id>] [--all] [--json]
 *   outlook folders create --name <text> [--dry-run] [--json]
 *
 * LIST: `GET /me/mailFolders` (docs/01-read-endpoints.md §1) returns only
 * **root-level** child folders — it does not recurse (§2, line 102: "To
 * enumerate all folders, each child folder's `childFolders` must be
 * traversed separately"). That `childFolders` collection is a real,
 * documented nested resource (confirmed by the attachments path pattern in
 * §7: `GET /me/mailFolders/{id}/childFolders/{id}/.../attachments/{id}`), so
 * `--parent <name|id>` here maps to `GET /me/mailFolders/{parent}/childFolders`
 * rather than an invented query parameter. Well-known folder names are
 * lowercase literals (docs/01 §2); an opaque folder id is passed through
 * unchanged (mirrors `move.ts`'s `resolveDestinationId`).
 *
 * CREATE: `POST /me/mailFolders` (docs/02-write-endpoints.md §7). Verbatim
 * from that doc: "displayName and isHidden are the only writable property
 * for a mailFolder object" — `parentFolderId` is explicitly **not** accepted
 * on create ("folders are always created at root"). So `create`
 * deliberately has no `--parent` support: passing one is a clean UsageError
 * rather than a silently-ignored, misleading flag. Per docs/02 §8 this write
 * is **UNSAFE to retry** ("no documented uniqueness constraint on
 * displayName; retry creates a duplicate folder"), so — like every mutation
 * in this CLI — it is wrapped in `guardMutation`.
 */

import { emit, guardMutation, type CommandContext, type CommandSpec } from '../command.js';
import { collectItems } from '../paginate.js';
import { renderFolders } from '../output.js';
import { isWellKnownFolder, type MailFolder } from '../types.js';
import { UsageError } from '../errors.js';

/** Well-known names are documented as lowercase literals; a raw folder id is opaque and case-sensitive. */
function resolveFolderId(folder: string): string {
  return isWellKnownFolder(folder) ? folder.toLowerCase() : folder;
}

function foldersListPath(parent: string | undefined): string {
  if (parent === undefined) return '/me/mailFolders';
  return `/me/mailFolders/${encodeURIComponent(resolveFolderId(parent))}/childFolders`;
}

async function listHandler(ctx: CommandContext): Promise<void> {
  const parent = ctx.args.str('parent');
  const path = foldersListPath(parent);
  const opts = ctx.args.bool('all') ? { all: true } : {};

  const folders = await collectItems<MailFolder>(ctx.client, path, opts);

  emit(ctx, folders, () => renderFolders(folders));
}

async function createHandler(ctx: CommandContext): Promise<void> {
  if (ctx.args.has('parent')) {
    throw new UsageError(
      'outlook folders create does not accept --parent: Microsoft Graph only ' +
        'creates mail folders at the mailbox root (docs/02-write-endpoints.md §7 — ' +
        '"folders are always created at root"; parentFolderId is not an accepted ' +
        'request parameter on POST /me/mailFolders).',
    );
  }

  const name = ctx.args.require('name');
  const body = { displayName: name };
  const path = '/me/mailFolders';
  const spec = ctx.client.buildSpec('POST', path, body);

  const created = await guardMutation(ctx, spec, async () => ctx.client.post<MailFolder>(path, body));

  if (created === undefined) return; // --dry-run already reported

  emit(ctx, created, () => `Created folder "${created.displayName}" (id: ${created.id})`);
}

async function foldersHandler(ctx: CommandContext): Promise<void> {
  const sub = ctx.args.positionals[0] ?? 'list';

  switch (sub) {
    case 'list':
      return listHandler(ctx);
    case 'create':
      return createHandler(ctx);
    default:
      throw new UsageError(`Unknown folders subcommand "${sub}". Expected "list" or "create".`);
  }
}

export const foldersCommand: CommandSpec = {
  name: 'folders',
  summary: 'List or create mail folders',
  usage:
    'outlook folders [list] [--parent <name|id>] [--all] [--json]\n' +
    '       outlook folders create --name <text> [--dry-run] [--json]',
  mutating: true,
  handler: foldersHandler,
};
