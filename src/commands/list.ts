/**
 * `outlook list` — list messages in a folder, newest first.
 *
 * Endpoint: `GET /me/mailFolders/{id}/messages` (docs/01-read-endpoints.md §2
 * table row "List messages in a specific folder"). `--folder` accepts either
 * a well-known folder name (validated via `isWellKnownFolder`) or an opaque
 * folder id — docs/01 §2 confirms well-known names are "usable in place of a
 * folder `{id}`, e.g. `GET /me/mailFolders/inbox/messages`. Work regardless
 * of the mailbox's display-language locale." An unrecognized `--folder`
 * value is treated as a raw folder id and percent-encoded onto the same
 * path shape.
 *
 * Read-only: `mutating: false`, no `guardMutation` (see get.ts — that
 * chokepoint is reserved for state-changing commands).
 *
 * Sort: `$orderby=receivedDateTime desc` alone — newest first (docs/01 §3.3).
 * `--unread` adds `$filter=isRead eq false` as its own independent query
 * param rather than folding it together with `$orderby` into one compound
 * expression: docs/01 §3.4 documents an exact, quotable constraint
 * (`InefficientFilter`: "The restriction or sort order is too complex for
 * this operation") that fires when `$orderby` properties don't also appear
 * in `$filter`, in the same order, before any non-`$orderby` properties.
 * Keeping `isRead` out of `$orderby` and `receivedDateTime` out of `$filter`
 * sidesteps that constraint entirely, and docs/01 §3.2 confirms
 * `$filter=isRead eq false` alone is a working example on its own.
 *
 * `$select` is kept small for list output (docs/01 §3.1 recommends
 * `$select` for latency) — just enough fields for `renderMessageList`.
 *
 * Paging: default is a single page sized by `--limit` (default 25; docs/01
 * §3.5 documents the messages `$top` range as 1-1000, so `--limit` is
 * validated against that same range). `--all` follows every
 * `@odata.nextLink` via `collectItems`'s own pagination loop instead.
 */

import { emit, type CommandContext, type CommandSpec } from '../command.js';
import { UsageError } from '../errors.js';
import { renderMessageList } from '../output.js';
import { collectItems } from '../paginate.js';
import { isWellKnownFolder, type Message } from '../types.js';

const SELECT = 'subject,from,receivedDateTime,isRead,hasAttachments,bodyPreview';
const ORDERBY = 'receivedDateTime desc';
const MAX_TOP = 1000;

function folderPath(folder: string): string {
  const segment = isWellKnownFolder(folder) ? folder.toLowerCase() : encodeURIComponent(folder);
  return `/me/mailFolders/${segment}/messages`;
}

function readLimit(ctx: CommandContext): number {
  if (!ctx.args.has('limit')) return 25;
  const limit = ctx.args.num('limit');
  if (limit === undefined || !Number.isInteger(limit) || limit < 1 || limit > MAX_TOP) {
    throw new UsageError(`--limit must be an integer between 1 and ${MAX_TOP}`);
  }
  return limit;
}

async function listHandler(ctx: CommandContext): Promise<void> {
  const folder = ctx.args.str('folder', 'inbox') ?? 'inbox';
  const limit = readLimit(ctx);
  const all = ctx.args.bool('all');
  const unread = ctx.args.bool('unread');

  const query: Record<string, unknown> = {
    $select: SELECT,
    $orderby: ORDERBY,
    $top: limit,
  };
  if (unread) query.$filter = 'isRead eq false';

  const messages = await collectItems<Message>(ctx.client, folderPath(folder), {
    query,
    limit,
    all,
  });

  emit(ctx, messages, () => renderMessageList(messages));
}

export const listCommand: CommandSpec = {
  name: 'list',
  summary: 'List messages in a folder, newest first',
  usage: 'outlook list [--folder <name|id>] [--limit <n>] [--all] [--unread] [--json]',
  mutating: false,
  handler: listHandler,
};
