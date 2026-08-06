/**
 * `outlook search <query>` — KQL/OData `$search` over messages
 * (docs/01-read-endpoints.md §4 "`$search` and KQL (message search)").
 *
 * Endpoint: `GET /me/messages?$search=...` by default, or
 * `GET /me/mailFolders/{id}/messages?$search=...` when `--folder` scopes the
 * search to one folder — docs/01 §3's intro line lists
 * `/me/mailFolders/{id}/messages` as one of the three endpoints that
 * "supports the OData Query Parameters" (of which `$search` is one),
 * alongside `/me/messages` and `/me/messages/{id}`.
 *
 * Quoting (docs/01 §4a, verbatim): "the search string itself must be wrapped
 * in double quotes inside the query value (`$search="pizza"`), per S12's own
 * examples." §4's general syntax rules add: "embedded double quotes or
 * backslashes must be backslash-escaped; other special characters must be
 * URL-encoded." This command wraps the ENTIRE positional query in one outer
 * pair of double quotes and backslash-escapes any embedded `"`/`\` before
 * handing it to `collectItems`/`client.buildSpec` as the `$search` value;
 * percent-encoding of the resulting string (quotes included) is left to
 * buildSpec, exactly as for every other query parameter.
 *
 * Deliberately absent: `$orderby` and `$filter`. docs/01 §3.7/§9 document
 * `$search` on messages as carrying its own **built-in sort** — "search
 * results for messages ... are sorted by the date and time the message was
 * sent" (S12) — and flag `$search`+`$orderby` on `/me/messages` as an
 * unverified interaction: "the CLI should defensively avoid the combination
 * ... and fall back to client-side sort if needed" (§9 item 1). §3.7 flags
 * `$search`+`$filter` the same way for messages (§9 item 2: "supported or
 * rejected" unconfirmed). Unlike `list.ts` (which always sets
 * `$orderby=receivedDateTime desc`), this command never sends `$orderby` or
 * `$filter` at all — no flag here maps to either — so the documented risk
 * has no code path that could trigger it.
 *
 * `--limit` (default 25) is validated against the range docs/01 documents
 * for `$search` on messages specifically: "up to 1,000 results" (§4a) — the
 * same 1-1000 ceiling §3.5 gives generally for `$top`. A `--limit` outside
 * that range can never be honored by the API, so it is rejected locally as
 * a `UsageError` (naming the ceiling) instead of being sent as a request
 * guaranteed to misbehave.
 *
 * `--folder` scoping mirrors `folders.ts`'s tested behavior: a recognized
 * well-known name (`isWellKnownFolder`) is lowercased (well-known names are
 * documented case-insensitive, docs/01 §2); anything else is treated as a
 * raw folder id and percent-encoded.
 *
 * Read-only: `mutating: false`, no `guardMutation` (see get.ts — that
 * chokepoint is reserved for state-changing commands).
 */

import { emit, type CommandContext, type CommandSpec } from '../command.js';
import { UsageError } from '../errors.js';
import { renderMessageList } from '../output.js';
import { collectItems } from '../paginate.js';
import { isWellKnownFolder, type Message } from '../types.js';

const MAX_RESULTS = 1000;
const SELECT = 'subject,from,receivedDateTime,isRead,hasAttachments,bodyPreview';

function folderSegment(folder: string): string {
  return isWellKnownFolder(folder) ? folder.toLowerCase() : encodeURIComponent(folder);
}

function searchPath(folder: string | undefined): string {
  if (folder === undefined) return '/me/messages';
  return `/me/mailFolders/${folderSegment(folder)}/messages`;
}

/**
 * Wraps the whole query in one outer pair of double quotes and
 * backslash-escapes embedded `"`/`\`, per docs/01 §4's quoting rules.
 * Backslashes are escaped first so the quote-escaping step doesn't
 * re-escape the backslashes it just inserted.
 */
function quoteSearch(query: string): string {
  const escaped = query.replace(/\\/g, '\\\\').replace(/"/g, '\\"');
  return `"${escaped}"`;
}

function readLimit(ctx: CommandContext): number {
  if (!ctx.args.has('limit')) return 25;
  const limit = ctx.args.num('limit');
  if (limit === undefined || !Number.isInteger(limit) || limit < 1 || limit > MAX_RESULTS) {
    throw new UsageError(
      `--limit must be an integer between 1 and ${MAX_RESULTS} — docs/01 §4a ` +
        `documents $search on messages as returning "up to 1,000 results"`,
    );
  }
  return limit;
}

async function searchHandler(ctx: CommandContext): Promise<void> {
  const query = ctx.args.requirePositional(0, '<query>');
  const limit = readLimit(ctx);
  const all = ctx.args.bool('all');
  const folder = ctx.args.str('folder');

  const messages = await collectItems<Message>(ctx.client, searchPath(folder), {
    query: { $search: quoteSearch(query), $select: SELECT, $top: limit },
    limit,
    all,
  });

  emit(ctx, messages, () => renderMessageList(messages));
}

export const searchCommand: CommandSpec = {
  name: 'search',
  summary: 'Search messages with KQL/$search',
  usage: 'outlook search <query> [--limit <n>] [--all] [--folder <name|id>] [--json]',
  mutating: false,
  handler: searchHandler,
};
