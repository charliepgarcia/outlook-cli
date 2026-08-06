/**
 * `outlook get` — GET /me/messages/{id} (docs/01-read-endpoints.md §1, §6.4).
 *
 * Read-only: `mutating: false`. Nothing here is irreversible, so there is no
 * `guardMutation` wrapper — that chokepoint exists for state-changing calls
 * only (see command.ts).
 *
 * The message id is positional, not a flag (`requirePositional` gives a
 * clean UsageError instead of an undefined-id crash). Graph message ids are
 * opaque strings (docs/01 §6.4 example: "AAMkADhMGAAA=") that can contain
 * characters unsafe to splice unencoded into a URL path segment, so the id
 * is percent-encoded before being joined onto `/me/messages/`.
 *
 * `body` is returned in HTML by default (docs/01 §6.2) unless a caller sends
 * `Prefer: outlook.body-content-type="text"` — this command never sends that
 * header. Instead `--html` is forwarded straight into `renderMessage`'s
 * `opts.html` (src/output.ts), which already knows how to choose between the
 * raw HTML and a plain-text conversion of whatever `contentType` came back.
 *
 * docs/01 §6.1/§7 documents `attachments` only as a *separate* collection
 * endpoint (`GET /me/messages/{id}/attachments`) — no `$select`/`$expand`
 * combination that inlines attachment metadata into a message-get response
 * is documented anywhere in docs/01. So there is deliberately no
 * `--attachments` flag here; adding one would invent an unverified query
 * shape.
 */

import { emit, type CommandContext, type CommandSpec } from '../command.js';
import { renderMessage } from '../output.js';
import type { Message } from '../types.js';

async function getHandler(ctx: CommandContext): Promise<void> {
  const id = ctx.args.requirePositional(0, '<messageId>');
  const path = `/me/messages/${encodeURIComponent(id)}`;

  const message = await ctx.client.get<Message>(path);

  emit(ctx, message, () => renderMessage(message, { html: ctx.args.bool('html') }));
}

export const getCommand: CommandSpec = {
  name: 'get',
  summary: 'Fetch and display a single message',
  usage: 'outlook get <messageId> [--html] [--json]',
  mutating: false,
  handler: getHandler,
};
