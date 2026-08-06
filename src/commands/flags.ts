/**
 * outlook mark-read / mark-unread / flag / unflag.
 *
 * All four are `PATCH /me/messages/{id}` (docs/02-write-endpoints.md §5).
 * Bodies are the exact shapes from that section:
 *   isRead:  { "isRead": true }            / { "isRead": false }
 *   flag:    { "flag": { "flagStatus": "flagged" } }
 *            { "flag": { "flagStatus": "notFlagged" } }
 * `flagStatus` is documented as exactly `notFlagged` | `flagged` | `complete`
 * (lowercase/camelCase, no capitalized variants) — see FlagStatus in types.ts.
 */

import type { CommandContext, CommandSpec } from '../command.js';
import { guardMutation, emit } from '../command.js';
import type { FlagStatus, Message } from '../types.js';

function messagePath(id: string): string {
  return `/me/messages/${encodeURIComponent(id)}`;
}

async function patchMessage(
  ctx: CommandContext,
  id: string,
  body: Record<string, unknown>,
  describe: (updated: Message) => string,
): Promise<void> {
  const path = messagePath(id);
  const spec = ctx.client.buildSpec('PATCH', path, body);

  await guardMutation(ctx, spec, async () => {
    const updated = await ctx.client.patch<Message>(path, body);
    emit(ctx, { id, message: updated }, () => describe(updated));
  });
}

async function setRead(ctx: CommandContext, isRead: boolean): Promise<void> {
  const id = ctx.args.requirePositional(0, '<id>');
  await patchMessage(ctx, id, { isRead }, () => `Marked ${id} as ${isRead ? 'read' : 'unread'}.`);
}

export const markReadCommand: CommandSpec = {
  name: 'mark-read',
  summary: 'Mark a message as read.',
  usage: 'outlook mark-read <id> [--dry-run] [--json]',
  mutating: true,
  handler: (ctx) => setRead(ctx, true),
};

export const markUnreadCommand: CommandSpec = {
  name: 'mark-unread',
  summary: 'Mark a message as unread.',
  usage: 'outlook mark-unread <id> [--dry-run] [--json]',
  mutating: true,
  handler: (ctx) => setRead(ctx, false),
};

async function setFlag(ctx: CommandContext, flagStatus: FlagStatus): Promise<void> {
  const id = ctx.args.requirePositional(0, '<id>');
  await patchMessage(
    ctx,
    id,
    { flag: { flagStatus } },
    () => `Set flag on ${id} to "${flagStatus}".`,
  );
}

export const flagCommand: CommandSpec = {
  name: 'flag',
  summary: 'Flag a message for follow-up.',
  usage: 'outlook flag <id> [--dry-run] [--json]',
  mutating: true,
  handler: (ctx) => setFlag(ctx, 'flagged'),
};

export const unflagCommand: CommandSpec = {
  name: 'unflag',
  summary: 'Clear a message\'s follow-up flag.',
  usage: 'outlook unflag <id> [--dry-run] [--json]',
  mutating: true,
  handler: (ctx) => setFlag(ctx, 'notFlagged'),
};
