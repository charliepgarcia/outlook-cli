/**
 * `outlook reply` / `reply-all` / `forward`.
 *
 * These use the one-shot Graph actions (`/reply`, `/replyAll`, `/forward`)
 * which compose AND send in a single call — irreversible, so every one is
 * wrapped in `guardMutation` for --dry-run, and none is ever retried.
 *
 * docs/02 §2 notes an inconsistency: the replyAll parameter table lists only
 * `comment`, while surrounding prose implies a `message` object may also be
 * accepted. We send only the documented `comment` field.
 *
 * Signatures: `comment` is plain text, so only a plain-text signature can be
 * appended here (an HTML-only configuration is a loud error, see
 * signature.ts). Exchange places the comment above the quoted original.
 */

import { emit, guardMutation, type CommandContext, type CommandSpec } from '../command.js';
import { UsageError } from '../errors.js';
import { appendTextSignature, resolveSignature } from '../signature.js';
import { describeRecipients, recipientsFromArgs } from './compose.js';

type ReplyKind = 'reply' | 'replyAll' | 'forward';

const ACTION_PATH: Record<ReplyKind, string> = {
  reply: 'reply',
  replyAll: 'replyAll',
  forward: 'forward',
};

async function runReplyAction(
  ctx: CommandContext,
  kind: ReplyKind,
  label: string,
): Promise<void> {
  const messageId = ctx.args.requirePositional(0, '<messageId>');
  const body = ctx.args.str('body') ?? '';
  const signature = await resolveSignature(ctx.args, ctx.env);
  const comment =
    signature === undefined ? body : appendTextSignature(body, signature);

  const payload: Record<string, unknown> = { comment };

  if (kind === 'forward') {
    const to = recipientsFromArgs(ctx.args, 'to');
    if (to.length === 0) {
      throw new UsageError(
        'forward requires at least one recipient. Use --to <address>.',
      );
    }
    payload['toRecipients'] = to;
  }

  const path = `/me/messages/${encodeURIComponent(messageId)}/${ACTION_PATH[kind]}`;
  const spec = ctx.client.buildSpec('POST', path, payload);

  const done = await guardMutation(ctx, spec, async () => {
    await ctx.client.post<void>(path, payload);
    return true;
  });

  if (done === undefined) return;

  const recipients = payload['toRecipients'];
  emit(
    ctx,
    {
      action: kind,
      messageId,
      sent: true,
      signatureApplied: signature !== undefined,
    },
    () =>
      `${label} sent for message ${messageId}` +
      (Array.isArray(recipients)
        ? ` to ${describeRecipients(recipients as never)}`
        : ''),
  );
}

export const replyCommand: CommandSpec = {
  name: 'reply',
  summary: 'Reply to the sender of a message',
  usage:
    'outlook reply <messageId> [--body <text>] ' +
    '[--signature <path> | --no-signature] [--dry-run] [--json]',
  mutating: true,
  handler: (ctx) => runReplyAction(ctx, 'reply', 'Reply'),
};

export const replyAllCommand: CommandSpec = {
  name: 'reply-all',
  summary: 'Reply to all recipients of a message',
  usage:
    'outlook reply-all <messageId> [--body <text>] ' +
    '[--signature <path> | --no-signature] [--dry-run] [--json]',
  mutating: true,
  handler: (ctx) => runReplyAction(ctx, 'replyAll', 'Reply-all'),
};

export const forwardCommand: CommandSpec = {
  name: 'forward',
  summary: 'Forward a message to other recipients',
  usage:
    'outlook forward <messageId> --to <addr> [--to <addr>...] ' +
    '[--body <text>] [--signature <path> | --no-signature] [--dry-run] [--json]',
  mutating: true,
  handler: (ctx) => runReplyAction(ctx, 'forward', 'Forward'),
};
