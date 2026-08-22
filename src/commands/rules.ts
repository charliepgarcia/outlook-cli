/**
 * `outlook rules list` — list inbox message rules.
 *
 * Graph endpoint: `GET /me/mailFolders/inbox/messageRules`
 * (see https://learn.microsoft.com/en-us/graph/api/mailfolder-list-messagerules)
 *
 * Each messageRule has: displayName, sequence, isEnabled, isReadOnly,
 * conditions, actions, exceptions. This is a read-only command — no
 * guardMutation needed, no --dry-run.
 *
 * Subcommand dispatch on positional 0 (`list` is the default):
 *   outlook rules [list] [--json]
 *
 * Human output renders a compact table showing each rule's sequence number,
 * name, enabled state, a brief conditions summary, and a brief actions summary.
 */

import { emit, type CommandContext, type CommandSpec } from '../command.js';
import type { GraphCollection } from '../types.js';

/* ---------- Graph response types (subset we render) ---------- */

interface RuleEmailAddress {
  emailAddress?: { address?: string; name?: string };
}

/** Conditions or exceptions on a message rule. */
export interface RuleConditions {
  fromAddresses?: RuleEmailAddress[];
  sentToAddresses?: RuleEmailAddress[];
  subjectContains?: string[];
  bodyContains?: string[];
  headerContains?: string[];
  senderContains?: string[];
  recipientContains?: string[];
  bodyOrSubjectContains?: string[];
  hasAttachments?: boolean;
  importance?: string;
  isAutomatic?: boolean;
  [key: string]: unknown;
}

/** Actions performed when a rule matches. */
export interface RuleActions {
  moveToFolder?: string;
  copyToFolder?: string;
  delete?: boolean;
  permanentDelete?: boolean;
  markAsRead?: boolean;
  markImportance?: string;
  forwardTo?: RuleEmailAddress[];
  forwardAsAttachmentTo?: RuleEmailAddress[];
  redirectTo?: RuleEmailAddress[];
  stopProcessingRules?: boolean;
  [key: string]: unknown;
}

export interface MessageRule {
  id?: string;
  displayName?: string;
  sequence?: number;
  isEnabled?: boolean;
  isReadOnly?: boolean;
  conditions?: RuleConditions;
  actions?: RuleActions;
  exceptions?: RuleConditions;
}

/* ---------- Human-readable summarisers ---------- */

function addrList(addrs?: RuleEmailAddress[]): string {
  if (!addrs || addrs.length === 0) return '';
  return addrs
    .map((a) => a.emailAddress?.address ?? a.emailAddress?.name ?? '?')
    .join(', ');
}

/** Short human string describing which messages match. */
export function summariseConditions(c?: RuleConditions): string {
  if (!c) return '(any)';
  const parts: string[] = [];

  const from = addrList(c.fromAddresses);
  if (from) parts.push(`from: ${from}`);

  const to = addrList(c.sentToAddresses);
  if (to) parts.push(`to: ${to}`);

  if (c.subjectContains?.length) parts.push(`subject contains: ${c.subjectContains.join(', ')}`);
  if (c.bodyContains?.length) parts.push(`body contains: ${c.bodyContains.join(', ')}`);
  if (c.senderContains?.length) parts.push(`sender contains: ${c.senderContains.join(', ')}`);
  if (c.hasAttachments) parts.push('has attachments');
  if (c.importance) parts.push(`importance: ${c.importance}`);

  return parts.length > 0 ? parts.join('; ') : '(any)';
}

/** Short human string describing what the rule does. */
export function summariseActions(a?: RuleActions): string {
  if (!a) return '(none)';
  const parts: string[] = [];

  if (a.moveToFolder) parts.push(`move to [${a.moveToFolder}]`);
  if (a.copyToFolder) parts.push(`copy to [${a.copyToFolder}]`);
  if (a.delete) parts.push('delete');
  if (a.permanentDelete) parts.push('permanent delete');
  if (a.markAsRead) parts.push('mark read');
  if (a.markImportance) parts.push(`mark ${a.markImportance}`);

  const fwd = addrList(a.forwardTo);
  if (fwd) parts.push(`forward to: ${fwd}`);
  const redir = addrList(a.redirectTo);
  if (redir) parts.push(`redirect to: ${redir}`);

  if (a.stopProcessingRules) parts.push('stop processing');

  return parts.length > 0 ? parts.join('; ') : '(none)';
}

export function renderRules(rules: MessageRule[]): string {
  if (rules.length === 0) return '(no rules)';

  const header = [
    '#'.padEnd(3),
    'ENABLED'.padEnd(8),
    'NAME'.padEnd(28),
    'CONDITIONS'.padEnd(36),
    'ACTIONS',
  ].join(' ');

  const rows = rules.map((r) => [
    String(r.sequence ?? '-').padEnd(3),
    (r.isEnabled ? 'yes' : 'no').padEnd(8),
    (r.displayName ?? '(unnamed)').slice(0, 28).padEnd(28),
    summariseConditions(r.conditions).slice(0, 36).padEnd(36),
    summariseActions(r.actions),
  ].join(' '));

  return [header, ...rows].join('\n');
}

/* ---------- Command handler ---------- */

const RULES_PATH = '/me/mailFolders/inbox/messageRules';

async function rulesHandler(ctx: CommandContext): Promise<void> {
  const sub = ctx.args.positionals[0] ?? 'list';

  if (sub !== 'list') {
    const { UsageError } = await import('../errors.js');
    throw new UsageError(`Unknown rules subcommand "${sub}". Expected "list".`);
  }

  const response = await ctx.client.get<GraphCollection<MessageRule>>(RULES_PATH);
  const rules = response.value;

  emit(ctx, rules, () => renderRules(rules));
}

export const rulesCommand: CommandSpec = {
  name: 'rules',
  summary: 'List inbox message rules',
  usage: 'outlook rules [list] [--json]',
  mutating: false,
  handler: rulesHandler,
};
