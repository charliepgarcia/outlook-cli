/**
 * Pure renderers: every function here returns a string. Nothing in this file
 * calls console.log — cli.ts owns printing. That split is what makes these
 * testable without capturing stdout.
 *
 * Graph omits fields liberally (S1/S8, docs/01-read-endpoints.md §6), so every
 * renderer treats optional/null fields as expected input, not an error case.
 */

import type {
  Attachment,
  MailFolder,
  Message,
  Recipient,
  UserProfile,
} from './types.js';
import { htmlToText } from './mime.js';

export interface RenderOpts {
  json: boolean;
}

const DATE_WIDTH = 20;
const FROM_WIDTH = 24;
const SUBJECT_WIDTH = 40;

export function renderJson(data: unknown): string {
  return JSON.stringify(data, null, 2);
}

/** "Name <address>" when a name is present, else the bare address. undefined if unresolvable. */
function recipientLabel(recipient: Recipient | null | undefined): string | undefined {
  const address = recipient?.emailAddress?.address;
  if (!address) return undefined;
  const name = recipient.emailAddress.name;
  return name ? `${name} <${address}>` : address;
}

export function formatRecipients(recipients?: Recipient[]): string {
  if (!recipients || recipients.length === 0) return '(none)';
  return recipients.map((r) => recipientLabel(r) ?? '(unknown)').join(', ');
}

function pad2(n: number): string {
  return n.toString().padStart(2, '0');
}

/** Deterministic UTC rendering so output doesn't vary with the host timezone. */
export function formatDate(iso?: string): string {
  if (!iso) return '(no date)';
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return '(no date)';
  const y = d.getUTCFullYear();
  const mo = pad2(d.getUTCMonth() + 1);
  const day = pad2(d.getUTCDate());
  const h = pad2(d.getUTCHours());
  const mi = pad2(d.getUTCMinutes());
  return `${y}-${mo}-${day} ${h}:${mi} UTC`;
}

export function truncate(text: string, max: number): string {
  if (max <= 0) return '';
  if (text.length <= max) return text;
  if (max === 1) return text.slice(0, 1);
  return `${text.slice(0, max - 1)}\u2026`;
}

function subjectOf(message: Message): string {
  return message.subject && message.subject.trim().length > 0
    ? message.subject
    : '(no subject)';
}

function fromOf(message: Message): string {
  return recipientLabel(message.from ?? message.sender) ?? '(unknown sender)';
}

/**
 * Aligned columns: unread marker, date, from, subject, attachment marker, id.
 * `*` marks isRead === false; blank covers read or unknown (Graph can omit isRead).
 * `@` marks hasAttachments === true.
 */
export function renderMessageList(messages: Message[]): string {
  if (messages.length === 0) return '(no messages)';

  const header = [
    ' '.padEnd(1),
    'DATE'.padEnd(DATE_WIDTH),
    'FROM'.padEnd(FROM_WIDTH),
    'SUBJECT'.padEnd(SUBJECT_WIDTH),
    ' '.padEnd(1),
    'ID',
  ].join(' ');

  const rows = messages.map((m) => {
    const unread = m.isRead === false ? '*' : ' ';
    const attach = m.hasAttachments ? '@' : ' ';
    const date = formatDate(m.receivedDateTime ?? m.sentDateTime);
    const from = truncate(fromOf(m), FROM_WIDTH);
    const subject = truncate(subjectOf(m), SUBJECT_WIDTH);
    return [
      unread,
      date.padEnd(DATE_WIDTH),
      from.padEnd(FROM_WIDTH),
      subject.padEnd(SUBJECT_WIDTH),
      attach,
      m.id,
    ].join(' ');
  });

  return [header, ...rows].join('\n');
}

function readState(message: Message): string {
  if (message.isRead === false) return 'no';
  if (message.isRead === true) return 'yes';
  return 'unknown';
}

function renderBody(message: Message, opts?: { html?: boolean }): string {
  const body = message.body;
  if (!body || !body.content) {
    return message.bodyPreview && message.bodyPreview.trim().length > 0
      ? message.bodyPreview
      : '(no body)';
  }
  if (body.contentType === 'html') {
    return opts?.html ? body.content : htmlToText(body.content);
  }
  return body.content;
}

/** Full headers, then a blank line, then the body (plain text unless opts.html). */
export function renderMessage(message: Message, opts?: { html?: boolean }): string {
  const lines: string[] = [
    `Subject: ${subjectOf(message)}`,
    `From: ${fromOf(message)}`,
    `To: ${formatRecipients(message.toRecipients)}`,
  ];

  if (message.ccRecipients && message.ccRecipients.length > 0) {
    lines.push(`Cc: ${formatRecipients(message.ccRecipients)}`);
  }

  lines.push(
    `Date: ${formatDate(message.receivedDateTime ?? message.sentDateTime)}`,
    `Importance: ${message.importance ?? 'normal'}`,
    `Read: ${readState(message)}`,
    `Attachments: ${message.hasAttachments ? 'yes' : 'no'}`,
    `Id: ${message.id}`,
    '',
    renderBody(message, opts),
  );

  return lines.join('\n');
}

export function renderFolders(folders: MailFolder[]): string {
  if (folders.length === 0) return '(no folders)';

  return folders
    .map((f) => {
      const name = f.displayName && f.displayName.trim() ? f.displayName : '(unnamed folder)';
      const counts: string[] = [];
      if (f.unreadItemCount !== undefined) counts.push(`${f.unreadItemCount} unread`);
      if (f.totalItemCount !== undefined) counts.push(`${f.totalItemCount} total`);
      const suffix = counts.length > 0 ? ` (${counts.join(', ')})` : '';
      return `${name}${suffix} — ${f.id}`;
    })
    .join('\n');
}

/** Strips a leading '#' and the 'microsoft.graph.' prefix (S9's own example omits the '#'). */
function attachmentKind(odataType?: string): string {
  if (!odataType) return 'unknown';
  const normalized = odataType.replace(/^#/, '');
  switch (normalized) {
    case 'microsoft.graph.fileAttachment':
      return 'file';
    case 'microsoft.graph.itemAttachment':
      return 'item';
    case 'microsoft.graph.referenceAttachment':
      return 'reference';
    default:
      return normalized;
  }
}

function formatSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

export function renderAttachments(attachments: Attachment[]): string {
  if (attachments.length === 0) return '(no attachments)';

  return attachments
    .map((a) => {
      const name = a.name && a.name.trim() ? a.name : '(unnamed attachment)';
      const kind = attachmentKind(a['@odata.type']);
      const size = a.size !== undefined ? formatSize(a.size) : '(unknown size)';
      const inline = a.isInline ? ' [inline]' : '';
      return `${name} — ${kind}, ${size}${inline} — ${a.id}`;
    })
    .join('\n');
}

export function renderProfile(profile: UserProfile): string {
  const lines = [
    `Name: ${profile.displayName ?? '(unknown)'}`,
    `Mail: ${profile.mail ?? profile.userPrincipalName ?? '(none)'}`,
  ];
  if (profile.jobTitle) lines.push(`Job title: ${profile.jobTitle}`);
  lines.push(`Id: ${profile.id}`);
  return lines.join('\n');
}
