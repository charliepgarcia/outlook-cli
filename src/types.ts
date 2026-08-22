/**
 * Microsoft Graph v1.0 resource shapes — the subset this CLI touches.
 * Sourced from docs/01-read-endpoints.md and docs/02-write-endpoints.md.
 */

export interface EmailAddress {
  address: string;
  name?: string;
}

export interface Recipient {
  emailAddress: EmailAddress;
}

/**
 * Graph accepts both "Text"/"HTML" and "text"/"html" on write and echoes
 * lowercase on read. We author lowercase everywhere (docs/02).
 */
export type BodyContentType = 'text' | 'html';

export interface ItemBody {
  contentType: BodyContentType;
  content: string;
}

export type FlagStatus = 'notFlagged' | 'flagged' | 'complete';

export interface FollowupFlag {
  flagStatus: FlagStatus;
}

export type Importance = 'low' | 'normal' | 'high';

export interface Message {
  id: string;
  subject?: string | null;
  bodyPreview?: string;
  body?: ItemBody;
  from?: Recipient | null;
  sender?: Recipient | null;
  toRecipients?: Recipient[];
  ccRecipients?: Recipient[];
  bccRecipients?: Recipient[];
  replyTo?: Recipient[];
  receivedDateTime?: string;
  sentDateTime?: string;
  hasAttachments?: boolean;
  isRead?: boolean;
  isDraft?: boolean;
  importance?: Importance;
  flag?: FollowupFlag;
  parentFolderId?: string;
  conversationId?: string;
  webLink?: string;
  internetMessageId?: string;
}

/** Fields writable when creating or updating a draft (docs/02 §3). */
export interface DraftInput {
  subject?: string;
  body?: ItemBody;
  toRecipients?: Recipient[];
  ccRecipients?: Recipient[];
  bccRecipients?: Recipient[];
  importance?: Importance;
}

export interface MailFolder {
  id: string;
  displayName: string;
  parentFolderId?: string;
  childFolderCount?: number;
  unreadItemCount?: number;
  totalItemCount?: number;
}

export type AttachmentODataType =
  | '#microsoft.graph.fileAttachment'
  | '#microsoft.graph.itemAttachment'
  | '#microsoft.graph.referenceAttachment';

export interface Attachment {
  id: string;
  name: string;
  contentType?: string | null;
  size?: number;
  isInline?: boolean;
  lastModifiedDateTime?: string;
  '@odata.type'?: AttachmentODataType | string;
  /** base64 — present on fileAttachment when fetched directly. */
  contentBytes?: string;
}

/** Body shape for POSTing an inline attachment (< 3 MB, docs/02 §6). */
export interface FileAttachmentUpload {
  '@odata.type': '#microsoft.graph.fileAttachment';
  name: string;
  contentType: string;
  contentBytes: string;
}

export interface UserProfile {
  id: string;
  displayName?: string;
  mail?: string | null;
  userPrincipalName?: string;
  jobTitle?: string | null;
}

/** Envelope returned by every Graph collection endpoint. */
export interface GraphCollection<T> {
  '@odata.context'?: string;
  '@odata.nextLink'?: string;
  '@odata.count'?: number;
  value: T[];
}

/** Graph error envelope (docs/00 §8). Note both innerError casings occur. */
export interface GraphErrorBody {
  error: {
    code: string;
    message: string;
    innerError?: Record<string, unknown>;
    innererror?: Record<string, unknown>;
  };
}

/**
 * Well-known folder names usable in place of a folder id (docs/01 §2).
 * Verified against user-list-mailfolders.
 */
export const WELL_KNOWN_FOLDERS = [
  'archive',
  'clutter',
  'conflicts',
  'conversationhistory',
  'deleteditems',
  'drafts',
  'inbox',
  'junkemail',
  'localfailures',
  'msgfolderroot',
  'outbox',
  'recoverableitemsdeletions',
  'scheduled',
  'searchfolders',
  'sentitems',
  'serverfailures',
  'syncissues',
] as const;

export type WellKnownFolder = (typeof WELL_KNOWN_FOLDERS)[number];

export function isWellKnownFolder(name: string): boolean {
  return (WELL_KNOWN_FOLDERS as readonly string[]).includes(name.toLowerCase());
}

/** OAuth2 token set as persisted to the local 0600 cache. */
export interface TokenSet {
  accessToken: string;
  refreshToken?: string;
  /** Epoch milliseconds. */
  expiresAt: number;
  scope?: string;
  tokenType?: string;
  /** Signed-in identity, for `auth status`. */
  account?: string;
}

/** Response from the /devicecode endpoint (docs/00 §2). */
export interface DeviceCodeResponse {
  device_code: string;
  user_code: string;
  verification_uri: string;
  expires_in: number;
  interval: number;
  message?: string;
}

/** Response from the /token endpoint (docs/00 §3, §4). */
export interface TokenResponse {
  access_token: string;
  refresh_token?: string;
  expires_in: number;
  scope?: string;
  token_type?: string;
  id_token?: string;
}

/** A fully-described Graph request, used for --dry-run rendering. */
export interface RequestSpec {
  method: string;
  url: string;
  headers?: Record<string, string>;
  body?: unknown;
}

/** Resolved runtime configuration. */
export interface OutlookConfig {
  clientId: string;
  tenantId: string;
  clientSecret?: string;
  tokenCachePath: string;
}

export const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';

export const DEFAULT_SCOPES = [
  'offline_access',
  'User.Read',
  'Mail.ReadWrite',
  'Mail.Send',
  'MailboxSettings.ReadWrite',
] as const;
