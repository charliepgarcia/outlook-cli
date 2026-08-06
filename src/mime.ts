/**
 * Attachment encoding + content-type sniffing + HTML-to-text conversion.
 * Inline attachment size rule sourced from docs/02-write-endpoints.md §6:
 * "This operation limits the size of the attachment you can add to under 3 MB."
 */

import { readFile, stat } from 'node:fs/promises';
import { basename, extname } from 'node:path';
import { OutlookCliError } from './errors.js';
import type { FileAttachmentUpload } from './types.js';

/** Inline-attachment ceiling per docs/02 §6 — "under 3 MB." */
export const MAX_INLINE_ATTACHMENT_BYTES = 3 * 1024 * 1024;

const CONTENT_TYPES: Record<string, string> = {
  '.pdf': 'application/pdf',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.gif': 'image/gif',
  '.txt': 'text/plain',
  '.md': 'text/markdown',
  '.csv': 'text/csv',
  '.json': 'application/json',
  '.zip': 'application/zip',
  '.docx':
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  '.xlsx':
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  '.pptx':
    'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  '.html': 'text/html',
  '.htm': 'text/html',
};

const DEFAULT_CONTENT_TYPE = 'application/octet-stream';

/** Best-effort MIME sniff from the filename extension. Falls back to octet-stream. */
export function guessContentType(filename: string): string {
  const ext = extname(filename).toLowerCase();
  return CONTENT_TYPES[ext] ?? DEFAULT_CONTENT_TYPE;
}

function formatBytes(n: number): string {
  const mb = n / (1024 * 1024);
  return `${mb.toFixed(2)} MB (${n} bytes)`;
}

/**
 * Read a local file and shape it into a Graph fileAttachment POST body.
 * Throws OutlookCliError (not a silent truncation) if the file exceeds the
 * 3 MB inline-attachment ceiling — larger files need an upload session,
 * which is not implemented in v1 (docs/02 §6, "Large files — upload session").
 */
export async function readFileAttachment(
  path: string,
): Promise<FileAttachmentUpload> {
  const stats = await stat(path);

  if (stats.size > MAX_INLINE_ATTACHMENT_BYTES) {
    throw new OutlookCliError(
      `Attachment "${path}" is ${formatBytes(stats.size)}, which exceeds the ` +
        `3 MB inline-attachment limit. Files between 3 MB and 150 MB require ` +
        `a Graph upload session (createUploadSession), which this CLI does ` +
        `not yet support. Choose a smaller file or split the content.`,
    );
  }

  const buf = await readFile(path);
  const name = basename(path);

  return {
    '@odata.type': '#microsoft.graph.fileAttachment',
    name,
    contentType: guessContentType(name),
    contentBytes: buf.toString('base64'),
  };
}

const HTML_ENTITIES: Record<string, string> = {
  '&amp;': '&',
  '&lt;': '<',
  '&gt;': '>',
  '&quot;': '"',
  '&#39;': "'",
  '&nbsp;': ' ',
};

const ENTITY_RE = /&(?:amp|lt|gt|quot|#39|nbsp);/g;

/**
 * Convert HTML to readable plain text. Not a full parser — strips script/style
 * blocks entirely, maps <br>/</p> to newlines, strips remaining tags, decodes
 * the common named entities, and collapses runs of blank lines.
 */
export function htmlToText(html: string): string {
  if (!html) return '';

  let text = html
    .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '')
    .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '');

  text = text
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p\s*>/gi, '\n\n')
    .replace(/<\/(div|tr|li|h[1-6])\s*>/gi, '\n')
    .replace(/<[^>]+>/g, '');

  text = text.replace(ENTITY_RE, (m) => HTML_ENTITIES[m] ?? m);

  text = text
    .split('\n')
    .map((line) => line.replace(/[ \t]+$/, ''))
    .join('\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();

  return text;
}
