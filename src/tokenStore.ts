/**
 * Local token cache: a single JSON file, chmod 0600. A corrupt or unreadable
 * cache must never crash the CLI — it just means "sign in again," so `load`
 * swallows every failure mode (missing file, bad JSON, wrong shape) and
 * returns `null` rather than throwing.
 */

import { chmod, mkdir, readFile, rm, writeFile } from 'node:fs/promises';
import { dirname } from 'node:path';
import type { TokenSet } from './types.js';

/** Owner read/write only — the file holds an access + refresh token. */
const FILE_MODE = 0o600;

export interface TokenStore {
  readonly path: string;
  load(): Promise<TokenSet | null>;
  save(tokens: TokenSet): Promise<void>;
  clear(): Promise<void>;
}

export function createTokenStore(path: string): TokenStore {
  return {
    path,

    async load(): Promise<TokenSet | null> {
      try {
        const raw = await readFile(path, 'utf8');
        const parsed: unknown = JSON.parse(raw);
        return isTokenSet(parsed) ? parsed : null;
      } catch {
        // ENOENT, EACCES, malformed JSON, a directory where a file should
        // be — all of it collapses to "no usable cache." Never throw.
        return null;
      }
    },

    async save(tokens: TokenSet): Promise<void> {
      await mkdir(dirname(path), { recursive: true });
      // `mode` on writeFile only takes effect when the file is *created*;
      // if it already exists with looser permissions, writeFile leaves them
      // alone. The explicit chmod is what actually guarantees 0600 either way.
      await writeFile(path, JSON.stringify(tokens, null, 2), { mode: FILE_MODE });
      await chmod(path, FILE_MODE);
    },

    async clear(): Promise<void> {
      // force: true swallows ENOENT — clearing an already-absent cache is a
      // no-op, not an error.
      await rm(path, { force: true });
    },
  };
}

/** Runtime shape guard — a corrupt cache must degrade to `null`, not a crash. */
function isTokenSet(value: unknown): value is TokenSet {
  if (typeof value !== 'object' || value === null) return false;
  const v = value as Record<string, unknown>;

  if (typeof v['accessToken'] !== 'string' || v['accessToken'].length === 0) return false;
  if (typeof v['expiresAt'] !== 'number' || !Number.isFinite(v['expiresAt'])) return false;

  if (v['refreshToken'] !== undefined && typeof v['refreshToken'] !== 'string') return false;
  if (v['scope'] !== undefined && typeof v['scope'] !== 'string') return false;
  if (v['tokenType'] !== undefined && typeof v['tokenType'] !== 'string') return false;
  if (v['account'] !== undefined && typeof v['account'] !== 'string') return false;

  return true;
}
