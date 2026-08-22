/**
 * Auto-load environment variables from `~/.afk/config/afk.env`.
 *
 * Agent sessions often lack exported GRAPH_CLIENT_ID because the shell
 * profile was never sourced. SETUP.md tells users to persist it into this
 * file, but the CLI never read it — until now.
 *
 * Rules:
 *   1. If the file does not exist → no-op (not an error).
 *   2. Lines starting with `#` (after trimming) are comments → skip.
 *   3. Blank lines → skip.
 *   4. Lines matching `KEY=value` (with optional surrounding quotes on the
 *      value — single or double) set `process.env[KEY]` **only if that key
 *      is not already set**. Existing env vars always win.
 *
 * Zero runtime dependencies — uses only `node:fs` and `node:os`.
 */

import { readFileSync } from 'node:fs';
import { homedir } from 'node:os';
import { join } from 'node:path';

const ENV_PATH = join(homedir(), '.afk', 'config', 'afk.env');

/**
 * Parse lines from a `.env`-style file and return key-value pairs.
 * Exported for testing; the main API is `loadEnvFile`.
 */
export function parseEnvLines(content: string): [string, string][] {
  const pairs: [string, string][] = [];
  for (const raw of content.split('\n')) {
    const line = raw.trim();
    if (line === '' || line.startsWith('#')) continue;

    const eqIdx = line.indexOf('=');
    if (eqIdx === -1) continue;

    const key = line.slice(0, eqIdx).trim();
    if (key === '') continue;

    let value = line.slice(eqIdx + 1).trim();
    // Strip matched surrounding quotes (single or double).
    if (
      value.length >= 2 &&
      ((value.startsWith('"') && value.endsWith('"')) ||
        (value.startsWith("'") && value.endsWith("'")))
    ) {
      value = value.slice(1, -1);
    }

    pairs.push([key, value]);
  }
  return pairs;
}

/**
 * Read `~/.afk/config/afk.env` and inject any missing keys into
 * `process.env`. An explicit `envPath` can be supplied for testing.
 */
export function loadEnvFile(envPath: string = ENV_PATH): void {
  let content: string;
  try {
    content = readFileSync(envPath, 'utf-8');
  } catch {
    // File doesn't exist or isn't readable — nothing to do.
    return;
  }

  for (const [key, value] of parseEnvLines(content)) {
    if (process.env[key] === undefined) {
      process.env[key] = value;
    }
  }
}
