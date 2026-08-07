/**
 * Dependency-free argv parser.
 *
 * Convention (no schema is passed in, so this is deliberately simple and
 * documented rather than "smart"): a bare `--flag` greedily consumes the
 * following token as its string value UNLESS that token looks like another
 * flag (starts with `-`, length > 1) or is the `--` terminator. This means a
 * boolean-style flag immediately followed by a positional the caller wants
 * kept separate (e.g. `outlook list --json inbox`) will swallow "inbox" as
 * `--json`'s value rather than leaving it positional. `.bool()` still comes
 * back `true` in that case (presence-based), but the positional is lost.
 * Callers should order positionals before trailing boolean flags (the usual
 * CLI convention: `outlook list inbox --json`) or use `--` to force a hard
 * boundary. This tradeoff is inherent to schema-less argv parsing (the same
 * ambiguity minimist has) and is called out here rather than silently
 * assumed.
 */

import { UsageError } from './errors.js';

export interface ParsedArgs {
  positionals: string[];
  raw: string[];
  has(name: string): boolean;
  str(name: string, fallback?: string): string | undefined;
  strAll(name: string): string[];
  num(name: string): number | undefined;
  bool(name: string): boolean;
  require(name: string): string;
  requirePositional(index: number, label: string): string;
  /** Every flag name seen on the command line. Drives unknown-flag rejection. */
  names(): string[];
}

type FlagValue = string | boolean;

/**
 * Explicit boolean spellings accepted by `.bool()`. Anything else supplied as
 * a value to a boolean flag is a UsageError rather than a silent coercion —
 * see the `bool` implementation for why that matters for `--yes`.
 */
const TRUE_VALUES = new Set(['true', '1', 'yes', 'on']);
const FALSE_VALUES = new Set(['false', '0', 'no', 'off', '']);

/** `-` alone is treated as a value (common "stdin" convention), not a flag. */
function looksLikeOption(tok: string | undefined): boolean {
  if (tok === undefined) return false;
  if (tok === '-') return false;
  return tok.startsWith('-') && tok.length > 1;
}

export function parseArgs(argv: string[]): ParsedArgs {
  const positionals: string[] = [];
  const store = new Map<string, FlagValue[]>();

  function record(name: string, value: FlagValue): void {
    const existing = store.get(name);
    if (existing) existing.push(value);
    else store.set(name, [value]);
  }

  let i = 0;
  let terminated = false;

  while (i < argv.length) {
    const tok = argv[i];
    if (tok === undefined) break;

    if (!terminated && tok === '--') {
      terminated = true;
      i++;
      continue;
    }

    if (terminated) {
      positionals.push(tok);
      i++;
      continue;
    }

    if (tok.startsWith('--') && tok.length > 2) {
      const body = tok.slice(2);
      const eq = body.indexOf('=');

      if (eq !== -1) {
        record(body.slice(0, eq), body.slice(eq + 1));
        i++;
        continue;
      }

      if (body.startsWith('no-') && body.length > 3) {
        record(body.slice(3), false);
        i++;
        continue;
      }

      const next = argv[i + 1];
      if (next !== undefined && !looksLikeOption(next)) {
        record(body, next);
        i += 2;
        continue;
      }

      record(body, true);
      i++;
      continue;
    }

    if (tok.startsWith('-') && tok.length > 1) {
      const body = tok.slice(1);
      const eq = body.indexOf('=');

      if (eq !== -1) {
        record(body.slice(0, eq), body.slice(eq + 1));
        i++;
        continue;
      }

      const next = argv[i + 1];
      if (next !== undefined && !looksLikeOption(next)) {
        record(body, next);
        i += 2;
        continue;
      }

      record(body, true);
      i++;
      continue;
    }

    positionals.push(tok);
    i++;
  }

  function lastValue(name: string): FlagValue | undefined {
    const list = store.get(name);
    if (!list || list.length === 0) return undefined;
    return list[list.length - 1];
  }

  const api: ParsedArgs = {
    positionals,
    raw: [...argv],

    has(name) {
      return store.has(name);
    },

    str(name, fallback) {
      const v = lastValue(name);
      return typeof v === 'string' ? v : fallback;
    },

    strAll(name) {
      const list = store.get(name);
      if (!list) return [];
      return list.filter((v): v is string => typeof v === 'string');
    },

    num(name) {
      const v = lastValue(name);
      if (v === undefined) return undefined;
      if (typeof v === 'boolean') {
        throw new UsageError(`--${name} requires a numeric value`);
      }
      if (v.trim() === '') {
        throw new UsageError(`--${name} expects a numeric value, got ""`);
      }
      const n = Number(v);
      if (Number.isNaN(n)) {
        throw new UsageError(`--${name} expects a numeric value, got "${v}"`);
      }
      return n;
    },

    /**
     * Presence-based, but NEVER value-blind.
     *
     * The old implementation returned `last !== false`, which meant any string
     * value was truthy — so `--yes=false` (a user trying to DECLINE an
     * irreversible purge) read as `true` and the deletion proceeded. Explicit
     * false spellings must win, and anything ambiguous must be a loud error
     * rather than a guess, because the guess that gets this wrong deletes
     * mail. A bare `--flag` is still `true`; `--no-flag` is still `false`.
     */
    bool(name) {
      const v = lastValue(name);
      if (v === undefined) return false;
      if (typeof v === 'boolean') return v;

      const normalized = v.trim().toLowerCase();
      if (TRUE_VALUES.has(normalized)) return true;
      if (FALSE_VALUES.has(normalized)) return false;

      // Almost always the schema-less greedy-consumption footgun documented at
      // the top of this file: `outlook list --json inbox` swallowed "inbox".
      // That used to silently drop the positional; now it says so.
      throw new UsageError(
        `--${name} does not take a value (got "${v}").\n` +
          `If "${v}" is a positional argument, put it before the flag: ` +
          `\`... ${v} --${name}\`.`,
      );
    },

    require(name) {
      const v = lastValue(name);
      if (typeof v === 'string') return v;
      throw new UsageError(`Missing required flag: --${name}`);
    },

    requirePositional(index, label) {
      const v = positionals[index];
      if (v === undefined) {
        throw new UsageError(`Missing required argument: ${label}`);
      }
      return v;
    },

    names() {
      return [...store.keys()];
    },
  };

  return api;
}
