/**
 * Dependency-free argv parser.
 *
 * Production callers provide an OptionSchema. In that mode unknown options are
 * rejected, value-taking options require a value, and booleans accept only a
 * bare flag, `--no-flag`, or an explicit `true`/`false` value. Without a schema
 * the original permissive parsing behavior is retained for low-level callers.
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
}

type FlagValue = string | boolean;

export interface OptionSchema {
  boolean: readonly string[];
  value: readonly string[];
  aliases?: Readonly<Record<string, string>>;
}

/** `-` alone is treated as a value (common "stdin" convention), not a flag. */
function looksLikeOption(tok: string | undefined): boolean {
  if (tok === undefined) return false;
  if (tok === '-') return false;
  return tok.startsWith('-') && tok.length > 1;
}

export function parseArgs(argv: string[], schema?: OptionSchema): ParsedArgs {
  const positionals: string[] = [];
  const store = new Map<string, FlagValue[]>();
  const booleanOptions = new Set(schema?.boolean ?? []);
  const valueOptions = new Set(schema?.value ?? []);

  function option(
    rawName: string,
    spelling: string,
  ): { name: string; kind: 'boolean' | 'value' } {
    const name = schema?.aliases?.[rawName] ?? rawName;
    if (booleanOptions.has(name)) return { name, kind: 'boolean' };
    if (valueOptions.has(name)) return { name, kind: 'value' };
    throw new UsageError(`Unknown option: ${spelling}`);
  }

  function booleanValue(value: string, spelling: string): boolean {
    if (value === 'true') return true;
    if (value === 'false') return false;
    throw new UsageError(`${spelling} expects true or false`);
  }

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
        const rawName = body.slice(0, eq);
        const value = body.slice(eq + 1);
        if (schema) {
          const resolved = option(rawName, `--${rawName}`);
          record(
            resolved.name,
            resolved.kind === 'boolean'
              ? booleanValue(value, `--${rawName}`)
              : value,
          );
        } else {
          record(rawName, value);
        }
        i++;
        continue;
      }

      if (body.startsWith('no-') && body.length > 3) {
        const rawName = body.slice(3);
        if (schema) {
          const resolved = option(rawName, `--no-${rawName}`);
          if (resolved.kind !== 'boolean') {
            throw new UsageError(`Unknown option: --no-${rawName}`);
          }
          record(resolved.name, false);
        } else {
          record(rawName, false);
        }
        i++;
        continue;
      }

      if (schema) {
        const resolved = option(body, `--${body}`);
        const next = argv[i + 1];
        if (resolved.kind === 'boolean') {
          if (next === 'true' || next === 'false') {
            record(resolved.name, booleanValue(next, `--${body}`));
            i += 2;
          } else {
            record(resolved.name, true);
            i++;
          }
          continue;
        }

        if (next === undefined || looksLikeOption(next)) {
          throw new UsageError(`--${body} requires a value`);
        }
        record(resolved.name, next);
        i += 2;
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
        const rawName = body.slice(0, eq);
        const value = body.slice(eq + 1);
        if (schema) {
          const resolved = option(rawName, `-${rawName}`);
          record(
            resolved.name,
            resolved.kind === 'boolean'
              ? booleanValue(value, `-${rawName}`)
              : value,
          );
        } else {
          record(rawName, value);
        }
        i++;
        continue;
      }

      if (schema) {
        const resolved = option(body, `-${body}`);
        const next = argv[i + 1];
        if (resolved.kind === 'boolean') {
          if (next === 'true' || next === 'false') {
            record(resolved.name, booleanValue(next, `-${body}`));
            i += 2;
          } else {
            record(resolved.name, true);
            i++;
          }
          continue;
        }

        if (next === undefined || looksLikeOption(next)) {
          throw new UsageError(`-${body} requires a value`);
        }
        record(resolved.name, next);
        i += 2;
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

    bool(name) {
      const list = store.get(name);
      if (!list || list.length === 0) return false;
      return list[list.length - 1] !== false;
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
  };

  return api;
}
