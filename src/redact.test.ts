import { describe, it, expect } from 'vitest';
import { MASK, redactObject } from './redact.js';

/** Wraps `leaf` in `depth` nested `{ wrapper: ... }` objects. */
function nestAt(depth: number, leaf: unknown): unknown {
  let obj: unknown = leaf;
  for (let i = 0; i < depth; i++) obj = { wrapper: obj };
  return obj;
}

/** Walks `depth` `.wrapper` hops down from `redactObject`'s output. */
function unwrap(value: unknown, depth: number): unknown {
  let cursor = value;
  for (let i = 0; i < depth; i++) cursor = (cursor as { wrapper: unknown }).wrapper;
  return cursor;
}

describe('redactObject — recursion depth cap', () => {
  /**
   * Regression coverage (docs/security-review-2026-08-07.md): the depth
   * guard used to `return value` past the cap, which handed back the
   * untouched subtree — a secret nested one level deeper than the cap
   * printed raw. It must mask instead.
   */
  it('masks a secret-keyed field at exactly the cap (depth 12) as before', () => {
    const input = nestAt(12, { access_token: 'SECRET_VALUE_LONG_ENOUGH_TO_MATTER' });
    const out = redactObject(input);
    const leaf = unwrap(out, 12) as { access_token: string };
    expect(leaf.access_token).toContain(MASK);
    expect(leaf.access_token).not.toContain('SECRET_VALUE_LONG_ENOUGH_TO_MATTER');
  });

  it('masks the whole subtree — not returns it raw — once past the depth cap (depth 13)', () => {
    const input = nestAt(13, { access_token: 'SECRET_VALUE_LONG_ENOUGH_TO_MATTER' });
    const out = redactObject(input);
    const atCap = unwrap(out, 13);

    // Previously: the object `{ access_token: '...' }` came back byte-for-byte,
    // so JSON.stringify(out) contained the raw secret. Now the whole subtree
    // collapses to the MASK string.
    expect(atCap).toBe(MASK);
    expect(JSON.stringify(out)).not.toContain('SECRET_VALUE_LONG_ENOUGH_TO_MATTER');
  });
});
