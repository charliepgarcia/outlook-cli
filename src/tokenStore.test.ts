import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { mkdtemp, readdir, rm, stat, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { createTokenStore } from './tokenStore.js';
import type { TokenSet } from './types.js';

describe('tokenStore', () => {
  let dir: string;

  beforeEach(async () => {
    dir = await mkdtemp(join(tmpdir(), 'outlook-cli-test-'));
  });

  afterEach(async () => {
    await rm(dir, { recursive: true, force: true });
  });

  const sampleTokens: TokenSet = {
    accessToken: 'access-abc',
    refreshToken: 'refresh-xyz',
    expiresAt: 1_700_000_000_000,
    scope: 'User.Read Mail.Send',
    tokenType: 'Bearer',
    account: 'user@example.com',
  };

  it('load() returns null when the file does not exist', async () => {
    const store = createTokenStore(join(dir, 'nested', 'token.json'));
    await expect(store.load()).resolves.toBeNull();
  });

  it('save() then load() round-trips the token set exactly', async () => {
    const store = createTokenStore(join(dir, 'token.json'));
    await store.save(sampleTokens);
    await expect(store.load()).resolves.toEqual(sampleTokens);
  });

  it('save() creates the parent directory owner-only (0700), not world-readable', async () => {
    // The file is already 0600; this stops the directory from advertising that
    // a token cache exists and when it was last written.
    const path = join(dir, 'fresh', 'token.json');
    await createTokenStore(path).save(sampleTokens);
    const st = await stat(join(dir, 'fresh'));
    expect(st.mode & 0o777).toBe(0o700);
  });

  it('save() creates the parent directory when missing (mkdir -p)', async () => {
    const path = join(dir, 'deeply', 'nested', 'dir', 'token.json');
    const store = createTokenStore(path);
    await store.save(sampleTokens);
    const st = await stat(path);
    expect(st.isFile()).toBe(true);
  });

  it('save() writes the file chmod 0600', async () => {
    const path = join(dir, 'token.json');
    await createTokenStore(path).save(sampleTokens);
    const st = await stat(path);
    expect(st.mode & 0o777).toBe(0o600);
  });

  it('save() forces 0600 even over a pre-existing, more permissive file', async () => {
    const path = join(dir, 'token.json');
    await writeFile(path, '{}', { mode: 0o644 });
    expect((await stat(path)).mode & 0o777).toBe(0o644); // sanity: confirm the loose starting mode

    await createTokenStore(path).save(sampleTokens);

    expect((await stat(path)).mode & 0o777).toBe(0o600);
  });

  it('load() returns null for corrupt JSON rather than throwing', async () => {
    const path = join(dir, 'token.json');
    await writeFile(path, '{ not valid json', { mode: 0o600 });
    await expect(createTokenStore(path).load()).resolves.toBeNull();
  });

  it('load() returns null for well-formed JSON that is not a TokenSet', async () => {
    const path = join(dir, 'token.json');
    await writeFile(path, JSON.stringify({ foo: 'bar' }), { mode: 0o600 });
    await expect(createTokenStore(path).load()).resolves.toBeNull();
  });

  it('load() returns null when accessToken is missing', async () => {
    const path = join(dir, 'token.json');
    await writeFile(path, JSON.stringify({ expiresAt: 123 }), { mode: 0o600 });
    await expect(createTokenStore(path).load()).resolves.toBeNull();
  });

  it('load() returns null when expiresAt is not a number', async () => {
    const path = join(dir, 'token.json');
    await writeFile(
      path,
      JSON.stringify({ accessToken: 'a', expiresAt: 'soon' }),
      { mode: 0o600 },
    );
    await expect(createTokenStore(path).load()).resolves.toBeNull();
  });

  it('clear() removes an existing file', async () => {
    const path = join(dir, 'token.json');
    const store = createTokenStore(path);
    await store.save(sampleTokens);
    await store.clear();
    await expect(store.load()).resolves.toBeNull();
  });

  it('clear() does not throw when the file is already absent', async () => {
    const store = createTokenStore(join(dir, 'never-created.json'));
    await expect(store.clear()).resolves.toBeUndefined();
  });

  it('exposes the configured path', () => {
    const path = join(dir, 'token.json');
    expect(createTokenStore(path).path).toBe(path);
  });
});

/**
 * Security fix regression coverage (docs/security-review-2026-08-07.md #4):
 * `save()` now writes to a `.{pid}-{ts}.token.tmp` sibling and renames it over
 * the target, so a process death mid-write can never leave a truncated or
 * half-merged cache behind.
 */
describe('tokenStore — atomic save', () => {
  let dir: string;

  beforeEach(async () => {
    dir = await mkdtemp(join(tmpdir(), 'outlook-cli-test-atomic-'));
  });

  afterEach(async () => {
    await rm(dir, { recursive: true, force: true });
  });

  const sampleTokens: TokenSet = {
    accessToken: 'access-abc',
    refreshToken: 'refresh-xyz',
    expiresAt: 1_700_000_000_000,
    scope: 'User.Read Mail.Send',
    tokenType: 'Bearer',
    account: 'user@example.com',
  };

  it('leaves no leftover .token.tmp file after save()', async () => {
    const path = join(dir, 'token.json');
    await createTokenStore(path).save(sampleTokens);

    expect(await readdir(dir)).toEqual(['token.json']);
  });

  it('saving over an existing token file replaces it completely, not a merge', async () => {
    const path = join(dir, 'token.json');
    const store = createTokenStore(path);
    await store.save({
      accessToken: 'first',
      expiresAt: 1,
      refreshToken: 'r1',
      scope: 'a b c',
      tokenType: 'Bearer',
      account: 'first@example.com',
    });

    const replacement: TokenSet = { accessToken: 'second', expiresAt: 2 };
    await store.save(replacement);

    await expect(store.load()).resolves.toEqual(replacement);
    expect(await readdir(dir)).toEqual(['token.json']);
  });

  it.skipIf(process.platform === 'win32')(
    'the renamed-into-place file is still chmod 0600 (POSIX only)',
    async () => {
      const path = join(dir, 'token.json');
      await createTokenStore(path).save(sampleTokens);

      const st = await stat(path);
      expect(st.mode & 0o777).toBe(0o600);
    },
  );
});

/**
 * `mode` field coverage (docs/security-review-2026-08-07.md #3): a corrupt or
 * unrecognized `mode` must be REJECTED by load(), not defaulted to 'full' —
 * silently upgrading an unreadable mode to full access would be a privilege
 * escalation on a corrupt cache.
 */
describe('tokenStore — mode field', () => {
  let dir: string;

  beforeEach(async () => {
    dir = await mkdtemp(join(tmpdir(), 'outlook-cli-test-mode-'));
  });

  afterEach(async () => {
    await rm(dir, { recursive: true, force: true });
  });

  it('save() then load() round-trips mode: "readonly"', async () => {
    const path = join(dir, 'token.json');
    const tokens: TokenSet = { accessToken: 'a', expiresAt: 1, mode: 'readonly' };
    const store = createTokenStore(path);

    await store.save(tokens);

    await expect(store.load()).resolves.toEqual(tokens);
  });

  it('save() then load() round-trips mode: "full"', async () => {
    const path = join(dir, 'token.json');
    const tokens: TokenSet = { accessToken: 'a', expiresAt: 1, mode: 'full' };
    const store = createTokenStore(path);

    await store.save(tokens);

    await expect(store.load()).resolves.toEqual(tokens);
  });

  it('load() rejects an unrecognized mode string instead of defaulting it to "full"', async () => {
    const path = join(dir, 'token.json');
    await writeFile(
      path,
      JSON.stringify({ accessToken: 'a', expiresAt: 123, mode: 'admin' }),
      { mode: 0o600 },
    );

    await expect(createTokenStore(path).load()).resolves.toBeNull();
  });

  it('load() accepts a cache with no mode field at all (pre-read-only-mode tokens)', async () => {
    const path = join(dir, 'token.json');
    const legacy = { accessToken: 'a', expiresAt: 123 };
    await writeFile(path, JSON.stringify(legacy), { mode: 0o600 });

    await expect(createTokenStore(path).load()).resolves.toEqual(legacy);
  });
});
