/**
 * The atomic-write path's fail-closed properties (review finding F-A4).
 *
 * `save()` writes a temp file and renames it over the target. The temp name
 * used to be `.${pid}-${Date.now()}.token.tmp` — guessable — and writeFile's
 * default O_CREAT|O_TRUNC FOLLOWS a symlink, so a link pre-planted at that
 * path would have received the access + refresh token at whatever destination
 * it pointed to. The name is now random and the open is exclusive ('wx').
 *
 * Asserting the open FLAG needs the fs module mocked, which would leak into
 * every other test in tokenStore.test.ts — hence a separate file (the same
 * split-by-concern reason attachments.download.test.ts exists).
 */
import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';

const writeFileSpy = vi.fn();

vi.mock('node:fs/promises', async () => {
  const actual = await vi.importActual<typeof import('node:fs/promises')>('node:fs/promises');
  return {
    ...actual,
    // Record the call, then write through for real so rename/stat still work.
    writeFile: (...args: Parameters<typeof actual.writeFile>) => {
      writeFileSpy(...args);
      return actual.writeFile(...args);
    },
  };
});

const { mkdtemp, rm, readdir } = await import('node:fs/promises');
const { tmpdir } = await import('node:os');
const { join } = await import('node:path');
const { createTokenStore } = await import('./tokenStore.js');
type TokenSet = import('./types.js').TokenSet;

const sampleTokens: TokenSet = {
  accessToken: 'access-abc',
  refreshToken: 'refresh-xyz',
  expiresAt: 1_700_000_000_000,
};

describe('tokenStore.save — fail-closed temp file', () => {
  let dir: string;

  beforeEach(async () => {
    writeFileSpy.mockClear();
    dir = await mkdtemp(join(tmpdir(), 'outlook-atomic-'));
  });

  afterEach(async () => {
    await rm(dir, { recursive: true, force: true });
  });

  it("opens the temp file with 'wx', so a pre-planted symlink is refused rather than followed", async () => {
    const store = createTokenStore(join(dir, 'token.json'));
    await store.save(sampleTokens);

    expect(writeFileSpy).toHaveBeenCalledTimes(1);
    const options = writeFileSpy.mock.calls[0]?.[2] as { flag?: string; mode?: number };
    expect(options.flag).toBe('wx');
    // 0600 still applied at creation, so the secret is never briefly readable.
    expect(options.mode).toBe(0o600);
  });

  it('uses an unpredictable temp name — two saves never reuse a path', async () => {
    const store = createTokenStore(join(dir, 'token.json'));
    await store.save(sampleTokens);
    await store.save({ ...sampleTokens, accessToken: 'access-def' });

    const first = String(writeFileSpy.mock.calls[0]?.[0]);
    const second = String(writeFileSpy.mock.calls[1]?.[0]);
    expect(first).not.toBe(second);
    // The old scheme embedded the pid and a millisecond timestamp; neither
    // should be inferable from the name any more.
    expect(first).not.toContain(String(process.pid));
    expect(first).toMatch(/\.[0-9a-f]{16}\.token\.tmp$/);
  });

  it('leaves no temp file behind on success', async () => {
    const store = createTokenStore(join(dir, 'token.json'));
    await store.save(sampleTokens);

    expect(await readdir(dir)).toEqual(['token.json']);
  });

  it('round-trips correctly through the hardened write', async () => {
    const store = createTokenStore(join(dir, 'token.json'));
    await store.save(sampleTokens);
    await expect(store.load()).resolves.toEqual(sampleTokens);
  });
});
