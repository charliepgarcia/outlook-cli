import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { createAuthenticator } from './auth.js';
import { createTokenStore } from './tokenStore.js';
import { NotAuthenticatedError } from './errors.js';
import type { OutlookConfig, TokenSet } from './types.js';

const cfg: OutlookConfig = {
  clientId: 'client-abc-123',
  tenantId: 'common',
  tokenCachePath: '/dev/null',
};

function jsonResponse(body: unknown, status = 200): Response {
  return new Response(JSON.stringify(body), {
    status,
    headers: { 'Content-Type': 'application/json' },
  });
}

/** Encodes a fake (unsigned, header/payload/sig) id_token carrying the given claims. */
function fakeIdToken(claims: Record<string, unknown>): string {
  const header = Buffer.from(JSON.stringify({ alg: 'none' })).toString('base64url');
  const payload = Buffer.from(JSON.stringify(claims)).toString('base64url');
  return `${header}.${payload}.`;
}

const TOKEN_BODY = {
  token_type: 'Bearer',
  scope: 'User.Read Mail.ReadWrite Mail.Send MailboxSettings.Read',
  expires_in: 3599,
  access_token: 'ACCESS_TOKEN_VALUE',
  refresh_token: 'REFRESH_TOKEN_VALUE',
  id_token: fakeIdToken({ preferred_username: 'user@example.com' }),
};

const DEVICE_CODE_BODY = {
  device_code: 'DEVICECODE_OPAQUE',
  user_code: 'ABCD1234',
  verification_uri: 'https://microsoft.com/devicelogin',
  expires_in: 900,
  interval: 5,
  message: 'To sign in...',
};

function instantSleep(): (ms: number) => Promise<void> {
  return vi.fn().mockResolvedValue(undefined);
}

/** Routes /devicecode vs /token so login() can be driven end-to-end against one mock. */
function loginFlowFetch(tokenBody: unknown = TOKEN_BODY, tokenStatus = 200) {
  return vi.fn().mockImplementation(async (url: string | URL) => {
    const u = String(url);
    if (u.includes('/devicecode')) return jsonResponse(DEVICE_CODE_BODY);
    if (u.includes('/token')) return jsonResponse(tokenBody, tokenStatus);
    throw new Error(`unexpected url in test: ${u}`);
  });
}

describe('createAuthenticator', () => {
  let dir: string;
  let tokenPath: string;

  beforeEach(async () => {
    dir = await mkdtemp(join(tmpdir(), 'outlook-cli-auth-test-'));
    tokenPath = join(dir, 'token.json');
  });

  afterEach(async () => {
    await rm(dir, { recursive: true, force: true });
  });

  it('login(): fresh happy path persists tokens and resolves the account from id_token', async () => {
    const store = createTokenStore(tokenPath);
    const fetchMock = loginFlowFetch();
    const now = () => 1_000_000;
    const authenticator = createAuthenticator(cfg, store, {
      fetchImpl: fetchMock,
      sleep: instantSleep(),
      now,
    });

    const onPrompt = vi.fn();
    const tokens = await authenticator.login(onPrompt);

    expect(onPrompt).toHaveBeenCalledWith(DEVICE_CODE_BODY);
    expect(tokens.accessToken).toBe('ACCESS_TOKEN_VALUE');
    expect(tokens.refreshToken).toBe('REFRESH_TOKEN_VALUE');
    expect(tokens.expiresAt).toBe(now() + 3599 * 1000);
    expect(tokens.account).toBe('user@example.com');

    const persisted = await store.load();
    expect(persisted).toEqual(tokens);
  });

  it('login(): authorization_pending then success still lands a signed-in session', async () => {
    const store = createTokenStore(tokenPath);
    let tokenCalls = 0;
    const fetchMock = vi.fn().mockImplementation(async (url: string | URL) => {
      const u = String(url);
      if (u.includes('/devicecode')) return jsonResponse(DEVICE_CODE_BODY);
      tokenCalls += 1;
      if (tokenCalls === 1) return jsonResponse({ error: 'authorization_pending' }, 400);
      return jsonResponse(TOKEN_BODY);
    });
    const authenticator = createAuthenticator(cfg, store, {
      fetchImpl: fetchMock,
      sleep: instantSleep(),
      now: () => 1_000_000,
    });

    const tokens = await authenticator.login(() => {});
    expect(tokens.accessToken).toBe('ACCESS_TOKEN_VALUE');
    expect(tokenCalls).toBe(2);
  });

  it('getAccessToken(): throws NotAuthenticatedError when no cache exists', async () => {
    const store = createTokenStore(tokenPath);
    const authenticator = createAuthenticator(cfg, store, {
      fetchImpl: vi.fn(),
      sleep: instantSleep(),
    });

    await expect(authenticator.getAccessToken()).rejects.toThrow(NotAuthenticatedError);
  });

  it('getAccessToken(): returns the cached access token when still fresh', async () => {
    const store = createTokenStore(tokenPath);
    const fresh: TokenSet = {
      accessToken: 'still-fresh',
      refreshToken: 'refresh-1',
      expiresAt: 1_000_000 + 10 * 60_000, // 10 minutes out
    };
    await store.save(fresh);

    const fetchMock = vi.fn();
    const authenticator = createAuthenticator(cfg, store, {
      fetchImpl: fetchMock,
      sleep: instantSleep(),
      now: () => 1_000_000,
    });

    await expect(authenticator.getAccessToken()).resolves.toBe('still-fresh');
    expect(fetchMock).not.toHaveBeenCalled();
  });

  it('getAccessToken(): refreshes when within 60s of expiry, and persists the new token', async () => {
    const store = createTokenStore(tokenPath);
    const almostExpired: TokenSet = {
      accessToken: 'about-to-expire',
      refreshToken: 'refresh-1',
      expiresAt: 1_000_000 + 30_000, // 30s out — inside the 60s skew window
    };
    await store.save(almostExpired);

    const fetchMock = vi.fn().mockResolvedValue(
      jsonResponse({
        token_type: 'Bearer',
        expires_in: 3600,
        access_token: 'refreshed-access-token',
        refresh_token: 'refreshed-refresh-token',
        scope: 'User.Read',
      }),
    );
    const authenticator = createAuthenticator(cfg, store, {
      fetchImpl: fetchMock,
      sleep: instantSleep(),
      now: () => 1_000_000,
    });

    const accessToken = await authenticator.getAccessToken();
    expect(accessToken).toBe('refreshed-access-token');

    const persisted = await store.load();
    expect(persisted?.accessToken).toBe('refreshed-access-token');
    expect(persisted?.refreshToken).toBe('refreshed-refresh-token');
    expect(persisted?.expiresAt).toBe(1_000_000 + 3600 * 1000);
  });

  it('refresh(): persists a NEW rotated refresh token, never the one just spent', async () => {
    const store = createTokenStore(tokenPath);
    const original: TokenSet = {
      accessToken: 'old-access',
      refreshToken: 'ORIGINAL_REFRESH',
      expiresAt: 1_000_000,
    };
    await store.save(original);

    const fetchMock = vi.fn().mockImplementation(async (_url: string | URL, init?: RequestInit) => {
      const body = new URLSearchParams(init?.body as string);
      expect(body.get('grant_type')).toBe('refresh_token');
      expect(body.get('refresh_token')).toBe('ORIGINAL_REFRESH');
      return jsonResponse({
        token_type: 'Bearer',
        expires_in: 3600,
        access_token: 'new-access',
        refresh_token: 'ROTATED_REFRESH',
      });
    });
    const authenticator = createAuthenticator(cfg, store, {
      fetchImpl: fetchMock,
      sleep: instantSleep(),
      now: () => 2_000_000,
    });

    const result = await authenticator.refresh();
    expect(result.refreshToken).toBe('ROTATED_REFRESH');

    const persisted = await store.load();
    expect(persisted?.refreshToken).toBe('ROTATED_REFRESH');
    expect(persisted?.refreshToken).not.toBe('ORIGINAL_REFRESH');
  });

  it('refresh(): if the server omits refresh_token on this call, keeps the previous one', async () => {
    const store = createTokenStore(tokenPath);
    await store.save({
      accessToken: 'old-access',
      refreshToken: 'KEEP_ME',
      expiresAt: 1_000_000,
    });

    const fetchMock = vi.fn().mockResolvedValue(
      jsonResponse({
        token_type: 'Bearer',
        expires_in: 3600,
        access_token: 'new-access',
        // no refresh_token in this response
      }),
    );
    const authenticator = createAuthenticator(cfg, store, {
      fetchImpl: fetchMock,
      sleep: instantSleep(),
      now: () => 2_000_000,
    });

    const result = await authenticator.refresh();
    expect(result.refreshToken).toBe('KEEP_ME');
  });

  it('refresh(): throws NotAuthenticatedError when there is no cache at all', async () => {
    const store = createTokenStore(tokenPath);
    const authenticator = createAuthenticator(cfg, store, {
      fetchImpl: vi.fn(),
      sleep: instantSleep(),
    });
    await expect(authenticator.refresh()).rejects.toThrow(NotAuthenticatedError);
  });

  it('refresh(): throws NotAuthenticatedError when the cache has no refresh token', async () => {
    const store = createTokenStore(tokenPath);
    await store.save({ accessToken: 'a', expiresAt: 1_000_000 });
    const authenticator = createAuthenticator(cfg, store, {
      fetchImpl: vi.fn(),
      sleep: instantSleep(),
    });
    await expect(authenticator.refresh()).rejects.toThrow(NotAuthenticatedError);
  });

  it('logout(): clears the cache so a subsequent status() reports signed out', async () => {
    const store = createTokenStore(tokenPath);
    await store.save({ accessToken: 'a', refreshToken: 'r', expiresAt: 9_999_999_999_999 });
    const authenticator = createAuthenticator(cfg, store, {
      fetchImpl: vi.fn(),
      sleep: instantSleep(),
    });

    await authenticator.logout();
    await expect(authenticator.status()).resolves.toEqual({ signedIn: false, expired: false });
  });

  it('status(): reports signed out with no cache', async () => {
    const store = createTokenStore(tokenPath);
    const authenticator = createAuthenticator(cfg, store, { fetchImpl: vi.fn() });
    await expect(authenticator.status()).resolves.toEqual({ signedIn: false, expired: false });
  });

  it('status(): reports signed in, not expired, and includes account/scopes but NEVER a raw token', async () => {
    const store = createTokenStore(tokenPath);
    await store.save({
      accessToken: 'SUPER_SECRET_ACCESS_TOKEN',
      refreshToken: 'SUPER_SECRET_REFRESH_TOKEN',
      expiresAt: 2_000_000,
      account: 'user@example.com',
      scope: 'User.Read Mail.Send',
    });
    const authenticator = createAuthenticator(cfg, store, {
      fetchImpl: vi.fn(),
      now: () => 1_000_000,
    });

    const status = await authenticator.status();
    expect(status).toEqual({
      signedIn: true,
      expired: false,
      expiresAt: 2_000_000,
      account: 'user@example.com',
      scopes: 'User.Read Mail.Send',
    });

    const serialized = JSON.stringify(status);
    expect(serialized).not.toContain('SUPER_SECRET_ACCESS_TOKEN');
    expect(serialized).not.toContain('SUPER_SECRET_REFRESH_TOKEN');
  });

  it('status(): reports expired: true once past expiresAt', async () => {
    const store = createTokenStore(tokenPath);
    await store.save({ accessToken: 'a', refreshToken: 'r', expiresAt: 1_000_000 });
    const authenticator = createAuthenticator(cfg, store, {
      fetchImpl: vi.fn(),
      now: () => 1_000_001,
    });

    await expect(authenticator.status()).resolves.toMatchObject({ signedIn: true, expired: true });
  });

  it('status(): a corrupt cache file degrades to signed-out rather than throwing', async () => {
    const store = createTokenStore(tokenPath);
    // Simulate corruption by writing invalid JSON straight through fs, bypassing save().
    const { writeFile } = await import('node:fs/promises');
    await writeFile(tokenPath, 'not json at all', { mode: 0o600 });

    const authenticator = createAuthenticator(cfg, store, { fetchImpl: vi.fn() });
    await expect(authenticator.status()).resolves.toEqual({ signedIn: false, expired: false });
  });
});
