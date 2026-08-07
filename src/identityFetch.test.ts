/**
 * Transport hardening on the identity-platform credential exchanges
 * (review finding F-A3).
 *
 * client.ts already sets `redirect: 'error'` and a per-attempt timeout on the
 * Graph call, on the reasoning that a 3xx to another origin replays the
 * credential. The three calls that carry the MOST sensitive secrets — the
 * refresh-token exchange and both device-code calls — had neither, and a
 * 307/308 preserves the request BODY, which is exactly where a `refresh_token`
 * and a `device_code` sit. These tests pin the hardening at the shared
 * chokepoint AND at each of the three call sites, so a future call site that
 * assembles its own fetch init fails here rather than in production.
 *
 * Lives in its own file rather than in auth.test.ts (327 lines) to stay under
 * the 350-LOC ceiling.
 */
import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { postForm } from './identityFetch.js';
import { requestDeviceCode, pollForToken } from './deviceCode.js';
import { createAuthenticator } from './auth.js';
import { createTokenStore } from './tokenStore.js';
import { REQUEST_TIMEOUT_MS, type OutlookConfig, type TokenSet } from './types.js';

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

const TOKEN_BODY = {
  token_type: 'Bearer',
  scope: 'User.Read Mail.Read MailboxSettings.Read',
  expires_in: 3599,
  access_token: 'ACCESS_TOKEN_VALUE',
  refresh_token: 'REFRESH_TOKEN_VALUE',
};

const DEVICE_CODE_BODY = {
  device_code: 'DEVICECODE_OPAQUE',
  user_code: 'ABCD1234',
  verification_uri: 'https://microsoft.com/devicelogin',
  expires_in: 900,
  interval: 5,
  message: 'To sign in...',
};

/** The RequestInit the mock fetch was handed on call `n`. */
function initOf(fetchMock: ReturnType<typeof vi.fn>, call = 0): RequestInit {
  return fetchMock.mock.calls[call]?.[1] as RequestInit;
}

describe('postForm — the single hardened credential-exchange chokepoint', () => {
  it('refuses to follow a redirect, so a 3xx cannot replay the body to another origin', async () => {
    const fetchMock = vi.fn().mockResolvedValue(jsonResponse(TOKEN_BODY));
    await postForm('https://login.microsoftonline.com/common/oauth2/v2.0/token', new URLSearchParams({ a: 'b' }), fetchMock);

    expect(initOf(fetchMock).redirect).toBe('error');
  });

  it('attaches an abort signal, so a hung socket cannot stall an unattended run forever', async () => {
    const fetchMock = vi.fn().mockResolvedValue(jsonResponse(TOKEN_BODY));
    await postForm('https://login.microsoftonline.com/common/oauth2/v2.0/token', new URLSearchParams({ a: 'b' }), fetchMock);

    const signal = initOf(fetchMock).signal;
    expect(signal).toBeInstanceOf(AbortSignal);
    expect(signal?.aborted).toBe(false);
  });

  it('still POSTs a form-encoded body — the hardening did not change the wire format', async () => {
    const fetchMock = vi.fn().mockResolvedValue(jsonResponse(TOKEN_BODY));
    await postForm('https://example.invalid/token', new URLSearchParams({ grant_type: 'refresh_token' }), fetchMock);

    const init = initOf(fetchMock);
    expect(init.method).toBe('POST');
    expect(init.headers).toEqual({ 'Content-Type': 'application/x-www-form-urlencoded' });
    expect(init.body).toBe('grant_type=refresh_token');
  });

  it('uses a per-attempt timeout, not one shared deadline across calls', async () => {
    const fetchMock = vi.fn().mockResolvedValue(jsonResponse(TOKEN_BODY));
    await postForm('https://example.invalid/token', new URLSearchParams(), fetchMock);
    await postForm('https://example.invalid/token', new URLSearchParams(), fetchMock);

    // A shared deadline would hand the same object to both attempts; the
    // device-code poll depends on each request getting its own.
    expect(initOf(fetchMock, 0).signal).not.toBe(initOf(fetchMock, 1).signal);
    expect(REQUEST_TIMEOUT_MS).toBe(60_000);
  });
});

describe('every credential-carrying call site routes through postForm', () => {
  it('requestDeviceCode (device-code request) is hardened', async () => {
    const fetchMock = vi.fn().mockResolvedValue(jsonResponse(DEVICE_CODE_BODY));
    await requestDeviceCode(cfg, ['Mail.Read'], fetchMock);

    expect(initOf(fetchMock).redirect).toBe('error');
    expect(initOf(fetchMock).signal).toBeInstanceOf(AbortSignal);
  });

  it('pollForToken (posts the device_code) is hardened on every poll attempt', async () => {
    const fetchMock = vi
      .fn()
      .mockResolvedValueOnce(jsonResponse({ error: 'authorization_pending' }, 400))
      .mockResolvedValueOnce(jsonResponse(TOKEN_BODY));

    await pollForToken(cfg, 'DEVICECODE_OPAQUE', {
      intervalSec: 5,
      expiresInSec: 900,
      fetchImpl: fetchMock,
      sleep: vi.fn().mockResolvedValue(undefined),
    });

    expect(fetchMock).toHaveBeenCalledTimes(2);
    for (const call of [0, 1]) {
      expect(initOf(fetchMock, call).redirect).toBe('error');
      expect(initOf(fetchMock, call).signal).toBeInstanceOf(AbortSignal);
    }
  });

  describe('exchangeRefreshToken (posts the refresh_token)', () => {
    let dir: string;

    beforeEach(async () => {
      dir = await mkdtemp(join(tmpdir(), 'outlook-identityfetch-'));
    });

    afterEach(async () => {
      await rm(dir, { recursive: true, force: true });
    });

    it('is hardened', async () => {
      const store = createTokenStore(join(dir, 'token.json'));
      const cached: TokenSet = {
        accessToken: 'stale',
        refreshToken: 'REFRESH_TOKEN_VALUE',
        expiresAt: 0,
        mode: 'readonly',
      };
      await store.save(cached);

      const fetchMock = vi.fn().mockResolvedValue(jsonResponse(TOKEN_BODY));
      const auth = createAuthenticator(cfg, store, { fetchImpl: fetchMock, now: () => 1_000 });
      await auth.refresh();

      expect(fetchMock).toHaveBeenCalledTimes(1);
      expect(initOf(fetchMock).redirect).toBe('error');
      expect(initOf(fetchMock).signal).toBeInstanceOf(AbortSignal);
      // The refresh_token really is in the body — which is why redirect:'error'
      // matters here and header-stripping would not have helped.
      expect(String(initOf(fetchMock).body)).toContain('refresh_token=REFRESH_TOKEN_VALUE');
    });
  });
});
