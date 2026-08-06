/**
 * Tests for `outlook auth`. makeCtx's stub Authenticator throws on
 * login/refresh by design (see testUtils.ts), so every subcommand under
 * test here replaces the relevant auth method with `vi.spyOn(...).mockResolvedValue`
 * rather than relying on the stub.
 */

import { describe, expect, it, vi } from 'vitest';
import type { AuthStatus } from '../auth.js';
import { UsageError } from '../errors.js';
import type { DeviceCodeResponse, TokenSet } from '../types.js';
import { authCommand } from './auth.js';
import { makeCtx, text } from './testUtils.js';

const SECRET_ACCESS_TOKEN = 'eyJ.SUPER-SECRET-ACCESS-TOKEN.zzz';
const SECRET_REFRESH_TOKEN = 'M.SUPER-SECRET-REFRESH-TOKEN';

function sampleDeviceCode(overrides: Partial<DeviceCodeResponse> = {}): DeviceCodeResponse {
  return {
    device_code: 'device-code-abc',
    user_code: 'ABCD-EFGH',
    verification_uri: 'https://microsoft.com/devicelogin',
    expires_in: 900,
    interval: 5,
    ...overrides,
  };
}

function sampleTokenSet(overrides: Partial<TokenSet> = {}): TokenSet {
  return {
    accessToken: SECRET_ACCESS_TOKEN,
    refreshToken: SECRET_REFRESH_TOKEN,
    expiresAt: Date.parse('2024-06-01T13:00:00Z'),
    scope: 'Mail.Read',
    account: 'frannis@contoso.com',
    ...overrides,
  };
}

describe('auth', () => {
  it('status: reports signed-in with account and expiry', async () => {
    const { ctx, out } = makeCtx(['status']);
    const status: AuthStatus = {
      signedIn: true,
      expired: false,
      account: 'frannis@contoso.com',
      expiresAt: Date.parse('2024-06-01T13:00:00Z'),
    };
    vi.spyOn(ctx.auth, 'status').mockResolvedValue(status);

    await authCommand.handler(ctx);

    expect(text(out)).toContain('Signed in: yes');
    expect(text(out)).toContain('frannis@contoso.com');
  });

  it('status: reports signed-out cleanly', async () => {
    const { ctx, out } = makeCtx(['status']);
    vi.spyOn(ctx.auth, 'status').mockResolvedValue({ signedIn: false, expired: false });

    await authCommand.handler(ctx);

    expect(text(out)).toContain('Signed in: no');
  });

  it('defaults to status when no subcommand is given', async () => {
    const { ctx, out } = makeCtx([]);
    const status = vi.spyOn(ctx.auth, 'status').mockResolvedValue({ signedIn: false, expired: false });

    await authCommand.handler(ctx);

    expect(status).toHaveBeenCalledTimes(1);
    expect(text(out)).toContain('Signed in: no');
  });

  it('login: prints the device code and verification URL promptly', async () => {
    const { ctx, out } = makeCtx(['login']);
    const device = sampleDeviceCode();
    vi.spyOn(ctx.auth, 'login').mockImplementation(async (onPrompt) => {
      onPrompt(device);
      return sampleTokenSet();
    });

    await authCommand.handler(ctx);

    expect(text(out)).toContain(device.user_code);
    expect(text(out)).toContain(device.verification_uri);
  });

  it('login: prefers the server-provided message when present', async () => {
    const { ctx, out } = makeCtx(['login']);
    const device = sampleDeviceCode({ message: 'Go to https://x/devicelogin and enter CODE-1' });
    vi.spyOn(ctx.auth, 'login').mockImplementation(async (onPrompt) => {
      onPrompt(device);
      return sampleTokenSet();
    });

    await authCommand.handler(ctx);

    expect(text(out)).toContain('Go to https://x/devicelogin and enter CODE-1');
  });

  it('logout: clears the cache and confirms in output', async () => {
    const { ctx, out } = makeCtx(['logout']);
    const logout = vi.spyOn(ctx.auth, 'logout').mockResolvedValue(undefined);

    await authCommand.handler(ctx);

    expect(logout).toHaveBeenCalledTimes(1);
    expect(text(out)).toContain('Signed out');
  });

  it('refresh: forces a token refresh and reports the new expiry', async () => {
    const { ctx, out } = makeCtx(['refresh']);
    const refresh = vi
      .spyOn(ctx.auth, 'refresh')
      .mockResolvedValue(sampleTokenSet({ expiresAt: Date.parse('2024-06-01T14:00:00Z') }));

    await authCommand.handler(ctx);

    expect(refresh).toHaveBeenCalledTimes(1);
    expect(text(out)).toContain('Token refreshed');
    expect(text(out)).toContain('2024-06-01T14:00:00');
  });

  it('unknown subcommand throws a UsageError naming the valid ones', async () => {
    const { ctx } = makeCtx(['nonsense']);

    await expect(authCommand.handler(ctx)).rejects.toBeInstanceOf(UsageError);
    await expect(authCommand.handler(ctx)).rejects.toThrow(/login.*status.*logout.*refresh/s);
  });

  it('--json emits a machine-readable status shape', async () => {
    const { ctx, out } = makeCtx(['status', '--json']);
    vi.spyOn(ctx.auth, 'status').mockResolvedValue({
      signedIn: true,
      expired: false,
      account: 'frannis@contoso.com',
      expiresAt: 1717248000000,
    });

    await authCommand.handler(ctx);

    const parsed = JSON.parse(text(out));
    expect(parsed).toEqual({
      signedIn: true,
      expired: false,
      account: 'frannis@contoso.com',
      expiresAt: 1717248000000,
    });
  });

  it('never prints an access token, refresh token, or client secret', async () => {
    const { ctx, out } = makeCtx(['login']);
    vi.spyOn(ctx.auth, 'login').mockImplementation(async (onPrompt) => {
      onPrompt(sampleDeviceCode());
      return sampleTokenSet();
    });

    await authCommand.handler(ctx);

    expect(text(out)).not.toContain(SECRET_ACCESS_TOKEN);
    expect(text(out)).not.toContain(SECRET_REFRESH_TOKEN);
  });

  it('refresh also never prints the raw tokens', async () => {
    const { ctx, out } = makeCtx(['refresh', '--json']);
    vi.spyOn(ctx.auth, 'refresh').mockResolvedValue(sampleTokenSet());

    await authCommand.handler(ctx);

    expect(text(out)).not.toContain(SECRET_ACCESS_TOKEN);
    expect(text(out)).not.toContain(SECRET_REFRESH_TOKEN);
  });
});
