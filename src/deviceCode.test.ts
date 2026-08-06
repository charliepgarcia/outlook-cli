import { describe, it, expect, vi } from 'vitest';
import {
  DEVICE_CODE_GRANT_TYPE,
  requestDeviceCode,
  pollForToken,
} from './deviceCode.js';
import { OutlookCliError } from './errors.js';
import type { OutlookConfig } from './types.js';

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

const DEVICE_CODE_BODY = {
  device_code: 'DEVICECODE_OPAQUE',
  user_code: 'ABCD1234',
  verification_uri: 'https://microsoft.com/devicelogin',
  expires_in: 900,
  interval: 5,
  message: 'To sign in, use a web browser...',
};

const TOKEN_BODY = {
  token_type: 'Bearer',
  scope: 'User.Read Mail.ReadWrite Mail.Send MailboxSettings.Read',
  expires_in: 3599,
  access_token: 'ACCESS_TOKEN_VALUE',
  refresh_token: 'REFRESH_TOKEN_VALUE',
  id_token: 'HEADER.PAYLOAD.SIG',
};

function instantSleep(): (ms: number) => Promise<void> {
  return vi.fn().mockResolvedValue(undefined);
}

describe('requestDeviceCode', () => {
  it('POSTs to the exact /devicecode endpoint with the exact grant-type-free form body', async () => {
    const fetchMock = vi.fn().mockResolvedValue(jsonResponse(DEVICE_CODE_BODY));

    const result = await requestDeviceCode(cfg, ['offline_access', 'User.Read'], fetchMock);

    expect(fetchMock).toHaveBeenCalledTimes(1);
    const [url, init] = fetchMock.mock.calls[0] as [string, RequestInit];
    expect(url).toBe('https://login.microsoftonline.com/common/oauth2/v2.0/devicecode');
    expect(init.method).toBe('POST');
    expect((init.headers as Record<string, string>)['Content-Type']).toBe(
      'application/x-www-form-urlencoded',
    );
    const body = new URLSearchParams(init.body as string);
    expect(body.get('client_id')).toBe('client-abc-123');
    expect(body.get('scope')).toBe('offline_access User.Read');
    expect(result).toEqual(DEVICE_CODE_BODY);
  });

  it('throws a GraphError on a non-2xx response', async () => {
    const fetchMock = vi.fn().mockResolvedValue(
      jsonResponse({ error: 'invalid_client', error_description: 'bad app' }, 400),
    );
    await expect(
      requestDeviceCode(cfg, ['User.Read'], fetchMock),
    ).rejects.toThrow();
  });
});

describe('pollForToken', () => {
  it('uses the exact /token endpoint and the exact device-code grant_type URN', async () => {
    expect(DEVICE_CODE_GRANT_TYPE).toBe('urn:ietf:params:oauth:grant-type:device_code');

    const fetchMock = vi.fn().mockResolvedValue(jsonResponse(TOKEN_BODY));
    const sleep = instantSleep();

    await pollForToken(cfg, 'DEVICECODE_OPAQUE', {
      intervalSec: 5,
      expiresInSec: 900,
      fetchImpl: fetchMock,
      sleep,
    });

    const [url, init] = fetchMock.mock.calls[0] as [string, RequestInit];
    expect(url).toBe('https://login.microsoftonline.com/common/oauth2/v2.0/token');
    const body = new URLSearchParams(init.body as string);
    expect(body.get('grant_type')).toBe('urn:ietf:params:oauth:grant-type:device_code');
    expect(body.get('client_id')).toBe('client-abc-123');
    expect(body.get('device_code')).toBe('DEVICECODE_OPAQUE');
  });

  it('fresh login happy path: succeeds immediately with no pending polls', async () => {
    const fetchMock = vi.fn().mockResolvedValue(jsonResponse(TOKEN_BODY));
    const sleep = instantSleep();

    const result = await pollForToken(cfg, 'DEVICECODE_OPAQUE', {
      intervalSec: 5,
      expiresInSec: 900,
      fetchImpl: fetchMock,
      sleep,
    });

    expect(result).toEqual(TOKEN_BODY);
    expect(fetchMock).toHaveBeenCalledTimes(1);
    expect(sleep).not.toHaveBeenCalled();
  });

  it('authorization_pending then success: keeps polling at the same interval', async () => {
    const fetchMock = vi
      .fn()
      .mockResolvedValueOnce(jsonResponse({ error: 'authorization_pending' }, 400))
      .mockResolvedValueOnce(jsonResponse({ error: 'authorization_pending' }, 400))
      .mockResolvedValueOnce(jsonResponse(TOKEN_BODY));
    const sleep = instantSleep();

    const result = await pollForToken(cfg, 'DEVICECODE_OPAQUE', {
      intervalSec: 5,
      expiresInSec: 900,
      fetchImpl: fetchMock,
      sleep,
    });

    expect(result).toEqual(TOKEN_BODY);
    expect(fetchMock).toHaveBeenCalledTimes(3);
    expect(sleep).toHaveBeenCalledTimes(2);
    expect(sleep).toHaveBeenNthCalledWith(1, 5000);
    expect(sleep).toHaveBeenNthCalledWith(2, 5000);
  });

  it('slow_down raises the polling interval by 5s and keeps polling', async () => {
    const fetchMock = vi
      .fn()
      .mockResolvedValueOnce(jsonResponse({ error: 'slow_down' }, 400))
      .mockResolvedValueOnce(jsonResponse({ error: 'slow_down' }, 400))
      .mockResolvedValueOnce(jsonResponse(TOKEN_BODY));
    const sleep = instantSleep();

    await pollForToken(cfg, 'DEVICECODE_OPAQUE', {
      intervalSec: 5,
      expiresInSec: 900,
      fetchImpl: fetchMock,
      sleep,
    });

    // First slow_down: 5s -> 10s. Second slow_down: 10s -> 15s. Both sleeps
    // observed are the POST-increment interval, since the increment happens
    // before the wait on the same iteration.
    expect(sleep).toHaveBeenNthCalledWith(1, 10_000);
    expect(sleep).toHaveBeenNthCalledWith(2, 15_000);
  });

  it('expired_token throws a clear OutlookCliError and stops polling', async () => {
    const fetchMock = vi.fn().mockResolvedValue(jsonResponse({ error: 'expired_token' }, 400));
    const sleep = instantSleep();

    await expect(
      pollForToken(cfg, 'DEVICECODE_OPAQUE', {
        intervalSec: 5,
        expiresInSec: 900,
        fetchImpl: fetchMock,
        sleep,
      }),
    ).rejects.toThrow(OutlookCliError);
    expect(fetchMock).toHaveBeenCalledTimes(1);
  });

  it('authorization_declined throws and stops polling', async () => {
    const fetchMock = vi
      .fn()
      .mockResolvedValue(jsonResponse({ error: 'authorization_declined' }, 400));

    await expect(
      pollForToken(cfg, 'DEVICECODE_OPAQUE', {
        intervalSec: 5,
        expiresInSec: 900,
        fetchImpl: fetchMock,
        sleep: instantSleep(),
      }),
    ).rejects.toThrow(OutlookCliError);
    expect(fetchMock).toHaveBeenCalledTimes(1);
  });

  it('bad_verification_code throws and stops polling', async () => {
    const fetchMock = vi
      .fn()
      .mockResolvedValue(jsonResponse({ error: 'bad_verification_code' }, 400));

    await expect(
      pollForToken(cfg, 'DEVICECODE_OPAQUE', {
        intervalSec: 5,
        expiresInSec: 900,
        fetchImpl: fetchMock,
        sleep: instantSleep(),
      }),
    ).rejects.toThrow(OutlookCliError);
    expect(fetchMock).toHaveBeenCalledTimes(1);
  });

  it('throws once the expires_in deadline would be overrun by the next wait', async () => {
    // interval 5s, expires 8s: first poll pending, then 5+5=10s >= 8s deadline -> throw
    // before ever sleeping past the window.
    const fetchMock = vi.fn().mockResolvedValue(jsonResponse({ error: 'authorization_pending' }, 400));

    await expect(
      pollForToken(cfg, 'DEVICECODE_OPAQUE', {
        intervalSec: 5,
        expiresInSec: 8,
        fetchImpl: fetchMock,
        sleep: instantSleep(),
      }),
    ).rejects.toThrow(OutlookCliError);
  });

  it('propagates a fatal non-polling error (e.g. invalid_client) immediately', async () => {
    const fetchMock = vi
      .fn()
      .mockResolvedValue(jsonResponse({ error: 'invalid_client', error_description: 'bad app' }, 400));

    await expect(
      pollForToken(cfg, 'DEVICECODE_OPAQUE', {
        intervalSec: 5,
        expiresInSec: 900,
        fetchImpl: fetchMock,
        sleep: instantSleep(),
      }),
    ).rejects.toThrow();
    expect(fetchMock).toHaveBeenCalledTimes(1);
  });
});
