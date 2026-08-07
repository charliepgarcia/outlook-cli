/**
 * OAuth2 device code flow — initiate + poll. Endpoint URLs, the grant_type
 * URN, and the five polling error codes are all specified verbatim in
 * docs/00-auth-and-fundamentals.md §2-§3. Nothing here is guessed.
 *
 * No network or real timers are hardwired: both `fetchImpl` and `sleep` are
 * injectable so tests run instantly against a fake transport.
 */

import { authorityBase } from './config.js';
import { OutlookCliError, parseGraphError } from './errors.js';
import { postForm } from './identityFetch.js';
import type { DeviceCodeResponse, OutlookConfig, TokenResponse } from './types.js';

/**
 * The device-code grant_type value, copied character-for-character from
 * docs/00 §3 (verified there against RFC 8628 §3.4's own registered URN —
 * there is no shorthand form; `device_code` alone fails).
 */
export const DEVICE_CODE_GRANT_TYPE = 'urn:ietf:params:oauth:grant-type:device_code';

/** docs/00 §2: devicecode response default; not a guaranteed contract, just the observed default. */
const DEFAULT_EXPIRES_IN_SEC = 900;
/** docs/00 §2: "Observed value in practice: 5." Used only when the caller has no better value. */
const DEFAULT_POLL_INTERVAL_SEC = 5;
/** docs/00 §3 + §9.1: on `slow_down`, permanently add 5s to the interval. */
const SLOW_DOWN_INCREMENT_MS = 5_000;

export async function requestDeviceCode(
  cfg: OutlookConfig,
  scopes: readonly string[],
  fetchImpl: typeof fetch = fetch,
): Promise<DeviceCodeResponse> {
  const url = `${authorityBase(cfg.tenantId)}/devicecode`;
  const params = new URLSearchParams({
    client_id: cfg.clientId,
    scope: scopes.join(' '),
  });

  const res = await postForm(url, params, fetchImpl);
  const bodyText = await res.text().catch(() => '');

  if (!res.ok) {
    throw parseGraphError(res.status, bodyText);
  }

  const parsed = parseJsonSafe(bodyText);
  if (!isDeviceCodeResponse(parsed)) {
    throw new OutlookCliError(
      'The devicecode endpoint returned a 200 response with an unrecognized body shape.',
    );
  }
  return parsed;
}

export interface PollForTokenOptions {
  /** Seconds between polls. Default: the devicecode response's own `interval`, or 5. */
  intervalSec?: number;
  /** Hard deadline in seconds. Default: the devicecode response's own `expires_in`, or 900. */
  expiresInSec?: number;
  fetchImpl?: typeof fetch;
  /** Injectable so tests advance instantly instead of waiting on real timers. */
  sleep?: (ms: number) => Promise<void>;
}

export async function pollForToken(
  cfg: OutlookConfig,
  deviceCode: string,
  opts: PollForTokenOptions = {},
): Promise<TokenResponse> {
  const fetchImpl = opts.fetchImpl ?? fetch;
  const sleep = opts.sleep ?? defaultSleep;
  const deadlineMs = (opts.expiresInSec ?? DEFAULT_EXPIRES_IN_SEC) * 1000;
  let intervalMs = (opts.intervalSec ?? DEFAULT_POLL_INTERVAL_SEC) * 1000;

  const url = `${authorityBase(cfg.tenantId)}/token`;
  const params = new URLSearchParams({
    grant_type: DEVICE_CODE_GRANT_TYPE,
    client_id: cfg.clientId,
    device_code: deviceCode,
  });

  // Virtual clock: driven by the sleep durations we actually request, not
  // wall time. This keeps the hard-deadline check meaningful under a fake
  // `sleep` that resolves instantly in tests, while still enforcing the
  // real accounting an un-mocked run would experience.
  let elapsedMs = 0;

  for (;;) {
    const res = await postForm(url, params, fetchImpl);
    const bodyText = await res.text().catch(() => '');

    if (res.ok) {
      const parsed = parseJsonSafe(bodyText);
      if (!isTokenResponse(parsed)) {
        throw new OutlookCliError(
          'The token endpoint returned a 200 response with an unrecognized body shape.',
        );
      }
      return parsed;
    }

    const code = extractErrorCode(bodyText);
    switch (code) {
      case 'authorization_pending':
        break;
      case 'slow_down':
        intervalMs += SLOW_DOWN_INCREMENT_MS;
        break;
      case 'expired_token':
        throw new OutlookCliError(
          'The device code expired before sign-in was completed. Run login again.',
        );
      case 'authorization_declined':
        throw new OutlookCliError('Sign-in was declined.');
      case 'bad_verification_code':
        throw new OutlookCliError(
          'The device code was not recognized by the server. This should not happen in normal ' +
            'operation — it indicates a bug in the polling loop, not a transient condition.',
        );
      default:
        // Any other error (invalid_client, invalid_request, ...) is fatal
        // and carries real diagnostic detail — surface it verbatim.
        throw parseGraphError(res.status, bodyText);
    }

    if (elapsedMs + intervalMs >= deadlineMs) {
      throw new OutlookCliError(
        'The device code expired before sign-in was completed (local deadline reached). Run login again.',
      );
    }

    await sleep(intervalMs);
    elapsedMs += intervalMs;
  }
}

function parseJsonSafe(text: string): unknown {
  try {
    return JSON.parse(text);
  } catch {
    return undefined;
  }
}

/** Only the `error` field's string value is a stable contract (docs/00 §3). */
function extractErrorCode(bodyText: string): string | undefined {
  const parsed = parseJsonSafe(bodyText);
  if (parsed !== null && typeof parsed === 'object' && 'error' in parsed) {
    const value = (parsed as Record<string, unknown>)['error'];
    if (typeof value === 'string') return value;
  }
  return undefined;
}

function isDeviceCodeResponse(value: unknown): value is DeviceCodeResponse {
  if (typeof value !== 'object' || value === null) return false;
  const v = value as Record<string, unknown>;
  return (
    typeof v['device_code'] === 'string' &&
    typeof v['user_code'] === 'string' &&
    typeof v['verification_uri'] === 'string' &&
    typeof v['expires_in'] === 'number' &&
    typeof v['interval'] === 'number' &&
    (v['message'] === undefined || typeof v['message'] === 'string')
  );
}

function isTokenResponse(value: unknown): value is TokenResponse {
  if (typeof value !== 'object' || value === null) return false;
  const v = value as Record<string, unknown>;
  return (
    typeof v['access_token'] === 'string' &&
    typeof v['expires_in'] === 'number' &&
    (v['refresh_token'] === undefined || typeof v['refresh_token'] === 'string') &&
    (v['scope'] === undefined || typeof v['scope'] === 'string') &&
    (v['token_type'] === undefined || typeof v['token_type'] === 'string') &&
    (v['id_token'] === undefined || typeof v['id_token'] === 'string')
  );
}

function defaultSleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}
