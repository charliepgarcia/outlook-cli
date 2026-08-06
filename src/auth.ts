/**
 * Token acquisition, refresh, and status — the surface the CLI's commands
 * actually call. Wraps deviceCode.ts (initiate + poll) and tokenStore.ts
 * (disk cache) behind a single `Authenticator`.
 *
 * Refresh-token rotation (docs/00 §4) is load-bearing: Microsoft may return
 * a NEW refresh token on every refresh call, and the old one can be silently
 * invalidated once that happens. Every refresh here persists whatever
 * refresh_token comes back — never the one that was just spent.
 */

import { authorityBase } from './config.js';
import { NotAuthenticatedError, OutlookCliError, parseGraphError } from './errors.js';
import { requestDeviceCode, pollForToken } from './deviceCode.js';
import type { TokenStore } from './tokenStore.js';
import type { DeviceCodeResponse, OutlookConfig, TokenResponse, TokenSet } from './types.js';
import { DEFAULT_SCOPES } from './types.js';

/** getAccessToken refreshes proactively once the token is this close to expiry. */
const REFRESH_SKEW_MS = 60_000;

const FORM_HEADERS = { 'Content-Type': 'application/x-www-form-urlencoded' };

export interface AuthStatus {
  signedIn: boolean;
  account?: string;
  expiresAt?: number;
  expired: boolean;
  scopes?: string;
}

export interface Authenticator {
  getAccessToken(): Promise<string>;
  login(onPrompt: (d: DeviceCodeResponse) => void): Promise<TokenSet>;
  logout(): Promise<void>;
  refresh(): Promise<TokenSet>;
  status(): Promise<AuthStatus>;
}

export interface AuthenticatorDeps {
  fetchImpl?: typeof fetch;
  sleep?: (ms: number) => Promise<void>;
  /** Injectable clock so expiry math is deterministic in tests. */
  now?: () => number;
}

export function createAuthenticator(
  cfg: OutlookConfig,
  store: TokenStore,
  deps: AuthenticatorDeps = {},
): Authenticator {
  const fetchImpl = deps.fetchImpl ?? fetch;
  const sleep = deps.sleep ?? defaultSleep;
  const now = deps.now ?? Date.now;

  async function refresh(): Promise<TokenSet> {
    const cached = await store.load();
    if (!cached) throw new NotAuthenticatedError();
    if (!cached.refreshToken) {
      throw new NotAuthenticatedError('cached session has no refresh token; run login again');
    }

    const tokenResponse = await exchangeRefreshToken(cfg, cached.refreshToken, fetchImpl);
    const nextTokens = toTokenSet(tokenResponse, now(), cached);
    await store.save(nextTokens);
    return nextTokens;
  }

  return {
    async getAccessToken(): Promise<string> {
      const cached = await store.load();
      if (!cached) throw new NotAuthenticatedError();

      if (cached.expiresAt - now() < REFRESH_SKEW_MS) {
        const refreshed = await refresh();
        return refreshed.accessToken;
      }
      return cached.accessToken;
    },

    async login(onPrompt: (d: DeviceCodeResponse) => void): Promise<TokenSet> {
      const deviceCode = await requestDeviceCode(cfg, DEFAULT_SCOPES, fetchImpl);
      onPrompt(deviceCode);

      const tokenResponse = await pollForToken(cfg, deviceCode.device_code, {
        intervalSec: deviceCode.interval,
        expiresInSec: deviceCode.expires_in,
        fetchImpl,
        sleep,
      });

      const tokens = toTokenSet(tokenResponse, now());
      await store.save(tokens);
      return tokens;
    },

    async logout(): Promise<void> {
      await store.clear();
    },

    refresh,

    async status(): Promise<AuthStatus> {
      const cached = await store.load();
      if (!cached) return { signedIn: false, expired: false };

      return {
        signedIn: true,
        expired: now() >= cached.expiresAt,
        expiresAt: cached.expiresAt,
        ...(cached.account !== undefined ? { account: cached.account } : {}),
        ...(cached.scope !== undefined ? { scopes: cached.scope } : {}),
      };
    },
  };
}

/** docs/00 §4: same /token endpoint, grant_type=refresh_token, no client_secret for a public client. */
async function exchangeRefreshToken(
  cfg: OutlookConfig,
  refreshToken: string,
  fetchImpl: typeof fetch,
): Promise<TokenResponse> {
  const url = `${authorityBase(cfg.tenantId)}/token`;
  const params = new URLSearchParams({
    grant_type: 'refresh_token',
    client_id: cfg.clientId,
    refresh_token: refreshToken,
    scope: DEFAULT_SCOPES.join(' '),
  });

  const res = await fetchImpl(url, {
    method: 'POST',
    headers: FORM_HEADERS,
    body: params.toString(),
  });
  const bodyText = await res.text().catch(() => '');

  if (!res.ok) {
    throw parseGraphError(res.status, bodyText);
  }

  let parsed: unknown;
  try {
    parsed = JSON.parse(bodyText);
  } catch {
    throw new OutlookCliError(
      'The token endpoint returned a 200 response with an unparseable body during refresh.',
    );
  }
  if (!isTokenResponse(parsed)) {
    throw new OutlookCliError(
      'The token endpoint returned a 200 response with an unrecognized body shape during refresh.',
    );
  }
  return parsed;
}

/**
 * Builds the TokenSet to persist. Rotation contract: `refresh_token` from
 * THIS response wins whenever present; we fall back to `previous`'s only
 * when the server omitted one on this call (still oscillates toward the
 * newest token Microsoft has issued, never re-sends a stale one).
 */
function toTokenSet(
  tokenResponse: TokenResponse,
  nowMs: number,
  previous?: TokenSet,
): TokenSet {
  const account = tokenResponse.id_token
    ? (decodeIdTokenAccount(tokenResponse.id_token) ?? previous?.account)
    : previous?.account;
  const refreshToken = tokenResponse.refresh_token ?? previous?.refreshToken;
  const scope = tokenResponse.scope ?? previous?.scope;
  const tokenType = tokenResponse.token_type ?? previous?.tokenType;

  return {
    accessToken: tokenResponse.access_token,
    expiresAt: nowMs + tokenResponse.expires_in * 1000,
    ...(refreshToken !== undefined ? { refreshToken } : {}),
    ...(scope !== undefined ? { scope } : {}),
    ...(tokenType !== undefined ? { tokenType } : {}),
    ...(account !== undefined ? { account } : {}),
  };
}

/**
 * Best-effort account resolution from an id_token's payload claims.
 * Signature is NEVER verified here — this is display metadata only, not an
 * authorization decision, and the token itself is discarded after this call
 * (TokenSet has no id_token field; nothing derived from it is ever logged).
 */
function decodeIdTokenAccount(idToken: string): string | undefined {
  try {
    const parts = idToken.split('.');
    if (parts.length !== 3) return undefined;
    const payloadSegment = parts[1];
    if (!payloadSegment) return undefined;

    const json = Buffer.from(payloadSegment, 'base64url').toString('utf8');
    const claims: unknown = JSON.parse(json);
    if (typeof claims !== 'object' || claims === null) return undefined;

    const c = claims as Record<string, unknown>;
    const candidates = [c['preferred_username'], c['email'], c['upn'], c['name']];
    for (const candidate of candidates) {
      if (typeof candidate === 'string' && candidate.length > 0) return candidate;
    }
    return undefined;
  } catch {
    return undefined;
  }
}

function isTokenResponse(value: unknown): value is TokenResponse {
  if (typeof value !== 'object' || value === null) return false;
  const v = value as Record<string, unknown>;
  return typeof v['access_token'] === 'string' && typeof v['expires_in'] === 'number';
}

function defaultSleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}
