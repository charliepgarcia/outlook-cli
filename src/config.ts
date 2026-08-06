/**
 * Configuration resolution. Env is injected (never read from the global
 * `process` inside library code) so tests can drive it directly.
 */

import { homedir } from 'node:os';
import { join } from 'node:path';
import { MissingCredentialsError } from './errors.js';
import type { OutlookConfig } from './types.js';

export type ConfigNeed = 'clientId' | 'clientSecret';

export interface EnvLike {
  [key: string]: string | undefined;
}

/** Default token cache: ~/.afk/state/outlook/token.json */
export function defaultTokenCachePath(env: EnvLike = {}): string {
  const override = trimmed(env['GRAPH_TOKEN_CACHE']);
  if (override) return override;
  const afkHome = trimmed(env['AFK_HOME']) ?? join(homedir(), '.afk');
  return join(afkHome, 'state', 'outlook', 'token.json');
}

/**
 * Resolve config from env, failing loud with the exact missing var names.
 * `needs` lets read-only paths (e.g. `auth status`) avoid demanding a secret.
 */
export function resolveConfig(
  env: EnvLike,
  needs: readonly ConfigNeed[] = ['clientId'],
): OutlookConfig {
  const missing: string[] = [];

  const clientId = trimmed(env['GRAPH_CLIENT_ID']);
  if (needs.includes('clientId') && !clientId) {
    missing.push('GRAPH_CLIENT_ID');
  }

  const clientSecret = trimmed(env['GRAPH_CLIENT_SECRET']);
  if (needs.includes('clientSecret') && !clientSecret) {
    missing.push('GRAPH_CLIENT_SECRET');
  }

  if (missing.length > 0) {
    throw new MissingCredentialsError(missing);
  }

  return {
    clientId: clientId ?? '',
    tenantId: trimmed(env['GRAPH_TENANT_ID']) ?? 'common',
    ...(clientSecret ? { clientSecret } : {}),
    tokenCachePath: defaultTokenCachePath(env),
  };
}

/** Identity-platform base for this tenant. */
export function authorityBase(tenantId: string): string {
  return `https://login.microsoftonline.com/${encodeURIComponent(tenantId)}/oauth2/v2.0`;
}

function trimmed(value: string | undefined): string | undefined {
  if (typeof value !== 'string') return undefined;
  const t = value.trim();
  return t.length > 0 ? t : undefined;
}
