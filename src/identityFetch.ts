/**
 * The single place a credential is POSTed to the identity platform.
 *
 * auth.ts (the refresh-token exchange) and deviceCode.ts (the device-code
 * request and the token poll) all send a secret in the request BODY — a
 * `refresh_token` or a `device_code`. client.ts already refuses to follow a
 * redirect for exactly this reason, but its guard covers the Graph host only,
 * and these three calls never touch it.
 *
 * A 307/308 preserves the method AND the body, so a redirected credential
 * exchange replays the secret to whichever origin answered. The cross-origin
 * `Authorization`-stripping that fetch does for free is no help here: the
 * secret is not in a header. Hence `redirect: 'error'` — the identity platform
 * does not redirect these endpoints, so a 3xx is hostile, not routine.
 *
 * Routing all three through one function means a fourth call site cannot
 * reintroduce the hole by assembling its own fetch init. Same reasoning that
 * puts the bearer-host check inside client.ts's `send()` rather than in
 * `buildUrl`, and `--dry-run` inside `guardMutation` rather than in each
 * command.
 */

import { REQUEST_TIMEOUT_MS } from './types.js';

const FORM_HEADERS = { 'Content-Type': 'application/x-www-form-urlencoded' };

export function postForm(
  url: string,
  params: URLSearchParams,
  fetchImpl: typeof fetch,
): Promise<Response> {
  return fetchImpl(url, {
    method: 'POST',
    headers: FORM_HEADERS,
    body: params.toString(),
    redirect: 'error',
    // Per ATTEMPT, deliberately not per operation: the device-code poll
    // legitimately runs for the lifetime of the code, so each request inside
    // that loop gets a fresh timeout rather than sharing one deadline.
    signal: AbortSignal.timeout(REQUEST_TIMEOUT_MS),
  });
}
