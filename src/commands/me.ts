/**
 * `outlook me` — GET /me (docs/01-read-endpoints.md §1).
 *
 * The smallest command in the CLI: fetches the signed-in user's profile and
 * renders it via renderProfile. docs/01 §1 calls `GET /me` out explicitly —
 * "This CLI should treat `GET /me` as its identity-check endpoint" — so this
 * command doubles as the natural "is my auth working?" smoke test: run it
 * with no arguments and a clean profile (or a clean error) tells you whether
 * the cached token is good.
 *
 * That means an auth failure here must never be swallowed. There is no
 * try/catch in this handler — a GraphError (e.g. InvalidAuthenticationToken)
 * or any other rejection from ctx.client.get propagates straight out to the
 * caller with the real Graph error code/message intact (see errors.ts),
 * which is exactly what makes this command useful as a smoke test.
 *
 * Read-only: `mutating: false`, no guardMutation — see command.ts.
 */

import { emit, type CommandContext, type CommandSpec } from '../command.js';
import { renderProfile } from '../output.js';
import type { UserProfile } from '../types.js';

/** Exactly the fields UserProfile (types.ts) declares — nothing more. */
const SELECT_FIELDS = 'id,displayName,mail,userPrincipalName,jobTitle';

async function meHandler(ctx: CommandContext): Promise<void> {
  const profile = await ctx.client.get<UserProfile>('/me', { $select: SELECT_FIELDS });

  emit(ctx, profile, () => renderProfile(profile));
}

export const meCommand: CommandSpec = {
  name: 'me',
  summary: "Show the signed-in user's profile — also a quick auth smoke-test",
  usage: 'outlook me [--json]',
  mutating: false,
  handler: meHandler,
};
