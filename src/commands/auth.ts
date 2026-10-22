/**
 * `outlook auth <subcommand>` — local credential lifecycle: login, status,
 * logout, refresh. Every other command in this CLI depends on this one:
 * without `outlook auth login` there is no cached token for `ctx.auth`
 * (auth.ts) to hand out, so nothing else can call Graph.
 *
 * Subcommand is positional 0 (`status` is the default when absent, since
 * "is anything even signed in?" is the harmless, side-effect-free question
 * to default to). Unknown subcommands are a UsageError listing the valid
 * ones.
 *
 * `mutating: false` on all four subcommands: they change LOCAL credential
 * cache state (a file under tokenStore.ts), never remote mailbox state.
 * `guardMutation`/`--dry-run` exist to protect a real mailbox from a
 * surprising Graph call (send/delete/move) — a local-only side effect has
 * nothing for --dry-run to preview or protect against, so this command
 * deliberately does not wrap itself in guardMutation (see command.ts's own
 * doc comment: that chokepoint is for state-changing *Graph* calls).
 *
 * SECURITY: never print accessToken, refreshToken, or any client secret.
 * Every render path below is built from an explicit allow-list of fields
 * (signedIn, expired, account, expiresAt, scopes) — the raw TokenSet /
 * AuthStatus objects from auth.ts are never handed whole to emit()/out(),
 * specifically so a future field added to either type can't silently start
 * leaking through this command.
 */

import { emit, type CommandContext, type CommandSpec } from '../command.js';
import { UsageError } from '../errors.js';
import { FILE_MODE_IS_ENFORCED } from '../tokenStore.js';
import type { DeviceCodeResponse, GrantMode, TokenSet } from '../types.js';

const SUBCOMMANDS = ['login', 'status', 'logout', 'refresh'] as const;

/** Non-secret facts we're willing to show the user, shared by all four subcommands. */
interface AuthFacts {
  signedIn: boolean;
  expired?: boolean;
  account?: string;
  expiresAt?: number;
  scopes?: string;
  mode?: GrantMode;
}

function renderFacts(facts: AuthFacts): string {
  const lines = [`Signed in: ${facts.signedIn ? 'yes' : 'no'}`];
  if (facts.expired !== undefined) lines.push(`Expired: ${facts.expired ? 'yes' : 'no'}`);
  if (facts.account) lines.push(`Account: ${facts.account}`);
  if (facts.expiresAt !== undefined) lines.push(`Expires: ${new Date(facts.expiresAt).toISOString()}`);
  if (facts.scopes) lines.push(`Scopes: ${facts.scopes}`);
  if (facts.mode) {
    lines.push(
      `Access: ${facts.mode === 'readonly' ? 'read-only (cannot send or delete)' : 'full read/write'}`,
    );
  }
  return lines.join('\n');
}

/** Non-secret projection of a freshly issued/refreshed TokenSet — drops accessToken/refreshToken/tokenType. */
function factsFromTokenSet(tokens: TokenSet): AuthFacts {
  return {
    signedIn: true,
    ...(tokens.account !== undefined ? { account: tokens.account } : {}),
    expiresAt: tokens.expiresAt,
    ...(tokens.scope !== undefined ? { scopes: tokens.scope } : {}),
    ...(tokens.mode !== undefined ? { mode: tokens.mode } : {}),
  };
}

/**
 * The device-code prompt IS the login UX — printed the moment auth.login()
 * hands it back, before the (potentially slow) polling loop even starts.
 * Prefers the server's own pre-formatted `message` (it already names both
 * the URL and the code); falls back to assembling one from the two fields
 * that actually matter if the server omits it.
 */
function printDevicePrompt(ctx: CommandContext, device: DeviceCodeResponse): void {
  ctx.out(
    device.message ??
      `To sign in, open ${device.verification_uri} and enter the code: ${device.user_code}`,
  );
}

async function loginHandler(ctx: CommandContext): Promise<void> {
  const readOnly = ctx.args.bool('readonly');
  const tokens = await ctx.auth.login(
    (device) => printDevicePrompt(ctx, device),
    { readOnly },
  );

  // The stored refresh token is only as private as the file holding it, and
  // on Windows the 0600 we request is not actually enforced. Say so once, at
  // the moment the secret is created, rather than implying a guarantee.
  if (!FILE_MODE_IS_ENFORCED) {
    ctx.err(
      'warning: on Windows the token cache cannot be restricted to your user ' +
        'account by file mode alone. Restrict the containing directory with an ' +
        'ACL, or sign in with --readonly to limit what a stolen token can do.',
    );
  }

  const facts = factsFromTokenSet(tokens);
  emit(ctx, facts, () => `Signed in.\n${renderFacts(facts)}`);
}

async function statusHandler(ctx: CommandContext): Promise<void> {
  const status = await ctx.auth.status();
  emit(ctx, status, () => renderFacts(status));
}

async function logoutHandler(ctx: CommandContext): Promise<void> {
  await ctx.auth.logout();
  emit(ctx, { signedIn: false }, () => 'Signed out.');
}

async function refreshHandler(ctx: CommandContext): Promise<void> {
  const tokens = await ctx.auth.refresh();
  const facts = factsFromTokenSet(tokens);
  emit(ctx, facts, () => `Token refreshed.\n${renderFacts(facts)}`);
}

async function authHandler(ctx: CommandContext): Promise<void> {
  // Note: ParsedArgs.str(name, fallback) keys off a *named* flag (--name),
  // not a positional index — positionals are read directly off the array.
  const subcommand = ctx.args.positionals[0] ?? 'status';
  switch (subcommand) {
    case 'login':
      return loginHandler(ctx);
    case 'status':
      return statusHandler(ctx);
    case 'logout':
      return logoutHandler(ctx);
    case 'refresh':
      return refreshHandler(ctx);
    default:
      throw new UsageError(
        `Unknown "outlook auth" subcommand: "${subcommand}". Valid subcommands: ${SUBCOMMANDS.join(', ')}.`,
      );
  }
}

export const authCommand: CommandSpec = {
  name: 'auth',
  summary: 'Manage local sign-in: login, status, logout, refresh',
  usage:
    'outlook auth [login|status|logout|refresh] [--readonly] [--json]\n' +
    '\n' +
    '  --readonly  On login: request Mail.Read instead of Mail.ReadWrite and\n' +
    '              drop Mail.Send, so the stored token cannot send, delete, or\n' +
    '              modify mail. Recommended for unattended/agent use.',
  mutating: false,
  handler: authHandler,
};
