/**
 * The contract every command binds to.
 *
 * The --dry-run short-circuit lives HERE, not in individual commands, so that
 * no command can forget it. Sending mail is irreversible and reaches real
 * people; `guardMutation` is the single chokepoint that makes the safe path
 * the default path.
 */

import type { ParsedArgs } from './args.js';
import type { Authenticator } from './auth.js';
import type { GraphClient } from './client.js';
import { renderRequest } from './dryRun.js';
import type { RequestSpec } from './types.js';

export interface CommandContext {
  args: ParsedArgs;
  /** --json: emit machine-readable output. */
  json: boolean;
  /** --dry-run: render the request, send nothing. */
  dryRun: boolean;
  client: GraphClient;
  auth: Authenticator;
  out: (line: string) => void;
  err: (line: string) => void;
}

export type CommandHandler = (ctx: CommandContext) => Promise<void>;

export interface CommandSpec {
  name: string;
  summary: string;
  usage: string;
  /** True when the command changes remote state — used to document --dry-run. */
  mutating: boolean;
  handler: CommandHandler;
}

/**
 * Wrap every state-changing Graph call. When --dry-run is set, prints the
 * exact request that WOULD have been issued and returns undefined without
 * touching the network.
 */
export async function guardMutation<T>(
  ctx: CommandContext,
  spec: RequestSpec,
  exec: () => Promise<T>,
): Promise<T | undefined> {
  if (ctx.dryRun) {
    if (ctx.json) {
      ctx.out(
        JSON.stringify({ dryRun: true, request: redactSpecForJson(spec) }, null, 2),
      );
    } else {
      ctx.out('DRY RUN — nothing was sent. The request below is what would go out:');
      ctx.out('');
      ctx.out(renderRequest(spec));
    }
    return undefined;
  }
  return exec();
}

function redactSpecForJson(spec: RequestSpec): RequestSpec {
  const { method, url, body } = spec;
  return { method, url, ...(body === undefined ? {} : { body }) };
}

/** Emit either JSON or a human rendering, chosen by --json. */
export function emit(
  ctx: CommandContext,
  data: unknown,
  human: () => string,
): void {
  ctx.out(ctx.json ? JSON.stringify(data, null, 2) : human());
}
