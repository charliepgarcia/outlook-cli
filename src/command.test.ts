/**
 * Tests for guardMutation and emit — the single chokepoint every mutating
 * command routes through for the --dry-run short-circuit, the readonly-scope
 * gate, and the --json/human output split.
 */

import { describe, expect, it, vi } from 'vitest';
import { makeCtx, text } from './commands/testUtils.js';
import { emit, guardMutation } from './command.js';
import { InsufficientScopeError } from './errors.js';
import { MASK } from './redact.js';
import type { RequestSpec } from './types.js';

const SEND_SPEC: RequestSpec = {
  method: 'POST',
  url: 'https://graph.microsoft.com/v1.0/me/sendMail',
  body: { message: { subject: 'hi' }, saveToSentItems: true },
};

const RAW_SECRET = 'SUPER-SECRET-ACCESS-TOKEN-VALUE-000111222';

/** A RequestSpec shaped like a token exchange body — where access_token lives. */
const SECRET_SPEC: RequestSpec = {
  method: 'POST',
  url: 'https://login.microsoftonline.com/common/oauth2/v2.0/token',
  body: { access_token: RAW_SECRET, token_type: 'Bearer', scope: 'Mail.Send' },
};

describe('guardMutation — dry run', () => {
  it('under dryRun:true does not invoke exec and prints the request', async () => {
    const { ctx, out } = makeCtx([]);
    ctx.dryRun = true;
    const exec = vi.fn(async () => ({ id: 'sent-1' }));

    const result = await guardMutation(ctx, SEND_SPEC, exec);

    expect(exec).not.toHaveBeenCalled();
    expect(result).toBeUndefined();
    expect(text(out)).toContain('DRY RUN');
    expect(text(out)).toContain(SEND_SPEC.url);
  });

  it('JSON dry-run redacts a secret-keyed body field — raw secret never appears', async () => {
    const { ctx, out } = makeCtx([]);
    ctx.dryRun = true;
    ctx.json = true;
    const exec = vi.fn(async () => ({ id: 'sent-1' }));

    await guardMutation(ctx, SECRET_SPEC, exec);

    expect(exec).not.toHaveBeenCalled();

    const rendered = text(out);
    expect(rendered).not.toContain(RAW_SECRET);
    expect(rendered).toContain(MASK);

    const parsed = JSON.parse(out[0] ?? '') as {
      dryRun: boolean;
      request: { body?: { access_token?: string } };
    };
    expect(parsed.dryRun).toBe(true);
    expect(parsed.request.body?.access_token).toContain(MASK);
    expect(parsed.request.body?.access_token).not.toContain(RAW_SECRET);
  });

  it('human dry-run also redacts the same secret-keyed body field (regression guard)', async () => {
    const { ctx, out } = makeCtx([]);
    ctx.dryRun = true;
    ctx.json = false;
    const exec = vi.fn(async () => ({ id: 'sent-1' }));

    await guardMutation(ctx, SECRET_SPEC, exec);

    expect(exec).not.toHaveBeenCalled();
    const rendered = text(out);
    expect(rendered).not.toContain(RAW_SECRET);
    expect(rendered).toContain(MASK);
  });
});

describe('guardMutation — readonly-scope gate', () => {
  it('rejects with InsufficientScopeError and never calls exec when signed in readonly', async () => {
    const { ctx } = makeCtx([]);
    ctx.dryRun = false;
    vi.spyOn(ctx.auth, 'status').mockResolvedValue({
      signedIn: true,
      expired: false,
      mode: 'readonly',
    });
    const exec = vi.fn(async () => ({ id: 'sent-1' }));

    await expect(guardMutation(ctx, SEND_SPEC, exec)).rejects.toBeInstanceOf(
      InsufficientScopeError,
    );
    expect(exec).not.toHaveBeenCalled();
  });

  it('does not fire when mode is "full"', async () => {
    const { ctx } = makeCtx([]);
    ctx.dryRun = false;
    vi.spyOn(ctx.auth, 'status').mockResolvedValue({
      signedIn: true,
      expired: false,
      mode: 'full',
    });
    const exec = vi.fn(async () => ({ id: 'sent-1' }));

    const result = await guardMutation(ctx, SEND_SPEC, exec);

    expect(exec).toHaveBeenCalledTimes(1);
    expect(result).toEqual({ id: 'sent-1' });
  });

  it('does not fire when mode is undefined (legacy pre-readonly token)', async () => {
    const { ctx } = makeCtx([]);
    ctx.dryRun = false;
    vi.spyOn(ctx.auth, 'status').mockResolvedValue({
      signedIn: true,
      expired: false,
    });
    const exec = vi.fn(async () => ({ id: 'sent-1' }));

    const result = await guardMutation(ctx, SEND_SPEC, exec);

    expect(exec).toHaveBeenCalledTimes(1);
    expect(result).toEqual({ id: 'sent-1' });
  });

  it('does not fire when signedIn is false — leaves NotAuthenticatedError to the downstream token fetch', async () => {
    const { ctx } = makeCtx([]);
    ctx.dryRun = false;
    vi.spyOn(ctx.auth, 'status').mockResolvedValue({
      signedIn: false,
      expired: false,
    });
    const exec = vi.fn(async () => ({ id: 'sent-1' }));

    const result = await guardMutation(ctx, SEND_SPEC, exec);

    expect(exec).toHaveBeenCalledTimes(1);
    expect(result).toEqual({ id: 'sent-1' });
  });
});

describe('emit', () => {
  it('emits JSON when json:true', () => {
    const { ctx, out } = makeCtx([]);
    ctx.json = true;
    const human = vi.fn(() => 'human string');

    emit(ctx, { foo: 'bar' }, human);

    expect(human).not.toHaveBeenCalled();
    expect(text(out)).toBe(JSON.stringify({ foo: 'bar' }, null, 2));
  });

  it('emits the human string when json:false', () => {
    const { ctx, out } = makeCtx([]);
    ctx.json = false;
    const human = vi.fn(() => 'human string');

    emit(ctx, { foo: 'bar' }, human);

    expect(human).toHaveBeenCalledTimes(1);
    expect(text(out)).toBe('human string');
  });
});
