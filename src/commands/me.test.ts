/**
 * Tests for `outlook me`. Mirrors get.test.ts's shape: a real GraphClient
 * from makeCtx (so buildSpec/URL-joining behave exactly as in production),
 * with the network boundary (`client.get`) stubbed via vi.spyOn. makeCtx's
 * injected fetch THROWS on purpose, so any leak past the stub fails loudly.
 */

import { describe, expect, it, vi } from 'vitest';
import type { UserProfile } from '../types.js';
import { meCommand } from './me.js';
import { makeCtx, text } from './testUtils.js';

const PROFILE: UserProfile = {
  id: 'user-1',
  displayName: 'Adele Vance',
  mail: 'adelev@contoso.com',
  userPrincipalName: 'adelev@contoso.onmicrosoft.com',
  jobTitle: 'Engineer',
};

describe('me', () => {
  it('fetches the profile and renders it as human-readable text', async () => {
    const { ctx, client, out } = makeCtx([]);
    vi.spyOn(client, 'get').mockResolvedValue(PROFILE);

    await meCommand.handler(ctx);

    expect(text(out)).toContain('Adele Vance');
    expect(text(out)).toContain('adelev@contoso.com');
    expect(text(out)).toContain('user-1');
  });

  it('--json emits the raw profile object, unmodified', async () => {
    const { ctx, client, out } = makeCtx(['--json']);
    vi.spyOn(client, 'get').mockResolvedValue(PROFILE);

    await meCommand.handler(ctx);

    expect(JSON.parse(text(out))).toEqual(PROFILE);
  });

  it('without --json, output is human text, not a JSON blob', async () => {
    const { ctx, client, out } = makeCtx([]);
    vi.spyOn(client, 'get').mockResolvedValue(PROFILE);

    await meCommand.handler(ctx);

    expect(() => JSON.parse(text(out))).toThrow();
  });

  it('calls the documented GET /me endpoint with the exact $select set that populates UserProfile', async () => {
    const { ctx, client } = makeCtx([]);
    const get = vi.spyOn(client, 'get').mockResolvedValue(PROFILE);

    await meCommand.handler(ctx);

    expect(get).toHaveBeenCalledTimes(1);
    expect(get.mock.calls[0]?.[0]).toBe('/me');
    expect(get.mock.calls[0]?.[1]).toEqual({
      $select: 'id,displayName,mail,userPrincipalName,jobTitle',
    });
  });

  it('propagates a client error rather than swallowing it (auth-failure smoke test)', async () => {
    const { ctx, client } = makeCtx([]);
    const authError = new Error('InvalidAuthenticationToken: token expired');
    vi.spyOn(client, 'get').mockRejectedValue(authError);

    await expect(meCommand.handler(ctx)).rejects.toBe(authError);
  });

  it('never touches the network directly — only through the stubbed client method', async () => {
    const { ctx, client, fetchImpl } = makeCtx([]);
    vi.spyOn(client, 'get').mockResolvedValue(PROFILE);

    await meCommand.handler(ctx);

    expect(fetchImpl).toHaveBeenCalledTimes(0);
  });
});
