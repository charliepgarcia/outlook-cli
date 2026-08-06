import { describe, it, expect } from 'vitest';
import { renderRequest } from './dryRun.js';
import type { RequestSpec } from './types.js';

describe('renderRequest', () => {
  it('masks a Bearer token in the Authorization header', () => {
    const spec: RequestSpec = {
      method: 'GET',
      url: 'https://graph.microsoft.com/v1.0/me',
      headers: {
        Authorization: 'Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiJ9.abcdefghijklmnop.signature',
        Prefer: 'IdType="ImmutableId"',
      },
    };

    const out = renderRequest(spec);

    expect(out).not.toContain('eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiJ9');
    expect(out).toContain('GET https://graph.microsoft.com/v1.0/me');
    expect(out).toContain('Prefer: IdType="ImmutableId"');
  });

  it('is safe to call with no token / no headers at all', () => {
    const spec: RequestSpec = { method: 'GET', url: 'https://graph.microsoft.com/v1.0/me' };

    expect(() => renderRequest(spec)).not.toThrow();
    const out = renderRequest(spec);
    expect(out).toBe('GET https://graph.microsoft.com/v1.0/me');
  });

  it('pretty-prints a redacted JSON body', () => {
    const spec: RequestSpec = {
      method: 'POST',
      url: 'https://graph.microsoft.com/v1.0/me/sendMail',
      headers: { 'Content-Type': 'application/json' },
      body: {
        message: { subject: 'hi', toRecipients: [{ emailAddress: { address: 'a@b.com' } }] },
        saveToSentItems: false,
      },
    };

    const out = renderRequest(spec);

    expect(out).toContain('POST https://graph.microsoft.com/v1.0/me/sendMail');
    expect(out).toContain('"subject": "hi"');
    expect(out).toContain('"saveToSentItems": false');
  });

  it('redacts secret-keyed fields inside the body', () => {
    const spec: RequestSpec = {
      method: 'POST',
      url: 'https://example.com/token',
      body: { access_token: 'super-secret-value-that-is-long', ok: true },
    };

    const out = renderRequest(spec);

    expect(out).not.toContain('super-secret-value-that-is-long');
    expect(out).toContain('[REDACTED]');
  });
});
