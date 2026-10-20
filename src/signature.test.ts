import { mkdtemp, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { beforeAll, describe, expect, it } from 'vitest';
import { parseArgs } from './args.js';
import { UsageError } from './errors.js';
import {
  applySignature,
  appendTextSignature,
  resolveSignature,
  signatureFlagRequested,
} from './signature.js';

let dir: string;
let textSig: string;
let htmlSig: string;

beforeAll(async () => {
  dir = await mkdtemp(join(tmpdir(), 'outlook-sig-'));
  textSig = join(dir, 'sig.txt');
  htmlSig = join(dir, 'sig.html');
  await writeFile(textSig, 'Griffin Long\nAgent AFK\n');
  await writeFile(htmlSig, '<b>Griffin</b>');
});

describe('resolveSignature', () => {
  it('returns nothing when nothing is configured', async () => {
    expect(await resolveSignature(parseArgs([]), {})).toBeUndefined();
  });

  it('reads both env-configured files', async () => {
    const sig = await resolveSignature(parseArgs([]), {
      OUTLOOK_SIGNATURE_FILE: textSig,
      OUTLOOK_SIGNATURE_HTML_FILE: htmlSig,
    });
    expect(sig).toEqual({ text: 'Griffin Long\nAgent AFK\n', html: '<b>Griffin</b>' });
  });

  it('--no-signature beats env configuration', async () => {
    const sig = await resolveSignature(parseArgs(['--no-signature']), {
      OUTLOOK_SIGNATURE_FILE: textSig,
    });
    expect(sig).toBeUndefined();
  });

  it('--signature=false is an opt-out too', async () => {
    const sig = await resolveSignature(parseArgs(['--signature=false']), {
      OUTLOOK_SIGNATURE_FILE: textSig,
    });
    expect(sig).toBeUndefined();
  });

  it('--signature <path> overrides env and picks format by extension', async () => {
    const env = { OUTLOOK_SIGNATURE_FILE: textSig };
    expect(await resolveSignature(parseArgs(['--signature', htmlSig]), env)).toEqual({
      html: '<b>Griffin</b>',
    });
    expect(await resolveSignature(parseArgs(['--signature', textSig]), {})).toEqual({
      text: 'Griffin Long\nAgent AFK\n',
    });
  });

  it('bare --signature with no path is a usage error', async () => {
    await expect(
      resolveSignature(parseArgs(['--signature']), {}),
    ).rejects.toBeInstanceOf(UsageError);
  });

  it('an unreadable configured file is a hard error naming its source', async () => {
    const err = await resolveSignature(parseArgs([]), {
      OUTLOOK_SIGNATURE_FILE: join(dir, 'missing.txt'),
    }).catch((e: unknown) => e);
    expect(err).toBeInstanceOf(UsageError);
    expect(String((err as Error).message)).toContain('OUTLOOK_SIGNATURE_FILE');
  });
});

describe('signatureFlagRequested', () => {
  it('distinguishes asking for a signature from opting out', () => {
    expect(signatureFlagRequested(parseArgs([]))).toBe(false);
    expect(signatureFlagRequested(parseArgs(['--no-signature']))).toBe(false);
    expect(signatureFlagRequested(parseArgs(['--signature']))).toBe(true);
    expect(signatureFlagRequested(parseArgs(['--signature', 'x.txt']))).toBe(true);
  });
});

describe('readSignatureFile (empty-file guard)', () => {
  it('item 1: empty file via env throws UsageError naming path and origin', async () => {
    const empty = join(dir, 'empty.txt');
    await writeFile(empty, '');
    const err = await resolveSignature(parseArgs([]), {
      OUTLOOK_SIGNATURE_FILE: empty,
    }).catch((e: unknown) => e);
    expect(err).toBeInstanceOf(UsageError);
    expect((err as Error).message).toContain(empty);
    expect((err as Error).message).toContain('OUTLOOK_SIGNATURE_FILE');
  });

  it('item 1: whitespace-only file via --signature throws UsageError naming path', async () => {
    const ws = join(dir, 'whitespace.txt');
    await writeFile(ws, '   \n  \t  \n');
    const err = await resolveSignature(parseArgs(['--signature', ws]), {}).catch(
      (e: unknown) => e,
    );
    expect(err).toBeInstanceOf(UsageError);
    expect((err as Error).message).toContain(ws);
  });
});

describe('applySignature', () => {
  it('appends a text signature with the RFC 3676 "-- " delimiter', () => {
    const body = applySignature(
      { contentType: 'text', content: 'Hi there\n\n' },
      { text: 'Griffin\n' },
    );
    expect(body).toEqual({ contentType: 'text', content: 'Hi there\n\n-- \nGriffin' });
  });

  it('signs an empty text body with just the delimiter block', () => {
    expect(appendTextSignature('', { text: 'G' })).toBe('-- \nG');
  });

  it('refuses an HTML-only signature on a plain-text body', () => {
    expect(() =>
      applySignature({ contentType: 'text', content: 'Hi' }, { html: '<b>G</b>' }),
    ).toThrow(UsageError);
  });

  it('inserts an HTML signature before </body>', () => {
    const body = applySignature(
      { contentType: 'html', content: '<html><body><p>Hi</p></body></html>' },
      { html: '<b>G</b>' },
    );
    expect(body.content).toBe(
      '<html><body><p>Hi</p><br><div class="outlook-cli-signature"><b>G</b></div></body></html>',
    );
  });

  it('item 4: multi-byte char (İ) before </body> does not shift splice offset', () => {
    // İ (U+0130) is one code unit but toLowerCase() maps it to two ('i\u0307'),
    // which would shift the old lastIndexOf-on-toLowerCase result by 1.
    const body = applySignature(
      { contentType: 'html', content: '<html><body>\u0130</body></html>' },
      { html: '<b>G</b>' },
    );
    expect(body.content).toBe(
      '<html><body>\u0130<br><div class="outlook-cli-signature"><b>G</b></div></body></html>',
    );
  });

  it('escapes a text signature used on an HTML body', () => {
    const body = applySignature(
      { contentType: 'html', content: '<p>Hi</p>' },
      { text: 'A & B <ceo>\nLine 2' },
    );
    expect(body.content).toBe(
      '<p>Hi</p><br><div class="outlook-cli-signature">A &amp; B &lt;ceo&gt;<br>Line 2</div>',
    );
  });
});
