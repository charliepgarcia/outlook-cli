import { describe, it, expect } from 'vitest';
import { sanitizeFilename } from './safeFilename.js';

const FALLBACK = 'fallback.bin';

describe('sanitizeFilename — ordinary names', () => {
  it('passes an unremarkable name through unchanged', () => {
    expect(sanitizeFilename('report.pdf', FALLBACK)).toBe('report.pdf');
  });
});

describe('sanitizeFilename — NTFS alternate data stream colon', () => {
  it('neutralizes the colon in report:secret', () => {
    const result = sanitizeFilename('report:secret', FALLBACK);
    expect(result).not.toContain(':');
    expect(result).toBe('report_secret');
  });
});

describe('sanitizeFilename — Windows reserved device names', () => {
  it.each(['CON', 'con', 'NUL', 'AUX', 'PRN', 'COM1', 'LPT9'])(
    'defuses the bare reserved name %s',
    (reserved) => {
      const result = sanitizeFilename(reserved, FALLBACK);
      expect(result).not.toBe(reserved);
      expect(result.toLowerCase()).not.toBe(reserved.toLowerCase());
      // Implementation prefixes an underscore rather than replacing.
      expect(result).toContain(reserved);
    },
  );

  it('defuses the reserved name even with an extension (CON.txt)', () => {
    const result = sanitizeFilename('CON.txt', FALLBACK);
    expect(result).not.toBe('CON.txt');
    expect(result.toLowerCase()).not.toBe('con.txt');
    expect(result).toBe('_CON.txt');
  });
});

describe('sanitizeFilename — illegal Win32 characters', () => {
  it('replaces < > " | ? * with an underscore', () => {
    const result = sanitizeFilename('a<b>c"d|e?f*g', FALLBACK);
    expect(result).toBe('a_b_c_d_e_f_g');
    for (const ch of ['<', '>', '"', '|', '?', '*']) {
      expect(result).not.toContain(ch);
    }
  });
});

describe('sanitizeFilename — control characters', () => {
  it('replaces control characters (\\u0000-\\u001f) with an underscore', () => {
    const result = sanitizeFilename('a\u0000b\u001fc', FALLBACK);
    expect(result).toBe('a_b_c');
    expect(result).not.toMatch(/[\u0000-\u001f]/);
  });
});

describe('sanitizeFilename — trailing dots and spaces', () => {
  it('strips a trailing dot so it cannot collide with the undotted name', () => {
    expect(sanitizeFilename('invoice.pdf.', FALLBACK)).toBe('invoice.pdf');
  });

  it('strips a trailing space so it cannot collide with the unspaced name', () => {
    expect(sanitizeFilename('invoice.pdf ', FALLBACK)).toBe('invoice.pdf');
  });

  it('collapses both forms onto the same result (the collision this closes)', () => {
    const bare = sanitizeFilename('invoice.pdf', FALLBACK);
    expect(sanitizeFilename('invoice.pdf.', FALLBACK)).toBe(bare);
    expect(sanitizeFilename('invoice.pdf ', FALLBACK)).toBe(bare);
  });
});

describe('sanitizeFilename — collapses to fallback', () => {
  it.each(['', '.', '..', '....', '   ', '. .'])(
    'returns the fallback for %j',
    (input) => {
      expect(sanitizeFilename(input, FALLBACK)).toBe(FALLBACK);
    },
  );
});

describe('sanitizeFilename — over-long names', () => {
  it('truncates a long name with an extension but preserves the extension', () => {
    const longName = 'c'.repeat(300) + '.pdf';
    const result = sanitizeFilename(longName, FALLBACK);

    expect(result.length).toBeLessThan(longName.length);
    expect(result.endsWith('.pdf')).toBe(true);
    expect(result.endsWith('.')).toBe(false);
    expect(result.endsWith(' ')).toBe(false);
  });

  it('truncates a long extensionless name and never ends in a dot or space', () => {
    // A run of spaces straddling the truncation boundary would re-expose a
    // trailing space if the implementation didn't re-check after slicing.
    const longName = 'a'.repeat(50) + ' '.repeat(200) + 'b'.repeat(50);
    const result = sanitizeFilename(longName, FALLBACK);

    expect(result.length).toBeLessThan(longName.length);
    expect(/[. ]$/.test(result)).toBe(false);
  });
});
