/**
 * Turn an attacker-controlled filename into one that is safe to create on
 * ANY host filesystem — specifically including Windows, whose rules are
 * stricter than POSIX in ways that are easy to miss on a mac.
 *
 * The name on a Graph attachment is remote data chosen by whoever sent the
 * mail. `basename()` alone defeats `../` traversal, but leaves a family of
 * Windows-only hazards intact:
 *
 *   `report:secret`  -> NTFS alternate data stream; the bytes after the colon
 *                       vanish from an ordinary directory listing.
 *   `CON`, `NUL`,    -> reserved DOS device names. Opening one addresses the
 *   `AUX`, `COM1`       DEVICE, not a file; the write can hang or be discarded.
 *   `invoice.pdf.`   -> trailing dot/space is silently stripped by Win32, so
 *   `invoice.pdf `      two distinct attachments collapse onto one path.
 *   `<`, `>`, `|`,   -> outright illegal in a Win32 filename.
 *   `"`, `*`, `?`
 *
 * Applied unconditionally rather than behind a `process.platform` check: a
 * name that is dangerous on Windows is merely ugly on POSIX, and downloads
 * synced or copied between machines should not change meaning en route.
 */

/** Illegal on Win32, plus the ADS colon. Control chars 0x00-0x1f too. */
// eslint-disable-next-line no-control-regex
const ILLEGAL_CHARS_RE = /[<>:"|?*\u0000-\u001f]/g;

/** DOS device names, with or without an extension (`CON`, `CON.txt`). */
const RESERVED_NAMES = new Set([
  'con',
  'prn',
  'aux',
  'nul',
  ...Array.from({ length: 9 }, (_, i) => `com${i + 1}`),
  ...Array.from({ length: 9 }, (_, i) => `lpt${i + 1}`),
]);

/** Win32 MAX_PATH leaves little room; keep a component well under it. */
const MAX_COMPONENT_LENGTH = 180;

function stem(name: string): string {
  const dot = name.indexOf('.');
  return dot === -1 ? name : name.slice(0, dot);
}

/**
 * @param name     the raw, untrusted filename (already basename'd by caller)
 * @param fallback used when sanitizing leaves nothing usable
 */
export function sanitizeFilename(name: string, fallback: string): string {
  let out = name.replace(ILLEGAL_CHARS_RE, '_');

  // Win32 strips these silently, which would let `a.pdf` and `a.pdf.` collide.
  out = out.replace(/[. ]+$/, '');
  out = out.trim();

  if (out === '' || out === '.' || out === '..') return fallback;

  // Prefix rather than replace, so the original name stays legible.
  if (RESERVED_NAMES.has(stem(out).toLowerCase())) {
    out = `_${out}`;
  }

  if (out.length > MAX_COMPONENT_LENGTH) {
    // Preserve the extension — it is what makes the file openable.
    const lastDot = out.lastIndexOf('.');
    const ext = lastDot > 0 ? out.slice(lastDot, lastDot + 16) : '';
    out = out.slice(0, MAX_COMPONENT_LENGTH - ext.length) + ext;
  }

  // Re-check: truncation can re-expose a trailing dot or space.
  out = out.replace(/[. ]+$/, '');
  return out === '' ? fallback : out;
}
