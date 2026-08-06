import { describe, it, expect, afterEach } from 'vitest';
import { mkdtemp, writeFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import {
  MAX_INLINE_ATTACHMENT_BYTES,
  guessContentType,
  readFileAttachment,
  htmlToText,
} from './mime.js';
import { OutlookCliError } from './errors.js';

describe('MAX_INLINE_ATTACHMENT_BYTES', () => {
  it('is exactly 3 MB', () => {
    expect(MAX_INLINE_ATTACHMENT_BYTES).toBe(3 * 1024 * 1024);
  });
});

describe('guessContentType', () => {
  it.each([
    ['report.pdf', 'application/pdf'],
    ['photo.png', 'image/png'],
    ['photo.jpg', 'image/jpeg'],
    ['photo.jpeg', 'image/jpeg'],
    ['anim.gif', 'image/gif'],
    ['notes.txt', 'text/plain'],
    ['readme.md', 'text/markdown'],
    ['data.csv', 'text/csv'],
    ['payload.json', 'application/json'],
    ['archive.zip', 'application/zip'],
    [
      'doc.docx',
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    ],
    [
      'sheet.xlsx',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    ],
    [
      'deck.pptx',
      'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    ],
    ['page.html', 'text/html'],
    ['page.htm', 'text/html'],
  ])('maps %s -> %s', (filename, expected) => {
    expect(guessContentType(filename)).toBe(expected);
  });

  it('is case-insensitive on extension', () => {
    expect(guessContentType('REPORT.PDF')).toBe('application/pdf');
  });

  it('falls back to application/octet-stream for unknown extensions', () => {
    expect(guessContentType('archive.tar.gz')).toBe('application/octet-stream');
  });

  it('falls back to application/octet-stream for no extension', () => {
    expect(guessContentType('README')).toBe('application/octet-stream');
  });
});

describe('readFileAttachment', () => {
  let dir: string | undefined;

  afterEach(async () => {
    if (dir) {
      await rm(dir, { recursive: true, force: true });
      dir = undefined;
    }
  });

  it('round-trips base64 content and shapes a fileAttachment payload', async () => {
    dir = await mkdtemp(join(tmpdir(), 'outlook-cli-mime-'));
    const path = join(dir, 'hello.txt');
    const original = Buffer.from('Hello, Outlook CLI! \u00e9\u00e8 binary-ish \x00\x01\x02', 'utf8');
    await writeFile(path, original);

    const result = await readFileAttachment(path);

    expect(result['@odata.type']).toBe('#microsoft.graph.fileAttachment');
    expect(result.name).toBe('hello.txt');
    expect(result.contentType).toBe('text/plain');

    const decoded = Buffer.from(result.contentBytes, 'base64');
    expect(decoded.equals(original)).toBe(true);
  });

  it('uses only the basename, dropping directory components', async () => {
    dir = await mkdtemp(join(tmpdir(), 'outlook-cli-mime-'));
    const nested = join(dir, 'nested');
    await writeFile(join(dir, 'sibling-marker.txt'), 'x');
    const path = join(dir, 'report.pdf');
    await writeFile(path, Buffer.from('%PDF-1.4 fake'));
    void nested;

    const result = await readFileAttachment(path);
    expect(result.name).toBe('report.pdf');
    expect(result.name).not.toContain('/');
  });

  it('throws OutlookCliError naming the actual size when the file exceeds 3 MB', async () => {
    dir = await mkdtemp(join(tmpdir(), 'outlook-cli-mime-'));
    const path = join(dir, 'big.bin');
    const big = Buffer.alloc(MAX_INLINE_ATTACHMENT_BYTES + 1024, 7);
    await writeFile(path, big);

    await expect(readFileAttachment(path)).rejects.toThrow(OutlookCliError);
    await expect(readFileAttachment(path)).rejects.toThrow(
      new RegExp(String(big.length)),
    );
    await expect(readFileAttachment(path)).rejects.toThrow(/upload session/i);
  });

  it('accepts a file exactly at the 3 MB boundary', async () => {
    dir = await mkdtemp(join(tmpdir(), 'outlook-cli-mime-'));
    const path = join(dir, 'exact.bin');
    const exact = Buffer.alloc(MAX_INLINE_ATTACHMENT_BYTES, 1);
    await writeFile(path, exact);

    const result = await readFileAttachment(path);
    expect(Buffer.from(result.contentBytes, 'base64').length).toBe(
      MAX_INLINE_ATTACHMENT_BYTES,
    );
  });
});

describe('htmlToText', () => {
  it('removes script and style block content entirely', () => {
    const html =
      '<html><head><style>body{color:red}</style></head>' +
      '<body><script>alert("hi")</script><p>Real content</p></body></html>';
    const text = htmlToText(html);
    expect(text).not.toMatch(/color:red/);
    expect(text).not.toMatch(/alert/);
    expect(text).toContain('Real content');
  });

  it('converts <br> to a newline', () => {
    const text = htmlToText('Line one<br>Line two<br/>Line three');
    expect(text.split('\n').map((l) => l.trim())).toEqual([
      'Line one',
      'Line two',
      'Line three',
    ]);
  });

  it('converts </p> to a paragraph break', () => {
    const text = htmlToText('<p>First para</p><p>Second para</p>');
    expect(text).toContain('First para');
    expect(text).toContain('Second para');
    expect(text).not.toMatch(/<p>|<\/p>/);
  });

  it('decodes the common HTML entities', () => {
    const text = htmlToText(
      '&amp; &lt; &gt; &quot; &#39; &nbsp; done',
    );
    expect(text).toContain('& < > " \' ');
    expect(text).toContain('done');
  });

  it('strips remaining tags without emitting raw markup', () => {
    const text = htmlToText('<div><span class="x">Hi <b>there</b></span></div>');
    expect(text).not.toMatch(/[<>]/);
    expect(text).toContain('Hi');
    expect(text).toContain('there');
  });

  it('collapses runs of blank lines', () => {
    const text = htmlToText('<p>A</p><p></p><p></p><p>B</p>');
    expect(text).not.toMatch(/\n{3,}/);
  });

  it('returns an empty string for empty input', () => {
    expect(htmlToText('')).toBe('');
  });

  it('handles plain text with no tags at all', () => {
    expect(htmlToText('just plain text')).toBe('just plain text');
  });
});
