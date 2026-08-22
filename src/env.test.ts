/**
 * Tests for the afk.env auto-loader (`src/env.ts`).
 *
 * All tests use `loadEnvFile(path)` with a temp file rather than touching the
 * real `~/.afk/config/afk.env`, and every test restores `process.env` after
 * running so side-effects never leak across cases.
 */

import { describe, expect, it, afterEach } from 'vitest';
import { writeFileSync, mkdirSync, rmSync } from 'node:fs';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import { loadEnvFile, parseEnvLines } from './env.js';

const TMP_DIR = join(tmpdir(), `outlook-cli-env-test-${process.pid}`);

function tmpFile(name: string, content: string): string {
  mkdirSync(TMP_DIR, { recursive: true });
  const p = join(TMP_DIR, name);
  writeFileSync(p, content, 'utf-8');
  return p;
}

afterEach(() => {
  // Clean env keys we may have set.
  delete process.env['TEST_ENV_KEY'];
  delete process.env['ANOTHER'];
  delete process.env['GRAPH_CLIENT_ID'];
  delete process.env['SPACED'];
  delete process.env['EMPTY_VAL'];
  try {
    rmSync(TMP_DIR, { recursive: true, force: true });
  } catch {
    // ignore
  }
});

describe('parseEnvLines', () => {
  it('parses simple KEY=value pairs', () => {
    expect(parseEnvLines('A=1\nB=two')).toEqual([
      ['A', '1'],
      ['B', 'two'],
    ]);
  });

  it('skips blank lines and comments', () => {
    const input = '# comment\n\nKEY=val\n  # indented comment\n\n';
    expect(parseEnvLines(input)).toEqual([['KEY', 'val']]);
  });

  it('strips double quotes from values', () => {
    expect(parseEnvLines('K="hello world"')).toEqual([['K', 'hello world']]);
  });

  it('strips single quotes from values', () => {
    expect(parseEnvLines("K='hello world'")).toEqual([['K', 'hello world']]);
  });

  it('does not strip mismatched quotes', () => {
    expect(parseEnvLines('K="hello\'')).toEqual([['K', '"hello\'']]);
  });

  it('handles value with = in it', () => {
    expect(parseEnvLines('K=a=b=c')).toEqual([['K', 'a=b=c']]);
  });

  it('skips lines without =', () => {
    expect(parseEnvLines('no-equals-here')).toEqual([]);
  });
});

describe('loadEnvFile', () => {
  it('loads keys from file into process.env', () => {
    const path = tmpFile('basic.env', 'TEST_ENV_KEY=hello\nANOTHER=world');
    loadEnvFile(path);

    expect(process.env['TEST_ENV_KEY']).toBe('hello');
    expect(process.env['ANOTHER']).toBe('world');
  });

  it('does NOT overwrite existing env vars', () => {
    process.env['TEST_ENV_KEY'] = 'original';
    const path = tmpFile('no-overwrite.env', 'TEST_ENV_KEY=from-file');
    loadEnvFile(path);

    expect(process.env['TEST_ENV_KEY']).toBe('original');
  });

  it('missing file is a silent no-op', () => {
    // No error thrown, no env changes.
    expect(() => loadEnvFile('/nonexistent/path/afk.env')).not.toThrow();
  });

  it('skips comments and blanks in the file', () => {
    const path = tmpFile(
      'comments.env',
      '# This is a comment\n\nGRAPH_CLIENT_ID=abc123\n  \n# trailing',
    );
    loadEnvFile(path);

    expect(process.env['GRAPH_CLIENT_ID']).toBe('abc123');
  });

  it('handles quoted values', () => {
    const path = tmpFile('quoted.env', 'TEST_ENV_KEY="quoted value"');
    loadEnvFile(path);

    expect(process.env['TEST_ENV_KEY']).toBe('quoted value');
  });

  it('handles values with spaces around the =', () => {
    const path = tmpFile('spaced.env', ' SPACED = value ');
    loadEnvFile(path);

    expect(process.env['SPACED']).toBe('value');
  });

  it('handles empty values', () => {
    const path = tmpFile('empty.env', 'EMPTY_VAL=');
    loadEnvFile(path);

    expect(process.env['EMPTY_VAL']).toBe('');
  });
});
