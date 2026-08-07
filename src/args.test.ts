import { describe, it, expect } from 'vitest';
import { parseArgs } from './args.js';
import { UsageError } from './errors.js';

describe('parseArgs — flag forms', () => {
  it('parses --flag=value', () => {
    const a = parseArgs(['--subject=hello world']);
    expect(a.str('subject')).toBe('hello world');
  });

  it('parses --flag value (space-separated)', () => {
    const a = parseArgs(['--subject', 'hello world']);
    expect(a.str('subject')).toBe('hello world');
  });

  it('parses a bare boolean --flag as true', () => {
    const a = parseArgs(['--json']);
    expect(a.bool('json')).toBe(true);
    expect(a.has('json')).toBe(true);
  });

  it('treats a bare flag followed by another flag as boolean (does not consume it as a value)', () => {
    const a = parseArgs(['--json', '--top', '5']);
    expect(a.bool('json')).toBe(true);
    expect(a.num('top')).toBe(5);
  });

  it('parses --no-flag as an explicit false', () => {
    const a = parseArgs(['--no-color']);
    expect(a.bool('color')).toBe(false);
    expect(a.has('color')).toBe(true);
  });

  it('parses short flags -x', () => {
    const a = parseArgs(['-q', 'urgent']);
    expect(a.str('q')).toBe('urgent');
  });

  it('parses a short bare boolean flag', () => {
    const a = parseArgs(['-v']);
    expect(a.bool('v')).toBe(true);
  });

  it('parses short flag with = form', () => {
    const a = parseArgs(['-q=urgent']);
    expect(a.str('q')).toBe('urgent');
  });

  it('does not treat a lone "-" as a flag (stdin convention)', () => {
    const a = parseArgs(['--file', '-']);
    expect(a.str('file')).toBe('-');
  });
});

describe('parseArgs — repeatable flags', () => {
  it('accumulates repeated --attach occurrences in order', () => {
    const a = parseArgs(['--attach', 'a.pdf', '--attach', 'b.png']);
    expect(a.strAll('attach')).toEqual(['a.pdf', 'b.png']);
  });

  it('does not comma-split strAll values — callers get the raw occurrence', () => {
    const a = parseArgs(['--to', 'a@example.com,b@example.com']);
    expect(a.strAll('to')).toEqual(['a@example.com,b@example.com']);
  });

  it('str() returns the LAST occurrence when a flag repeats', () => {
    const a = parseArgs(['--top', '5', '--top', '10']);
    expect(a.str('top')).toBe('10');
  });

  it('strAll returns an empty array for an absent flag', () => {
    const a = parseArgs([]);
    expect(a.strAll('attach')).toEqual([]);
  });
});

describe('parseArgs — positionals and the -- terminator', () => {
  it('collects bare tokens as positionals', () => {
    const a = parseArgs(['list', 'inbox']);
    expect(a.positionals).toEqual(['list', 'inbox']);
  });

  it('treats everything after -- as positional, even flag-shaped tokens', () => {
    const a = parseArgs(['send', '--', '--not-a-flag', '-x']);
    expect(a.positionals).toEqual(['send', '--not-a-flag', '-x']);
  });

  it('a second -- after termination is itself just a positional', () => {
    const a = parseArgs(['--', '--', 'a']);
    expect(a.positionals).toEqual(['--', 'a']);
  });

  it('preserves the original argv on .raw', () => {
    const argv = ['--json', 'inbox'];
    const a = parseArgs(argv);
    expect(a.raw).toEqual(['--json', 'inbox']);
  });
});

describe('parseArgs — require()', () => {
  it('returns the value when present', () => {
    const a = parseArgs(['--id', 'abc123']);
    expect(a.require('id')).toBe('abc123');
  });

  it('throws UsageError naming the flag when missing', () => {
    const a = parseArgs([]);
    expect(() => a.require('id')).toThrow(UsageError);
    expect(() => a.require('id')).toThrow(/--id/);
  });
});

describe('parseArgs — requirePositional()', () => {
  it('returns the positional at the given index', () => {
    const a = parseArgs(['move', 'msg-123', 'archive']);
    expect(a.requirePositional(1, 'message id')).toBe('msg-123');
  });

  it('throws UsageError naming the label when missing', () => {
    const a = parseArgs(['move']);
    expect(() => a.requirePositional(1, 'message id')).toThrow(UsageError);
    expect(() => a.requirePositional(1, 'message id')).toThrow(/message id/);
  });
});

describe('parseArgs — num()', () => {
  it('parses a valid integer', () => {
    const a = parseArgs(['--top', '25']);
    expect(a.num('top')).toBe(25);
  });

  it('parses a valid float', () => {
    const a = parseArgs(['--ratio', '1.5']);
    expect(a.num('ratio')).toBe(1.5);
  });

  it('returns undefined when the flag is absent', () => {
    const a = parseArgs([]);
    expect(a.num('top')).toBeUndefined();
  });

  it('throws UsageError on a non-numeric value rather than yielding NaN', () => {
    const a = parseArgs(['--top', 'abc']);
    expect(() => a.num('top')).toThrow(UsageError);
    expect(() => a.num('top')).toThrow(/--top/);
  });

  it('throws UsageError when the flag was given as a bare boolean', () => {
    const a = parseArgs(['--top']);
    expect(() => a.num('top')).toThrow(UsageError);
  });

  it('throws UsageError on an empty string value', () => {
    const a = parseArgs(['--top=']);
    expect(() => a.num('top')).toThrow(UsageError);
  });
});

describe('parseArgs — has() and bool() on absent flags', () => {
  it('has() is false for an absent flag', () => {
    const a = parseArgs([]);
    expect(a.has('json')).toBe(false);
  });

  it('bool() is false for an absent flag', () => {
    const a = parseArgs([]);
    expect(a.bool('json')).toBe(false);
  });

  it('str() returns the fallback for an absent flag', () => {
    const a = parseArgs([]);
    expect(a.str('folder', 'inbox')).toBe('inbox');
  });

  it('str() returns undefined for an absent flag with no fallback', () => {
    const a = parseArgs([]);
    expect(a.str('folder')).toBeUndefined();
  });
});

describe('parseArgs — realistic composite command line', () => {
  it('handles a send-with-attachments style invocation', () => {
    const a = parseArgs([
      'send',
      '--to',
      'a@example.com',
      '--to',
      'b@example.com',
      '--subject=Q3 numbers',
      '--attach',
      './report.pdf',
      '--attach',
      './chart.png',
      '--no-save-sent',
      '--',
      '--literal-body-not-a-flag',
    ]);

    expect(a.positionals).toEqual(['send', '--literal-body-not-a-flag']);
    expect(a.strAll('to')).toEqual(['a@example.com', 'b@example.com']);
    expect(a.str('subject')).toBe('Q3 numbers');
    expect(a.strAll('attach')).toEqual(['./report.pdf', './chart.png']);
    expect(a.bool('save-sent')).toBe(false);
  });
});

describe('parseArgs — bool() strict boolean coercion', () => {
  it('a bare --yes is true', () => {
    const a = parseArgs(['--yes']);
    expect(a.bool('yes')).toBe(true);
  });

  it('--no-yes is false', () => {
    const a = parseArgs(['--no-yes']);
    expect(a.bool('yes')).toBe(false);
  });

  it('THE CRITICAL ONE: --yes=false is false — this previously returned true and permanently deleted mail', () => {
    const a = parseArgs(['--yes=false']);
    expect(a.bool('yes')).toBe(false);
  });

  it.each(['0', 'no', 'off', ''])('--yes=%s is false', (value) => {
    const a = parseArgs([`--yes=${value}`]);
    expect(a.bool('yes')).toBe(false);
  });

  it.each(['true', '1', 'yes', 'on'])('--yes=%s is true', (value) => {
    const a = parseArgs([`--yes=${value}`]);
    expect(a.bool('yes')).toBe(true);
  });

  it('is case-insensitive: --yes=FALSE is false', () => {
    const a = parseArgs(['--yes=FALSE']);
    expect(a.bool('yes')).toBe(false);
  });

  it('trims surrounding whitespace: --yes= true  is true', () => {
    const a = parseArgs(['--yes= true ']);
    expect(a.bool('yes')).toBe(true);
  });

  it('throws UsageError on an ambiguous value, naming the value', () => {
    const a = parseArgs(['--json', 'inbox']);
    expect(() => a.bool('json')).toThrow(UsageError);
    expect(() => a.bool('json')).toThrow(/inbox/);
  });
});

describe('parseArgs — names()', () => {
  it('returns every flag name seen on the command line', () => {
    const a = parseArgs(['--json', '--top', '5', '--no-color']);
    expect(a.names().sort()).toEqual(['color', 'json', 'top']);
  });

  it('returns an empty array for a positional-only argv', () => {
    const a = parseArgs(['list', 'inbox']);
    expect(a.names()).toEqual([]);
  });
});
