import { describe, expect, it } from 'vitest';
import { parseArgs } from './args.js';
import { optionSchemaForCommand } from './cliOptions.js';
import { UsageError } from './errors.js';

function parse(command: string, argv: string[]) {
  return parseArgs(argv, optionSchemaForCommand(command));
}

describe('strict CLI option parsing', () => {
  it('rejects a misspelled --dry-run before command execution', () => {
    expect(() => parse('send', ['--dryrun'])).toThrow(UsageError);
    expect(() => parse('send', ['--dryrun'])).toThrow(/Unknown option: --dryrun/);
  });

  it('rejects a valid option when it belongs to a different command', () => {
    expect(() => parse('send', ['--force'])).toThrow(/Unknown option: --force/);
  });

  it('parses explicit false confirmation safely in both supported forms', () => {
    expect(parse('purge', ['message-id', '--yes=false']).bool('yes')).toBe(false);
    expect(parse('purge', ['message-id', '--yes', 'false']).bool('yes')).toBe(false);
  });

  it('does not consume ordinary positionals after a bare boolean', () => {
    const args = parse('purge', ['--yes', 'message-id']);
    expect(args.bool('yes')).toBe(true);
    expect(args.positionals).toEqual(['message-id']);
  });

  it('rejects invalid explicit boolean values instead of treating them as true', () => {
    expect(() => parse('purge', ['message-id', '--yes=nope'])).toThrow(
      /--yes expects true or false/,
    );
  });

  it('accepts documented repeatable send options and --no-save-sent', () => {
    const args = parse('send', [
      '--to',
      'a@example.com',
      '--to=b@example.com',
      '--subject',
      'Hello',
      '--no-save-sent',
      '--dry-run',
    ]);

    expect(args.strAll('to')).toEqual(['a@example.com', 'b@example.com']);
    expect(args.str('subject')).toBe('Hello');
    expect(args.bool('save-sent')).toBe(false);
    expect(args.bool('dry-run')).toBe(true);
  });

  it('supports command help through the -h alias', () => {
    expect(parse('list', ['-h']).bool('help')).toBe(true);
  });

  it('requires values for value-taking options', () => {
    expect(() => parse('send', ['--subject'])).toThrow(/--subject requires a value/);
  });

  it.each([
    ['auth', ['status', '--json']],
    ['me', ['--json']],
    ['list', ['--folder', 'inbox', '--limit', '10', '--all', '--unread']],
    ['get', ['id', '--html']],
    ['search', ['query', '--folder', 'inbox', '--limit', '10', '--all']],
    ['folders', ['create', '--name', 'Archive', '--parent', 'inbox', '--all']],
    ['attachments', ['download', 'id', '--out', '.', '--all', '--force']],
    ['reply', ['id', '--body', 'Thanks']],
    ['reply-all', ['id', '--body', 'Thanks']],
    ['forward', ['id', '--to', 'a@example.com', '--body', 'FYI']],
    [
      'draft',
      [
        'create', '--to', 'a@example.com', '--cc', 'c@example.com',
        '--bcc', 'b@example.com', '--subject', 'Hi', '--body', 'Text',
        '--body-file', 'body.txt', '--attach', 'a.txt', '--html', '--all',
        '--limit', '10',
      ],
    ],
    ['move', ['id', '--folder', 'archive']],
    ['delete', ['id', '--dry-run']],
    ['purge', ['id', '--yes', '--dry-run']],
    ['mark-read', ['id', '--dry-run']],
    ['mark-unread', ['id', '--dry-run']],
    ['flag', ['id', '--dry-run']],
    ['unflag', ['id', '--dry-run']],
  ] as const)('accepts existing options for %s', (command, argv) => {
    expect(() => parse(command, [...argv])).not.toThrow();
  });
});
