/**
 * Tests for flagSpec.ts — the per-command flag allowlist that fail-closes on
 * typos (see the file header there for the `--dryrun` incident this exists
 * to prevent).
 *
 * The command modules are imported directly, NOT via cli.ts: cli.ts runs
 * top-level routing/registry-building code on import and does not export its
 * COMMANDS array, so importing it here would be both unnecessary and risky.
 * Each command module already exports its CommandSpec(s) individually (as
 * every sibling command test file relies on), so the registry is rebuilt
 * locally from those exports for the completeness assertions below.
 */

import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { assertKnownFlags, COMMAND_FLAGS, GLOBAL_FLAGS } from './flagSpec.js';
import { UsageError } from './errors.js';

import { attachmentsCommand } from './commands/attachments.js';
import { authCommand } from './commands/auth.js';
import { contactsCommand } from './commands/contacts.js';
import { draftCommand } from './commands/draft.js';
import {
  flagCommand,
  markReadCommand,
  markUnreadCommand,
  unflagCommand,
} from './commands/flags.js';
import { foldersCommand } from './commands/folders.js';
import { getCommand } from './commands/get.js';
import { listCommand } from './commands/list.js';
import { meCommand } from './commands/me.js';
import { deleteCommand, moveCommand, purgeCommand } from './commands/move.js';
import {
  forwardCommand,
  replyAllCommand,
  replyCommand,
} from './commands/reply.js';
import { rulesCommand } from './commands/rules.js';
import { searchCommand } from './commands/search.js';
import { sendCommand } from './commands/send.js';

/**
 * Rebuilt independently of cli.ts's (unexported) COMMANDS array, from each
 * command module's own exports — the same exports cli.ts itself imports.
 * Keyed by the identifier cli.ts uses, so the drift test below can check
 * cli.ts's registry against this map by name.
 */
const COMMANDS_BY_IDENTIFIER = {
  authCommand,
  meCommand,
  listCommand,
  getCommand,
  searchCommand,
  foldersCommand,
  attachmentsCommand,
  contactsCommand,
  sendCommand,
  replyCommand,
  replyAllCommand,
  forwardCommand,
  draftCommand,
  moveCommand,
  deleteCommand,
  purgeCommand,
  markReadCommand,
  markUnreadCommand,
  flagCommand,
  unflagCommand,
  rulesCommand,
} as const;

const REGISTERED_COMMANDS = Object.values(COMMANDS_BY_IDENTIFIER);

describe('assertKnownFlags — known flags pass', () => {
  it("accepts flags that are in the command's own list (list: folder, limit)", () => {
    expect(() => assertKnownFlags('list', ['folder', 'limit'])).not.toThrow();
  });

  it('accepts json, dry-run, help, and h individually on a representative command', () => {
    expect(() => assertKnownFlags('list', ['json'])).not.toThrow();
    expect(() => assertKnownFlags('list', ['dry-run'])).not.toThrow();
    expect(() => assertKnownFlags('list', ['help'])).not.toThrow();
    expect(() => assertKnownFlags('list', ['h'])).not.toThrow();
  });

  it('accepts GLOBAL_FLAGS on every registered command', () => {
    for (const command of REGISTERED_COMMANDS) {
      expect(() => assertKnownFlags(command.name, [...GLOBAL_FLAGS])).not.toThrow();
    }
  });
});

describe('assertKnownFlags — unknown flags throw', () => {
  it('throws UsageError for a flag not in the allowlist', () => {
    expect(() => assertKnownFlags('list', ['bogus'])).toThrow(UsageError);
  });

  it('THE CRITICAL ONE: rejects the --dryrun typo and suggests --dry-run — this is the typo that used to send real mail', () => {
    expect(() => assertKnownFlags('send', ['dryrun'])).toThrow(UsageError);
    expect(() => assertKnownFlags('send', ['dryrun'])).toThrow(/Did you mean --dry-run\?/);
  });

  it("throws for an unknown flag on a command with an empty own-flag list ('me')", () => {
    expect(() => assertKnownFlags('me', ['bogus'])).toThrow(UsageError);
  });
});

describe('assertKnownFlags — unregistered command name (fail closed)', () => {
  it('throws UsageError rather than silently allowing everything', () => {
    expect(() => assertKnownFlags('not-a-real-command', [])).toThrow(UsageError);
    expect(() => assertKnownFlags('not-a-real-command', [])).toThrow(
      /no flag specification registered/,
    );
  });
});

describe('COMMAND_FLAGS — completeness against the real command registry', () => {
  it("knows every command cli.ts registers (catches registry drift, e.g. 'rules' added without a flag entry)", () => {
    const source = readFileSync(new URL('./cli.ts', import.meta.url), 'utf8');
    const block = /const COMMANDS\b[^=]*=\s*\[([\s\S]*?)\];/.exec(source)?.[1] ?? '';
    const registered = block.match(/\b\w+Command\b/g) ?? [];
    expect(registered.length).toBeGreaterThan(0);
    const known = Object.keys(COMMANDS_BY_IDENTIFIER);
    for (const identifier of registered) {
      expect(known, `cli.ts registers ${identifier}; add it to this test and to COMMAND_FLAGS`).toContain(identifier);
    }
    expect(registered.length).toBe(known.length);
  });

  it('accepts the rules create and contacts flags, and rejects typos on them', () => {
    expect(() =>
      assertKnownFlags('rules', ['name', 'folder', 'from', 'sender-contains', 'mark-read', 'stop-processing', 'enabled', 'sequence']),
    ).not.toThrow();
    expect(() => assertKnownFlags('contacts', ['top', 'all', 'limit', 'select', 'filter'])).not.toThrow();
    expect(() => assertKnownFlags('rules', ['dryrun'])).toThrow(/Did you mean --dry-run\?/);
    expect(() => assertKnownFlags('contacts', ['tpo'])).toThrow(UsageError);
  });

  it('has an entry for every registered command', () => {
    for (const command of REGISTERED_COMMANDS) {
      expect(COMMAND_FLAGS[command.name]).toBeDefined();
    }
  });

  it('has no stale entries that do not correspond to a registered command', () => {
    const registeredNames = new Set(REGISTERED_COMMANDS.map((command) => command.name));
    for (const name of Object.keys(COMMAND_FLAGS)) {
      expect(registeredNames.has(name)).toBe(true);
    }
  });
});
