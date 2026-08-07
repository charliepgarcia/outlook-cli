import type { OptionSchema } from './args.js';

interface CommandOptions {
  boolean?: readonly string[];
  value?: readonly string[];
}

const COMMAND_OPTIONS: Readonly<Record<string, CommandOptions>> = {
  auth: {},
  me: {},
  list: { boolean: ['all', 'unread'], value: ['folder', 'limit'] },
  get: { boolean: ['html'] },
  search: { boolean: ['all'], value: ['folder', 'limit'] },
  folders: { boolean: ['all'], value: ['parent', 'name'] },
  attachments: { boolean: ['all', 'force'], value: ['out'] },
  send: {
    boolean: ['html', 'save-sent'],
    value: ['to', 'cc', 'bcc', 'subject', 'body', 'body-file', 'attach'],
  },
  reply: { value: ['body'] },
  'reply-all': { value: ['body'] },
  forward: { value: ['to', 'body'] },
  draft: {
    boolean: ['all', 'html'],
    value: ['limit', 'to', 'cc', 'bcc', 'subject', 'body', 'body-file', 'attach'],
  },
  move: { value: ['folder'] },
  delete: {},
  purge: { boolean: ['yes'] },
  'mark-read': {},
  'mark-unread': {},
  flag: {},
  unflag: {},
};

const COMMON_BOOLEAN = ['json', 'dry-run', 'help'] as const;

export function optionSchemaForCommand(command: string): OptionSchema {
  const specific = COMMAND_OPTIONS[command] ?? {};
  return {
    boolean: [...COMMON_BOOLEAN, ...(specific.boolean ?? [])],
    value: specific.value ?? [],
    aliases: { h: 'help' },
  };
}