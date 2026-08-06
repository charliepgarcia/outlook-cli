import { describe, it, expect } from 'vitest';
import {
  renderJson,
  renderFolders,
  renderAttachments,
  renderProfile,
  formatRecipients,
  formatDate,
  truncate,
} from './output.js';
import type { Attachment, MailFolder, Recipient, UserProfile } from './types.js';

// renderMessageList / renderMessage coverage lives in output.message.test.ts
// (split out to stay under the file-size gate — see that file's header comment).

function recipient(address: string, name?: string): Recipient {
  return { emailAddress: name ? { address, name } : { address } };
}

describe('renderJson', () => {
  it('pretty-prints with 2-space indent', () => {
    expect(renderJson({ a: 1 })).toBe(JSON.stringify({ a: 1 }, null, 2));
  });

  it('round-trips arrays and nested objects', () => {
    const data = { list: [1, 2, { nested: true }] };
    const out = renderJson(data);
    expect(JSON.parse(out)).toEqual(data);
  });
});

describe('formatDate', () => {
  it('formats a valid ISO date deterministically in UTC', () => {
    expect(formatDate('2024-03-05T14:30:00Z')).toBe('2024-03-05 14:30 UTC');
  });

  it('returns a placeholder for undefined', () => {
    expect(formatDate(undefined)).toBe('(no date)');
  });

  it('returns a placeholder for an unparseable string', () => {
    expect(formatDate('not-a-date')).toBe('(no date)');
  });
});

describe('truncate', () => {
  it('leaves short text untouched', () => {
    expect(truncate('hello', 10)).toBe('hello');
  });

  it('truncates and adds an ellipsis when over max', () => {
    const out = truncate('hello world', 8);
    expect(out.length).toBe(8);
    expect(out.endsWith('\u2026')).toBe(true);
  });

  it('handles max of 0', () => {
    expect(truncate('hello', 0)).toBe('');
  });

  it('handles max of 1', () => {
    expect(truncate('hello', 1)).toBe('h');
  });
});

describe('formatRecipients', () => {
  it('renders "(none)" for an empty array', () => {
    expect(formatRecipients([])).toBe('(none)');
  });

  it('renders "(none)" for undefined', () => {
    expect(formatRecipients(undefined)).toBe('(none)');
  });

  it('renders a single recipient with name and address', () => {
    expect(formatRecipients([recipient('a@example.com', 'Alice')])).toBe(
      'Alice <a@example.com>',
    );
  });

  it('renders a single recipient with no name as the bare address', () => {
    expect(formatRecipients([recipient('a@example.com')])).toBe('a@example.com');
  });

  it('joins many recipients with a comma', () => {
    const out = formatRecipients([
      recipient('a@example.com', 'Alice'),
      recipient('b@example.com'),
      recipient('c@example.com', 'Carol'),
    ]);
    expect(out).toBe('Alice <a@example.com>, b@example.com, Carol <c@example.com>');
  });
});

describe('renderFolders', () => {
  it('renders an explicit empty-state line for an empty array', () => {
    expect(renderFolders([])).toBe('(no folders)');
  });

  it('renders folder display name, id, and counts', () => {
    const folders: MailFolder[] = [
      {
        id: 'folder-1',
        displayName: 'Inbox',
        unreadItemCount: 5,
        totalItemCount: 42,
      },
    ];
    const out = renderFolders(folders);
    expect(out).toContain('Inbox');
    expect(out).toContain('folder-1');
    expect(out).toContain('5 unread');
    expect(out).toContain('42 total');
  });

  it('handles a folder with no counts at all', () => {
    const folders: MailFolder[] = [{ id: 'folder-2', displayName: 'Custom' }];
    expect(() => renderFolders(folders)).not.toThrow();
    expect(renderFolders(folders)).toContain('Custom');
  });
});

describe('renderAttachments', () => {
  it('renders an explicit empty-state line for an empty array', () => {
    expect(renderAttachments([])).toBe('(no attachments)');
  });

  it('renders a fileAttachment with size and kind', () => {
    const attachments: Attachment[] = [
      {
        id: 'att-1',
        name: 'report.pdf',
        size: 2048,
        '@odata.type': '#microsoft.graph.fileAttachment',
      },
    ];
    const out = renderAttachments(attachments);
    expect(out).toContain('report.pdf');
    expect(out).toContain('file');
    expect(out).toContain('att-1');
  });

  it('recognizes an itemAttachment discriminator', () => {
    const attachments: Attachment[] = [
      {
        id: 'att-2',
        name: 'Meeting invite',
        '@odata.type': '#microsoft.graph.itemAttachment',
      },
    ];
    expect(renderAttachments(attachments)).toContain('item');
  });

  it('recognizes a referenceAttachment discriminator', () => {
    const attachments: Attachment[] = [
      {
        id: 'att-3',
        name: 'Shared doc',
        '@odata.type': '#microsoft.graph.referenceAttachment',
      },
    ];
    expect(renderAttachments(attachments)).toContain('reference');
  });

  it('tolerates the non-#-prefixed odata.type form seen in S9 examples', () => {
    const attachments: Attachment[] = [
      {
        id: 'att-4',
        name: 'x.pdf',
        '@odata.type': 'microsoft.graph.fileAttachment',
      },
    ];
    expect(renderAttachments(attachments)).toContain('file');
  });

  it('marks inline attachments', () => {
    const attachments: Attachment[] = [
      {
        id: 'att-5',
        name: 'inline.png',
        isInline: true,
        '@odata.type': '#microsoft.graph.fileAttachment',
      },
    ];
    expect(renderAttachments(attachments)).toContain('[inline]');
  });

  it('handles missing size gracefully', () => {
    const attachments: Attachment[] = [{ id: 'att-6', name: 'unsized.bin' }];
    expect(() => renderAttachments(attachments)).not.toThrow();
    expect(renderAttachments(attachments)).toContain('(unknown size)');
  });
});

describe('renderProfile', () => {
  it('renders display name, mail, and id', () => {
    const profile: UserProfile = {
      id: 'user-1',
      displayName: 'Adele Vance',
      mail: 'adelev@contoso.com',
    };
    const out = renderProfile(profile);
    expect(out).toContain('Adele Vance');
    expect(out).toContain('adelev@contoso.com');
    expect(out).toContain('user-1');
  });

  it('falls back to userPrincipalName when mail is null', () => {
    const profile: UserProfile = {
      id: 'user-2',
      mail: null,
      userPrincipalName: 'adelev@contoso.onmicrosoft.com',
    };
    expect(renderProfile(profile)).toContain('adelev@contoso.onmicrosoft.com');
  });

  it('handles a profile with only an id', () => {
    const profile: UserProfile = { id: 'user-3' };
    expect(() => renderProfile(profile)).not.toThrow();
    const out = renderProfile(profile);
    expect(out).toContain('user-3');
    expect(out).toContain('(unknown)');
  });
});
