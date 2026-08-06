/**
 * renderMessageList / renderMessage coverage, split out of output.test.ts
 * to stay under the 350-LOC file-size gate (extraction, not trimming —
 * see scripts/check-file-size.sh).
 */
import { describe, it, expect } from 'vitest';
import { renderMessageList, renderMessage } from './output.js';
import type { Message, Recipient } from './types.js';

function recipient(address: string, name?: string): Recipient {
  return { emailAddress: name ? { address, name } : { address } };
}

describe('renderMessageList', () => {
  it('renders an explicit "no messages" line for an empty array, never a blank string', () => {
    const out = renderMessageList([]);
    expect(out).toBe('(no messages)');
    expect(out.trim().length).toBeGreaterThan(0);
  });

  it('does not crash on a message with missing subject and null from', () => {
    const messages: Message[] = [
      {
        id: 'msg-1',
        subject: null,
        from: null,
        sender: null,
        receivedDateTime: '2024-01-01T00:00:00Z',
      },
    ];
    expect(() => renderMessageList(messages)).not.toThrow();
    const out = renderMessageList(messages);
    expect(out).toContain('(no subject)');
    expect(out).toContain('(unknown sender)');
    expect(out).toContain('msg-1');
  });

  it('does not crash on a message with subject entirely absent (undefined)', () => {
    const messages: Message[] = [{ id: 'msg-2' }];
    expect(() => renderMessageList(messages)).not.toThrow();
    expect(renderMessageList(messages)).toContain('(no subject)');
  });

  it('marks unread messages with a leading marker', () => {
    const messages: Message[] = [
      { id: 'm-unread', subject: 'Hi', isRead: false },
    ];
    const lines = renderMessageList(messages).split('\n');
    const row = lines[1];
    expect(row).toBeDefined();
    expect(row?.startsWith('*')).toBe(true);
  });

  it('does not mark read messages', () => {
    const messages: Message[] = [{ id: 'm-read', subject: 'Hi', isRead: true }];
    const lines = renderMessageList(messages).split('\n');
    const row = lines[1];
    expect(row).toBeDefined();
    expect(row?.startsWith('*')).toBe(false);
  });

  it('marks messages that have attachments', () => {
    const messages: Message[] = [
      { id: 'm-att', subject: 'Files', hasAttachments: true },
    ];
    expect(renderMessageList(messages)).toContain('@');
  });

  it('renders multiple rows with a header line first', () => {
    const messages: Message[] = [
      { id: 'm1', subject: 'First' },
      { id: 'm2', subject: 'Second' },
    ];
    const lines = renderMessageList(messages).split('\n');
    expect(lines.length).toBe(3);
    expect(lines[0]).toMatch(/DATE/);
    expect(lines[0]).toMatch(/FROM/);
    expect(lines[0]).toMatch(/SUBJECT/);
  });

  it('truncates a very long subject rather than breaking column alignment', () => {
    const longSubject = 'x'.repeat(200);
    const messages: Message[] = [{ id: 'm-long', subject: longSubject }];
    const out = renderMessageList(messages);
    expect(out).not.toContain(longSubject);
  });
});

describe('renderMessage', () => {
  it('renders full headers then body', () => {
    const message: Message = {
      id: 'msg-1',
      subject: 'Meet for lunch?',
      from: recipient('frannis@contoso.com', 'Frannis'),
      toRecipients: [recipient('me@contoso.com')],
      receivedDateTime: '2024-06-01T12:00:00Z',
      importance: 'high',
      isRead: true,
      hasAttachments: false,
      body: { contentType: 'text', content: 'The new cafeteria is open.' },
    };
    const out = renderMessage(message);
    expect(out).toContain('Subject: Meet for lunch?');
    expect(out).toContain('Frannis <frannis@contoso.com>');
    expect(out).toContain('Importance: high');
    expect(out).toContain('The new cafeteria is open.');
  });

  it('strips HTML to plain text by default', () => {
    const message: Message = {
      id: 'msg-html',
      subject: 'HTML body',
      body: { contentType: 'html', content: '<p>Hello <b>world</b></p>' },
    };
    const out = renderMessage(message);
    expect(out).not.toMatch(/<[a-zA-Z/][^>]*>/);
    expect(out).toContain('Hello');
    expect(out).toContain('world');
  });

  it('preserves raw HTML when opts.html is true', () => {
    const message: Message = {
      id: 'msg-html-2',
      subject: 'HTML body',
      body: { contentType: 'html', content: '<p>Hello <b>world</b></p>' },
    };
    const out = renderMessage(message, { html: true });
    expect(out).toContain('<p>Hello <b>world</b></p>');
  });

  it('falls back to bodyPreview when body is absent', () => {
    const message: Message = {
      id: 'msg-preview',
      subject: 'No full body',
      bodyPreview: 'Just a preview.',
    };
    const out = renderMessage(message);
    expect(out).toContain('Just a preview.');
  });

  it('renders "(no body)" when neither body nor bodyPreview exist', () => {
    const message: Message = { id: 'msg-empty', subject: 'Empty' };
    expect(renderMessage(message)).toContain('(no body)');
  });

  it('does not crash with null subject and null from/sender', () => {
    const message: Message = {
      id: 'msg-null',
      subject: null,
      from: null,
      sender: null,
    };
    expect(() => renderMessage(message)).not.toThrow();
    const out = renderMessage(message);
    expect(out).toContain('(no subject)');
    expect(out).toContain('(unknown sender)');
  });

  it('omits the Cc line when there are no cc recipients', () => {
    const message: Message = { id: 'msg-no-cc', subject: 'x' };
    expect(renderMessage(message)).not.toContain('Cc:');
  });

  it('includes the Cc line when cc recipients are present', () => {
    const message: Message = {
      id: 'msg-cc',
      subject: 'x',
      ccRecipients: [recipient('cc@example.com')],
    };
    expect(renderMessage(message)).toContain('Cc: cc@example.com');
  });
});
