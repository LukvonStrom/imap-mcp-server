import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { promises as fsp } from 'fs';
import { join } from 'path';
import { tmpdir } from 'os';
import { exportTools } from '../src/tools/export-tools.js';
import type { MessageExportRow } from '../src/types/index.js';

let handler: Function;

const mockServer = {
  registerTool: vi.fn((name: string, _schema: any, fn: Function) => {
    if (name === 'imap_export_messages') handler = fn;
  }),
};

const folders = [
  { name: 'INBOX', delimiter: '/', attributes: [] },
  { name: 'Junk Email', delimiter: '/', attributes: [], specialUse: '\\Junk' },
  { name: 'Deleted Items', delimiter: '/', attributes: [], specialUse: '\\Trash' },
  { name: 'Drafts', delimiter: '/', attributes: [], specialUse: '\\Drafts' },
  { name: 'Newsletters', delimiter: '/', attributes: [] },
];

function row(partial: Partial<MessageExportRow> & { folder: string; uid: number; from: string }): MessageExportRow {
  const domain = partial.from.split('@')[1];
  return {
    date: '2026-08-01T10:00:00.000Z',
    fromName: '',
    fromDomain: domain,
    replyTo: '',
    to: ['me@outlook.com'],
    ccCount: 0,
    subject: 'Subject, with "quotes"',
    seen: true,
    flagged: false,
    answered: false,
    size: 1000,
    hasAttachments: false,
    listId: '',
    hasListUnsubscribe: false,
    precedence: '',
    autoSubmitted: '',
    messageId: `<${partial.uid}@${domain}>`,
    inReplyTo: '',
    ...partial,
  };
}

// Six unread newsletter messages in INBOX, three receipts filed to Newsletters, one junk.
const data: Record<string, MessageExportRow[]> = {
  INBOX: [
    ...[1, 2, 3, 4, 5, 6].map(uid => row({ folder: 'INBOX', uid, from: 'news@deals.example', seen: false, listId: '<deals.example>', hasListUnsubscribe: true })),
    row({ folder: 'INBOX', uid: 7, from: 'boss@work.example', seen: true }),
  ],
  Newsletters: [8, 9, 10].map(uid => row({ folder: 'Newsletters', uid, from: 'weekly@paper.example', seen: true, listId: '<weekly.paper.example>' })),
  'Junk Email': [row({ folder: 'Junk Email', uid: 11, from: 'spam@bad.example', seen: false })],
  'Deleted Items': [row({ folder: 'Deleted Items', uid: 12, from: 'old@trash.example' })],
  Drafts: [row({ folder: 'Drafts', uid: 13, from: 'me@outlook.com' })],
};

const mockImapService = {
  listFolders: vi.fn(async () => folders),
  exportFolderRows: vi.fn(async (_acc: string, folder: string, opts: any, onRow: Function) => {
    let rows = data[folder] || [];
    if (opts.limit) rows = rows.slice(-opts.limit);
    for (const r of rows) await onRow(r);
    return rows.length;
  }),
};

const mockAccountManager = {
  resolveAccountId: (id: string) => id,
  getAccount: () => ({ id: 'acc1', name: 'My Outlook' }),
};

describe('imap_export_messages', () => {
  let dir: string;
  let prevDownloadDir: string | undefined;

  beforeEach(async () => {
    vi.clearAllMocks();
    prevDownloadDir = process.env.IMAP_DOWNLOAD_DIR;
    dir = await fsp.mkdtemp(join(tmpdir(), 'imap-mcp-export-'));
    process.env.IMAP_DOWNLOAD_DIR = dir;
    exportTools(mockServer as any, mockImapService as any, mockAccountManager as any);
  });

  afterEach(async () => {
    if (prevDownloadDir === undefined) delete process.env.IMAP_DOWNLOAD_DIR;
    else process.env.IMAP_DOWNLOAD_DIR = prevDownloadDir;
    await fsp.rm(dir, { recursive: true, force: true });
  });

  it('writes JSONL under the download dir, skips Trash/Drafts by default, and derives rule candidates', async () => {
    const result = await handler({ accountId: 'acc1', includeJunk: true, includeTrash: false, limitPerFolder: 5000, format: 'jsonl', summaryTop: 30, minMessagesForRule: 3 });
    const payload = JSON.parse(result.content[0].text);

    expect(payload.success).toBe(true);
    expect(payload.path.startsWith(join(dir, 'exports'))).toBe(true);
    expect(payload.rowCount).toBe(11);
    expect(payload.folders).toEqual({ INBOX: 7, 'Junk Email': 1, Newsletters: 3 });
    expect(Object.keys(payload.folders)).not.toContain('Deleted Items');
    expect(Object.keys(payload.folders)).not.toContain('Drafts');

    const lines = (await fsp.readFile(payload.path, 'utf8')).trim().split('\n');
    expect(lines).toHaveLength(11);
    expect(JSON.parse(lines[0])).toMatchObject({ folder: 'INBOX', uid: 1, fromDomain: 'deals.example' });

    expect(payload.summary.topDomains[0]).toMatchObject({ value: 'deals.example', messages: 6, unreadPercent: 100 });
    const byValue = Object.fromEntries(payload.ruleCandidates.map((c: any) => [c.match.value, c]));
    expect(byValue['deals.example'].suggestion).toMatch(/newsletter/i);
    expect(byValue['paper.example']).toMatchObject({ dominantFolder: 'Newsletters', dominantFolderPercent: 100 });
    expect(byValue['paper.example'].suggestion).toContain('move to Newsletters');
    // Actively read personal mail must not become a rule candidate.
    expect(byValue['work.example']).toBeUndefined();
  });

  it('writes CSV with a header, quoting, and semicolon-joined recipients, honouring an explicit folder list and filename', async () => {
    const result = await handler({ accountId: 'acc1', folders: ['INBOX'], includeJunk: true, includeTrash: false, limitPerFolder: 2, format: 'csv', filename: '../../escape.csv', summaryTop: 5, minMessagesForRule: 5 });
    const payload = JSON.parse(result.content[0].text);

    expect(payload.path).toBe(join(dir, 'exports', 'escape.csv'));
    expect(payload.rowCount).toBe(2);
    const text = await fsp.readFile(payload.path, 'utf8');
    const [header, first] = text.split('\n');
    expect(header.startsWith('folder,uid,date,from,')).toBe(true);
    expect(first).toContain('"Subject, with ""quotes"""');
    expect(first).toContain('me@outlook.com');
    expect(mockImapService.exportFolderRows).toHaveBeenCalledTimes(1);
  });

  it('rejects unknown folders and bad dates without writing anything', async () => {
    await expect(handler({ accountId: 'acc1', folders: ['Nope'], includeJunk: true, includeTrash: false, limitPerFolder: 10, format: 'jsonl', summaryTop: 5, minMessagesForRule: 5 }))
      .rejects.toThrow(/Unknown or non-selectable folder/);
    await expect(handler({ accountId: 'acc1', since: 'last week', includeJunk: true, includeTrash: false, limitPerFolder: 10, format: 'jsonl', summaryTop: 5, minMessagesForRule: 5 }))
      .rejects.toThrow(/YYYY-MM-DD/);
    await expect(fsp.readdir(join(dir, 'exports'))).rejects.toThrow();
  });
});
