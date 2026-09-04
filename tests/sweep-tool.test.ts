import { describe, it, expect, vi, beforeEach } from 'vitest';
import { sweepTools, sweepCutoff, SWEEP_UID_LIST_CAP } from '../src/tools/sweep-tools.js';
import { classifySpecialFolder } from '../src/utils/search-folders.js';

// Capture the handler and config registered for imap_sweep.
let sweepHandler: Function;
let sweepConfig: any;

const mockServer = {
  registerTool: vi.fn((name: string, config: any, handler: Function) => {
    if (name === 'imap_sweep') {
      sweepHandler = handler;
      sweepConfig = config;
    }
  }),
};

const mockImapService = {
  listFolders: vi.fn(),
  searchEmails: vi.fn(),
  moveUids: vi.fn(),
  markAsRead: vi.fn(),
  bulkDelete: vi.fn(),
  createFolder: vi.fn(),
};

const mockAccountManager = { resolveAccountId: (id?: string) => id ?? 'acc1' };

const DAY = 24 * 60 * 60 * 1000;
const daysAgo = (n: number) => new Date(Date.now() - n * DAY);
const msg = (uid: number, from: string, ageDays: number, flags: string[] = []) => ({
  uid,
  date: daysAgo(ageDays),
  from,
  to: ['me@example.com'],
  subject: `msg ${uid}`,
  messageId: `<${uid}@x>`,
  flags,
  customKeywords: [],
});

const folders = [
  { name: 'INBOX', delimiter: '/', attributes: [] },
  { name: 'Archive', delimiter: '/', attributes: [], children: [
    { name: 'Archive/Newsletters', delimiter: '/', attributes: [] },
  ] },
  { name: 'Deleted', delimiter: '/', attributes: [], specialUse: '\\Trash' },
  { name: 'Junk Email', delimiter: '/', attributes: [], specialUse: '\\Junk' },
];

const parse = (result: any) => JSON.parse(result.content[0].text);

/** Route each FROM search to a fixed result set, keyed by the `from` criterion. */
function searchBySender(map: Record<string, any[]>) {
  mockImapService.searchEmails.mockImplementation(async (_acc: string, _folder: string, criteria: any) => {
    return map[criteria.from] ?? [];
  });
}

describe('imap_sweep', () => {
  beforeEach(() => {
    vi.clearAllMocks();
    sweepTools(mockServer as any, mockImapService as any, mockAccountManager as any);
    mockImapService.listFolders.mockResolvedValue(folders);
    mockImapService.searchEmails.mockResolvedValue([]);
    mockImapService.moveUids.mockResolvedValue({ moved: 0, failed: 0, errors: [] });
    mockImapService.markAsRead.mockImplementation(async (_a: string, _f: string, uids: number[]) => ({
      success: true, marked: [...uids], failed: [],
    }));
    mockImapService.bulkDelete.mockResolvedValue({ deleted: 0, failed: 0, errors: [] });
    mockImapService.createFolder.mockResolvedValue({ path: 'x', created: true, alreadyExisted: false });
  });

  it('is registered with a title and destructive-but-idempotent annotations', () => {
    expect(sweepHandler).toBeDefined();
    expect(sweepConfig.title).toBeTruthy();
    expect(sweepConfig.annotations).toEqual({
      readOnlyHint: false,
      destructiveHint: true,
      idempotentHint: true,
      openWorldHint: false,
    });
    expect(sweepConfig.inputSchema.dryRun.parse(undefined)).toBe(true);
    expect(sweepConfig.inputSchema.action.parse(undefined)).toBe('move');
    expect(sweepConfig.inputSchema.chunkSize.parse(undefined)).toBe(200);
    expect(() => sweepConfig.inputSchema.senders.parse([])).toThrow();
    expect(sweepConfig.inputSchema.senders.parse('["a@x.com","b@y.com"]')).toEqual(['a@x.com', 'b@y.com']);
  });

  describe('cutoff', () => {
    it('is date-only, computed in UTC as today minus N days at midnight', () => {
      const now = new Date('2026-09-04T15:45:12Z');
      const cutoff = sweepCutoff(7, now);
      expect(cutoff.toISOString()).toBe('2026-08-28T00:00:00.000Z');
      expect(sweepCutoff(0, now).toISOString()).toBe('2026-09-04T00:00:00.000Z');
      // Crosses a month boundary correctly.
      expect(sweepCutoff(10, new Date('2026-03-05T01:00:00Z')).toISOString()).toBe('2026-02-23T00:00:00.000Z');
    });

    it('is passed to the search as a `before` Date and echoed as cutoffDate', async () => {
      const result = parse(await sweepHandler({ senders: ['news@x.com'], olderThanDays: 7, targetFolder: 'Archive' }));
      expect(mockImapService.searchEmails).toHaveBeenCalledOnce();
      const criteria = mockImapService.searchEmails.mock.calls[0][2];
      expect(criteria.from).toBe('news@x.com');
      expect(criteria.before).toBeInstanceOf(Date);
      expect(criteria.before.toISOString()).toBe(`${result.cutoffDate}T00:00:00.000Z`);
      expect(result.cutoffDate).toBe(sweepCutoff(7).toISOString().slice(0, 10));
      expect(criteria.seen).toBeUndefined();
    });

    it('maps onlyUnread / onlySeen onto the seen criterion', async () => {
      await sweepHandler({ senders: ['a@x.com'], olderThanDays: 1, onlyUnread: true, targetFolder: 'Archive' });
      expect(mockImapService.searchEmails.mock.calls[0][2].seen).toBe(false);
      vi.clearAllMocks();
      mockImapService.listFolders.mockResolvedValue(folders);
      await sweepHandler({ senders: ['a@x.com'], olderThanDays: 1, onlySeen: true, targetFolder: 'Archive' });
      expect(mockImapService.searchEmails.mock.calls[0][2].seen).toBe(true);
    });
  });

  describe('dry run (the default)', () => {
    it('returns the per-sender plan and never calls a mutating service method', async () => {
      searchBySender({
        'news@x.com': [msg(10, 'news@x.com', 30), msg(11, 'news@x.com', 20), msg(12, 'news@x.com', 9)],
        '@promo.example': [msg(20, 'deals@promo.example', 12)],
      });

      const result = parse(await sweepHandler({
        accountId: 'acc1',
        senders: ['news@x.com', '@promo.example'],
        olderThanDays: 7,
        targetFolder: 'Archive/Newsletters',
      }));

      expect(result.success).toBe(true);
      expect(result.dryRun).toBe(true);
      expect(result.folder).toBe('INBOX');
      expect(result.action).toBe('move');
      expect(result.targetFolder).toBe('Archive/Newsletters');
      expect(result.targetFolderExists).toBe(true);
      expect(result.olderThanDays).toBe(7);
      expect(result.keepLatest).toBe(0);
      expect(result.perSender).toHaveLength(2);
      const [news, promo] = result.perSender;
      expect(news).toMatchObject({ sender: 'news@x.com', matched: 3, qualifying: 3, keptUids: [], uids: [10, 11, 12], truncated: false });
      expect(new Date(news.oldest).getTime()).toBeLessThan(new Date(news.newest).getTime());
      expect(promo).toMatchObject({ sender: '@promo.example', matched: 1, qualifying: 1, uids: [20] });
      expect(result.totalMatched).toBe(4);
      expect(result.totalPlanned).toBe(4);
      expect(result.totalActioned).toBe(0);
      expect(result.kept).toBe(0);
      expect(result.uids).toEqual([10, 11, 12, 20]);
      expect(result.errors).toEqual([]);
      expect(result.message).toMatch(/dry run/i);

      expect(mockImapService.moveUids).not.toHaveBeenCalled();
      expect(mockImapService.bulkDelete).not.toHaveBeenCalled();
      expect(mockImapService.markAsRead).not.toHaveBeenCalled();
      expect(mockImapService.createFolder).not.toHaveBeenCalled();
    });

    it('also stays read-only for action delete with confirmDelete', async () => {
      searchBySender({ 'a@x.com': [msg(1, 'a@x.com', 40)] });
      const result = parse(await sweepHandler({ senders: ['a@x.com'], olderThanDays: 30, action: 'delete', confirmDelete: true }));
      expect(result.dryRun).toBe(true);
      expect(result.totalPlanned).toBe(1);
      expect(mockImapService.bulkDelete).not.toHaveBeenCalled();
    });

    it('caps echoed UID lists and flags truncation', async () => {
      const many = Array.from({ length: SWEEP_UID_LIST_CAP + 25 }, (_, i) => msg(1000 + i, 'bulk@x.com', 60 + i));
      searchBySender({ 'bulk@x.com': many });
      const result = parse(await sweepHandler({ senders: ['bulk@x.com'], olderThanDays: 30, targetFolder: 'Archive' }));
      expect(result.perSender[0].qualifying).toBe(many.length);
      expect(result.perSender[0].uids).toHaveLength(SWEEP_UID_LIST_CAP);
      expect(result.perSender[0].truncated).toBe(true);
      expect(result.uids).toHaveLength(SWEEP_UID_LIST_CAP);
      expect(result.truncated).toBe(true);
      expect(result.totalPlanned).toBe(many.length);
    });

    it('reports a per-sender search failure without aborting the other senders', async () => {
      mockImapService.searchEmails.mockImplementation(async (_a: string, _f: string, c: any) => {
        if (c.from === 'bad@x.com') throw new Error('SEARCH failed');
        return [msg(5, 'ok@x.com', 10)];
      });
      const result = parse(await sweepHandler({ senders: ['bad@x.com', 'ok@x.com'], olderThanDays: 1, targetFolder: 'Archive' }));
      expect(result.success).toBe(false);
      expect(result.errors).toEqual(['Search for sender "bad@x.com" failed: SEARCH failed']);
      expect(result.perSender[0]).toMatchObject({ sender: 'bad@x.com', matched: 0, error: 'SEARCH failed' });
      expect(result.perSender[1]).toMatchObject({ sender: 'ok@x.com', qualifying: 1 });
      expect(result.totalPlanned).toBe(1);
    });
  });

  describe('keepLatest', () => {
    it('keeps the newest N per sender (by date, then UID) out of the action set', async () => {
      searchBySender({
        'news@x.com': [msg(10, 'news@x.com', 30), msg(13, 'news@x.com', 8), msg(11, 'news@x.com', 20), msg(12, 'news@x.com', 8)],
        'other@y.com': [msg(50, 'other@y.com', 15), msg(51, 'other@y.com', 14)],
      });

      const result = parse(await sweepHandler({ senders: ['news@x.com', 'other@y.com'], olderThanDays: 7, keepLatest: 2, targetFolder: 'Archive' }));

      const [news, other] = result.perSender;
      // 12 and 13 share a date; the higher UID (13) is the newer one.
      expect(news.keptUids).toEqual([13, 12]);
      expect(news.uids).toEqual([10, 11]);
      expect(news.matched).toBe(4);
      expect(news.qualifying).toBe(2);
      expect(other.keptUids).toEqual([51, 50]);
      expect(other.uids).toEqual([]);
      expect(other.qualifying).toBe(0);
      expect(result.kept).toBe(4);
      expect(result.totalMatched).toBe(6);
      expect(result.uids).toEqual([10, 11]);
    });
  });

  describe('union of senders', () => {
    it('processes a message matched by two senders exactly once', async () => {
      const shared = msg(7, 'news@corp.example', 20);
      searchBySender({
        'news@corp.example': [shared, msg(8, 'news@corp.example', 15)],
        'corp.example': [shared, msg(9, 'hr@corp.example', 15)],
      });

      const result = parse(await sweepHandler({ senders: ['news@corp.example', 'corp.example'], olderThanDays: 7, targetFolder: 'Archive', dryRun: false }));

      expect(result.totalMatched).toBe(3);
      expect(result.totalPlanned).toBe(3);
      expect(result.uids).toEqual([7, 8, 9]);
      expect(mockImapService.moveUids).toHaveBeenCalledOnce();
      expect(mockImapService.moveUids.mock.calls[0][2]).toEqual([7, 8, 9]);
    });

    it('lets a keepLatest hold for one sender win over a match by another sender', async () => {
      const newest = msg(30, 'news@corp.example', 8);
      searchBySender({
        'news@corp.example': [msg(29, 'news@corp.example', 12), newest],
        'corp.example': [newest, msg(31, 'hr@corp.example', 7.2)],
      });

      const result = parse(await sweepHandler({ senders: ['news@corp.example', 'corp.example'], olderThanDays: 7, keepLatest: 1, targetFolder: 'Archive' }));

      expect(result.perSender[0].keptUids).toEqual([30]);
      expect(result.perSender[1].keptUids).toEqual([31]);
      expect(result.uids).toEqual([29]);
      expect(result.kept).toBe(2);
    });
  });

  describe('guards', () => {
    it.each([
      ['empty senders array', { senders: [], olderThanDays: 1 }, /without senders/],
      ['blank senders', { senders: ['  ', ''], olderThanDays: 1 }, /without senders/],
      ['missing senders', { olderThanDays: 1 }, /without senders/],
      ['negative olderThanDays', { senders: ['a@x.com'], olderThanDays: -1 }, /olderThanDays/],
      ['non-integer olderThanDays', { senders: ['a@x.com'], olderThanDays: 1.5 }, /olderThanDays/],
      ['onlyUnread + onlySeen', { senders: ['a@x.com'], olderThanDays: 1, onlyUnread: true, onlySeen: true }, /mutually exclusive/],
      ['delete without confirmDelete', { senders: ['a@x.com'], olderThanDays: 1, action: 'delete', dryRun: false }, /confirmDelete/],
      ['delete with confirmDelete false', { senders: ['a@x.com'], olderThanDays: 1, action: 'delete', confirmDelete: false, dryRun: false }, /confirmDelete/],
      ['move without targetFolder', { senders: ['a@x.com'], olderThanDays: 1, action: 'move', dryRun: false }, /targetFolder/],
      ['moveAndMarkRead without targetFolder', { senders: ['a@x.com'], olderThanDays: 1, action: 'moveAndMarkRead', dryRun: false }, /targetFolder/],
      ['target equals source', { senders: ['a@x.com'], olderThanDays: 1, folder: 'INBOX', targetFolder: 'INBOX', dryRun: false }, /differ/],
    ])('refuses %s before touching the mailbox', async (_label, input, pattern) => {
      const result = parse(await sweepHandler(input));
      expect(result.success).toBe(false);
      expect(result.error).toMatch(pattern);
      expect(result.totalActioned).toBe(0);
      expect(mockImapService.listFolders).not.toHaveBeenCalled();
      expect(mockImapService.searchEmails).not.toHaveBeenCalled();
      expect(mockImapService.moveUids).not.toHaveBeenCalled();
      expect(mockImapService.markAsRead).not.toHaveBeenCalled();
      expect(mockImapService.bulkDelete).not.toHaveBeenCalled();
    });

    it('refuses a Trash source (by SPECIAL-USE flag) unless allowSpecialFolders is set', async () => {
      const refused = parse(await sweepHandler({ senders: ['a@x.com'], olderThanDays: 1, folder: 'Deleted', targetFolder: 'Archive' }));
      expect(refused.success).toBe(false);
      expect(refused.error).toMatch(/Trash/);
      expect(mockImapService.searchEmails).not.toHaveBeenCalled();

      const allowed = parse(await sweepHandler({ senders: ['a@x.com'], olderThanDays: 1, folder: 'Deleted', targetFolder: 'Archive', allowSpecialFolders: true }));
      expect(allowed.success).toBe(true);
      expect(mockImapService.searchEmails).toHaveBeenCalledWith('acc1', 'Deleted', expect.anything());
    });

    it('refuses a Junk source by name when the folder list carries no SPECIAL-USE flag', async () => {
      mockImapService.listFolders.mockResolvedValue([{ name: 'INBOX', delimiter: '/', attributes: [] }]);
      const result = parse(await sweepHandler({ senders: ['a@x.com'], olderThanDays: 1, folder: 'INBOX/Spam', targetFolder: 'Archive' }));
      expect(result.success).toBe(false);
      expect(result.error).toMatch(/Junk\/Spam/);
      expect(mockImapService.searchEmails).not.toHaveBeenCalled();
    });

    it('classifySpecialFolder prefers SPECIAL-USE over the name and falls back to name heuristics', () => {
      expect(classifySpecialFolder('Deleted', folders)).toBe('trash');
      expect(classifySpecialFolder('Junk Email', folders)).toBe('spam');
      expect(classifySpecialFolder('INBOX', folders)).toBeNull();
      expect(classifySpecialFolder('Archive/Newsletters', folders)).toBeNull();
      expect(classifySpecialFolder('[Gmail]/Trash')).toBe('trash');
      expect(classifySpecialFolder('Bulk Mail')).toBe('spam');
      expect(classifySpecialFolder('Projects/Trash Talk')).toBeNull();
    });

    it('refuses a real move into a missing target unless createFolder is set, and only creates when there is work', async () => {
      searchBySender({ 'a@x.com': [msg(1, 'a@x.com', 10)] });

      const refused = parse(await sweepHandler({ senders: ['a@x.com'], olderThanDays: 1, targetFolder: 'Nope', dryRun: false }));
      expect(refused.success).toBe(false);
      expect(refused.error).toMatch(/does not exist/);
      expect(mockImapService.createFolder).not.toHaveBeenCalled();
      expect(mockImapService.moveUids).not.toHaveBeenCalled();

      mockImapService.moveUids.mockResolvedValueOnce({ moved: 1, failed: 0, errors: [] });
      const created = parse(await sweepHandler({ senders: ['a@x.com'], olderThanDays: 1, targetFolder: 'Nope', createFolder: true, dryRun: false }));
      expect(created.success).toBe(true);
      expect(created.targetFolderCreated).toBe(true);
      expect(mockImapService.createFolder).toHaveBeenCalledWith('acc1', 'Nope');
      expect(mockImapService.moveUids).toHaveBeenCalledWith('acc1', 'INBOX', [1], 'Nope', 200);

      // A dry run only reports that the target is missing.
      vi.clearAllMocks();
      mockImapService.listFolders.mockResolvedValue(folders);
      searchBySender({ 'a@x.com': [msg(1, 'a@x.com', 10)] });
      const dry = parse(await sweepHandler({ senders: ['a@x.com'], olderThanDays: 1, targetFolder: 'Nope', createFolder: true }));
      expect(dry.targetFolderExists).toBe(false);
      expect(mockImapService.createFolder).not.toHaveBeenCalled();
    });
  });

  describe('dryRun:false', () => {
    it('moves the union in chunks via moveUids and reports counts and errors', async () => {
      searchBySender({
        'news@x.com': [msg(1, 'news@x.com', 10), msg(2, 'news@x.com', 11), msg(3, 'news@x.com', 12)],
        'promo@y.com': [msg(4, 'promo@y.com', 13), msg(5, 'promo@y.com', 14)],
      });
      mockImapService.moveUids.mockResolvedValueOnce({ moved: 4, failed: 1, errors: ['Failed to move UIDs 5-5 to Archive/Newsletters: boom'] });

      const result = parse(await sweepHandler({
        accountId: 'acc1',
        folder: 'INBOX',
        senders: ['news@x.com', 'promo@y.com'],
        olderThanDays: 7,
        targetFolder: 'Archive/Newsletters',
        dryRun: false,
        chunkSize: 2,
      }));

      expect(mockImapService.moveUids).toHaveBeenCalledWith('acc1', 'INBOX', [1, 2, 3, 4, 5], 'Archive/Newsletters', 2);
      expect(mockImapService.markAsRead).not.toHaveBeenCalled();
      expect(mockImapService.bulkDelete).not.toHaveBeenCalled();
      expect(result.dryRun).toBe(false);
      expect(result.success).toBe(false);
      expect(result.totalMatched).toBe(5);
      expect(result.totalPlanned).toBe(5);
      expect(result.totalActioned).toBe(4);
      expect(result.failed).toBe(1);
      expect(result.errors).toEqual(['Failed to move UIDs 5-5 to Archive/Newsletters: boom']);
      expect(result.message).toBe('4/5 message(s) moved to "Archive/Newsletters", 1 failed.');
    });

    it('marks read in chunkSize batches for action markRead', async () => {
      searchBySender({ 'a@x.com': [msg(1, 'a@x.com', 3), msg(2, 'a@x.com', 4), msg(3, 'a@x.com', 5)] });

      const result = parse(await sweepHandler({ senders: ['a@x.com'], olderThanDays: 2, action: 'markRead', dryRun: false, chunkSize: 2 }));

      expect(mockImapService.markAsRead).toHaveBeenCalledTimes(2);
      expect(mockImapService.markAsRead.mock.calls[0].slice(1)).toEqual(['INBOX', [1, 2]]);
      expect(mockImapService.markAsRead.mock.calls[1].slice(1)).toEqual(['INBOX', [3]]);
      expect(mockImapService.moveUids).not.toHaveBeenCalled();
      expect(result.success).toBe(true);
      expect(result.totalActioned).toBe(3);
      expect(result.targetFolder).toBeUndefined();
    });

    it('marks read and then moves for action moveAndMarkRead', async () => {
      searchBySender({ 'a@x.com': [msg(1, 'a@x.com', 3), msg(2, 'a@x.com', 4)] });
      const order: string[] = [];
      mockImapService.markAsRead.mockImplementation(async (_a: string, _f: string, uids: number[]) => {
        order.push('mark');
        return { success: true, marked: [...uids], failed: [] };
      });
      mockImapService.moveUids.mockImplementation(async () => {
        order.push('move');
        return { moved: 2, failed: 0, errors: [] };
      });

      const result = parse(await sweepHandler({ senders: ['a@x.com'], olderThanDays: 2, action: 'moveAndMarkRead', targetFolder: 'Archive', dryRun: false }));

      expect(order).toEqual(['mark', 'move']);
      expect(result.markedRead).toBe(2);
      expect(result.totalActioned).toBe(2);
      expect(result.success).toBe(true);
    });

    it('deletes through the Trash-aware bulkDelete path when confirmDelete is true', async () => {
      searchBySender({ 'a@x.com': [msg(1, 'a@x.com', 40), msg(2, 'a@x.com', 41)] });
      mockImapService.bulkDelete.mockResolvedValueOnce({ deleted: 2, failed: 0, errors: [] });

      const result = parse(await sweepHandler({ senders: ['a@x.com'], olderThanDays: 30, action: 'delete', confirmDelete: true, dryRun: false, chunkSize: 50 }));

      expect(mockImapService.bulkDelete).toHaveBeenCalledWith('acc1', 'INBOX', [1, 2], 50);
      expect(mockImapService.moveUids).not.toHaveBeenCalled();
      expect(result.totalActioned).toBe(2);
      expect(result.success).toBe(true);
      expect(result.message).toBe('2/2 message(s) deleted.');
    });

    it('does nothing (and creates no folder) when nothing qualifies', async () => {
      const result = parse(await sweepHandler({ senders: ['a@x.com'], olderThanDays: 1, targetFolder: 'Nope', createFolder: true, dryRun: false }));
      expect(result.success).toBe(true);
      expect(result.totalActioned).toBe(0);
      expect(result.message).toMatch(/nothing qualifies/i);
      expect(mockImapService.createFolder).not.toHaveBeenCalled();
      expect(mockImapService.moveUids).not.toHaveBeenCalled();
    });
  });
});
