import { describe, it, expect, vi, beforeEach } from 'vitest';
import { ImapService, EXPORT_FETCH_CHUNK } from '../src/services/imap-service.js';
import type { ImapAccount } from '../src/types/index.js';

// Large hosted mailboxes reject a single FETCH over tens of thousands of UIDs
// ("Command failed" on a 30k-message Outlook.com INBOX). exportFolderRows must
// fetch in bounded chunks and still emit every row exactly once, in order.
const fetchCalls: number[][] = [];
const TOTAL = EXPORT_FETCH_CHUNK * 2 + 37;

vi.mock('imapflow', () => ({
  ImapFlow: class {
    usable = true;
    constructor(_options: any) {}
    connect() { return Promise.resolve(); }
    logout() { return Promise.resolve(); }
    on() {}
    getMailboxLock() { return Promise.resolve({ release() {} }); }
    search() { return Promise.resolve(Array.from({ length: TOTAL }, (_, i) => i + 1)); }
    async *fetch(uids: number[]) {
      fetchCalls.push([...uids]);
      for (const uid of uids) {
        yield {
          uid,
          flags: new Set(uid % 2 ? ['\\Seen'] : []),
          internalDate: new Date('2026-01-01T00:00:00Z'),
          envelope: { from: [{ address: `s${uid}@example.com`, name: 'S' }], to: [], subject: `m${uid}` },
          size: 10,
        };
      }
    }
  },
}));

describe('ImapService.exportFolderRows chunking', () => {
  let service: ImapService;
  const account: ImapAccount = { id: 'a', name: 'A', host: 'imap.example.com', port: 993, user: 'u', password: 'p', tls: true };

  beforeEach(() => {
    fetchCalls.length = 0;
    service = new ImapService();
  });

  it('splits the UID set into bounded FETCH commands and streams every row once', async () => {
    await service.connect(account);
    const seen: number[] = [];
    const count = await service.exportFolderRows('a', 'INBOX', {}, row => { seen.push(row.uid); });

    expect(count).toBe(TOTAL);
    expect(seen).toEqual(Array.from({ length: TOTAL }, (_, i) => i + 1));
    expect(fetchCalls).toHaveLength(3);
    expect(fetchCalls.map(c => c.length)).toEqual([EXPORT_FETCH_CHUNK, EXPORT_FETCH_CHUNK, 37]);
  });

  it('honours limit by keeping only the newest UIDs before chunking', async () => {
    await service.connect(account);
    const seen: number[] = [];
    await service.exportFolderRows('a', 'INBOX', { limit: 5 }, row => { seen.push(row.uid); });
    expect(seen).toEqual([TOTAL - 4, TOTAL - 3, TOTAL - 2, TOTAL - 1, TOTAL]);
    expect(fetchCalls).toHaveLength(1);
  });
});
