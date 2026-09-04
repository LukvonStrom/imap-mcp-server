import { describe, it, expect, vi, beforeEach } from 'vitest';
import { ImapService } from '../src/services/imap-service.js';
import type { ImapAccount } from '../src/types/index.js';

// Every ImapFlow the service builds is recorded here so the test can see
// reconnects (a new instance) and the exact UID MOVE commands issued.
// `vi.mock` is hoisted above the imports, so the shared state it closes over
// has to be hoisted too.
const shared = vi.hoisted(() => ({
  instances: [] as unknown[],
  moveCalls: [] as Array<{ range: string; target: string; options: any }>,
  failOn: new Set<string>(),
  releases: 0,
}));

vi.mock('imapflow', () => ({
  ImapFlow: class {
    usable = true;
    constructor(public options: any) {
      shared.instances.push(this);
    }
    connect() { return Promise.resolve(); }
    logout() { return Promise.resolve(); }
    on() {}
    async getMailboxLock(_path: string) {
      return { release: () => { shared.releases++; } };
    }
    async messageMove(range: string, target: string, options: any) {
      shared.moveCalls.push({ range, target, options });
      if (shared.failOn.has(range)) {
        throw new Error(`MOVE rejected for ${range}`);
      }
      return { path: 'INBOX', destination: target, uidMap: new Map() };
    }
  },
}));

const account: ImapAccount = {
  id: 'acct',
  name: 'Sweep test',
  host: 'imap.example.com',
  port: 993,
  user: 'user@example.com',
  password: 'secret',
  tls: true,
};

describe('ImapService.moveUids', () => {
  let service: ImapService;

  beforeEach(async () => {
    shared.instances.length = 0;
    shared.moveCalls.length = 0;
    shared.failOn.clear();
    shared.releases = 0;
    vi.spyOn(console, 'log').mockImplementation(() => {});
    service = new ImapService();
    await service.connect(account);
  });

  it('issues one UID MOVE per chunk with a UID-set string and returns counts', async () => {
    const result = await service.moveUids('acct', 'INBOX', [1, 2, 3, 4, 5], 'Archive', 2);

    expect(shared.moveCalls.map(c => c.range)).toEqual(['1,2', '3,4', '5']);
    expect(shared.moveCalls.every(c => c.target === 'Archive' && c.options?.uid === true)).toBe(true);
    expect(result).toEqual({ moved: 5, failed: 0, errors: [] });
    expect(shared.releases).toBe(3); // the mailbox lock is released after every chunk
    expect(shared.instances).toHaveLength(1);
  });

  it('records a failed chunk, keeps going, and reconnects for the next chunk', async () => {
    shared.failOn.add('3,4');

    const result = await service.moveUids('acct', 'INBOX', [1, 2, 3, 4, 5, 6], 'Archive', 2);

    expect(shared.moveCalls.map(c => c.range)).toEqual(['1,2', '3,4', '5,6']);
    expect(result.moved).toBe(4);
    expect(result.failed).toBe(2);
    expect(result.errors).toEqual(['Failed to move UIDs 3-4 to Archive: MOVE rejected for 3,4']);
    expect(shared.releases).toBe(3); // released in finally even for the failed chunk
    // The failed chunk flagged the connection dead; the next chunk got a fresh client.
    expect(shared.instances).toHaveLength(2);
  });

  it('defaults to chunks of 200 and is a no-op for an empty UID list', async () => {
    const uids = Array.from({ length: 401 }, (_, i) => i + 1);
    const result = await service.moveUids('acct', 'INBOX', uids, 'Archive');
    expect(shared.moveCalls).toHaveLength(3);
    expect(shared.moveCalls[0].range.split(',')).toHaveLength(200);
    expect(shared.moveCalls[2].range).toBe('401');
    expect(result.moved).toBe(401);

    shared.moveCalls.length = 0;
    expect(await service.moveUids('acct', 'INBOX', [], 'Archive')).toEqual({ moved: 0, failed: 0, errors: [] });
    expect(shared.moveCalls).toHaveLength(0);
  });
});
