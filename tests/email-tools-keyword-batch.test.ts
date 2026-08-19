import { describe, it, expect, vi, beforeEach } from 'vitest';
import { emailTools } from '../src/tools/email-tools.js';

const handlers = new Map<string, Function>();

const mockServer = {
  registerTool: vi.fn((name: string, _schema: any, handler: Function) => {
    handlers.set(name, handler);
  }),
};

const mockImapService = {
  addKeywordToUids: vi.fn(),
  removeKeywordFromUids: vi.fn(),
};

const mockAccountManager = { resolveAccountId: (id: string) => id };
const parse = (result: any) => JSON.parse(result.content[0].text);

describe('imap_add_keyword / imap_remove_keyword with UID lists', () => {
  beforeEach(() => {
    vi.clearAllMocks();
    handlers.clear();
    emailTools(mockServer as any, mockImapService as any, mockAccountManager as any, {} as any);
  });

  it('still accepts a single uid and phrases the message in the singular', async () => {
    mockImapService.addKeywordToUids.mockResolvedValueOnce(1);

    const parsed = parse(await handlers.get('imap_add_keyword')!({
      accountId: 'acc1', folder: 'INBOX', uid: 42, keyword: '$imapmcpChecked',
    }));

    expect(mockImapService.addKeywordToUids).toHaveBeenCalledWith('acc1', 'INBOX', [42], '$imapmcpChecked');
    expect(parsed.message).toBe('Keyword "$imapmcpChecked" added to email 42');
  });

  it('tags many messages in one call', async () => {
    mockImapService.addKeywordToUids.mockResolvedValueOnce(3);

    const parsed = parse(await handlers.get('imap_add_keyword')!({
      accountId: 'acc1', folder: 'Unsortiert', uid: [1, 2, 3], keyword: '$imapmcpChecked',
    }));

    expect(mockImapService.addKeywordToUids).toHaveBeenCalledWith('acc1', 'Unsortiert', [1, 2, 3], '$imapmcpChecked');
    expect(parsed.count).toBe(3);
    expect(parsed.message).toBe('Keyword "$imapmcpChecked" added to 3 emails');
  });

  it('clears a keyword from many messages in one call', async () => {
    mockImapService.removeKeywordFromUids.mockResolvedValueOnce(1231);

    const parsed = parse(await handlers.get('imap_remove_keyword')!({
      accountId: 'acc1', folder: 'Unsortiert', uid: [1, 2, 3], keyword: '$imapmcpChecked',
    }));

    expect(mockImapService.removeKeywordFromUids).toHaveBeenCalledWith('acc1', 'Unsortiert', [1, 2, 3], '$imapmcpChecked');
    expect(parsed.count).toBe(1231);
  });

  it('recovers a uid list a client stringified (issue #127)', async () => {
    // The tools take the value after Zod preprocessing, so exercise the schema.
    const schema = (mockServer.registerTool as any).mock.calls
      .find((c: any[]) => c[0] === 'imap_remove_keyword')[1].inputSchema.uid;

    expect(schema.safeParse('[10,20,30]').data).toEqual([10, 20, 30]);
    expect(schema.safeParse(7).data).toBe(7);
  });
});
