import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { promises as fsp } from 'fs';
import { join } from 'path';
import { tmpdir } from 'os';

// Regression tests for the hostile-audit findings (2026-09-03):
//  H1 imap_download_attachment.savePath could write anywhere on disk
//  H2 imap_update_account could redirect a stored credential to a new host
//  M3 an unparseable date silently dropped a bulk-delete criterion

let downloadHandler: Function;
let bulkDeleteHandler: Function;
let searchHandler: Function;
let updateHandler: Function;
let addHandler: Function;
let addOAuthHandler: Function;

const capture = (map: Record<string, (fn: Function) => void>) => ({
  registerTool: vi.fn((name: string, _cfg: any, fn: Function) => { map[name]?.(fn); }),
});

const emailServer = capture({
  imap_download_attachment: fn => { downloadHandler = fn; },
  imap_bulk_delete_by_search: fn => { bulkDeleteHandler = fn; },
  imap_search_emails: fn => { searchHandler = fn; },
});
const accountServer = capture({
  imap_update_account: fn => { updateHandler = fn; },
  imap_add_account: fn => { addHandler = fn; },
  imap_add_oauth_account: fn => { addOAuthHandler = fn; },
});

const mockImapService: any = {
  getAttachmentContent: vi.fn(async () => ({ content: Buffer.from('payload'), contentType: 'text/plain', filename: 'note.txt' })),
  searchEmails: vi.fn(async () => []),
  bulkDelete: vi.fn(),
  disconnect: vi.fn(),
};
const mockSmtpService: any = { disconnect: vi.fn() };

const passwordAccount = { id: 'pw', name: 'Work', host: 'imap.example.com', port: 993, user: 'u', password: 'secret', tls: true };
const oauthAccount = { id: 'oa', name: 'Outlook', host: 'outlook.office365.com', port: 993, user: 'me@outlook.com', password: '', tls: true, authType: 'oauth2', oauth: { provider: 'microsoft', clientId: 'cid', tenant: 'consumers', refreshToken: 'r', scopes: [] } };

const accounts: any[] = [passwordAccount, oauthAccount];
const mockAccountManager: any = {
  resolveAccountId: (id: string) => id,
  getAccount: vi.fn((id: string) => accounts.find(a => a.id === id)),
  getAllAccounts: vi.fn(() => accounts),
  updateAccount: vi.fn(async (id: string, updates: any) => ({ ...accounts.find(a => a.id === id), ...updates })),
  addAccount: vi.fn(async (a: any) => ({ ...a, id: 'new' })),
};
const mockOAuthService: any = {
  startDeviceCode: vi.fn(async () => ({ flowId: 'f', userCode: 'CODE', verificationUri: 'https://microsoft.com/devicelogin', expiresAt: Date.now() + 60000, message: 'm', interval: 5 })),
};

describe('security hardening', () => {
  let dir: string;
  let outside: string;

  beforeEach(async () => {
    vi.clearAllMocks();
    dir = await fsp.mkdtemp(join(tmpdir(), 'imap-mcp-sec-dl-'));
    outside = await fsp.mkdtemp(join(tmpdir(), 'imap-mcp-sec-out-'));
    process.env.IMAP_ATTACHMENT_ROOTS = dir;
    const { emailTools } = await import('../src/tools/email-tools.js');
    const { accountTools } = await import('../src/tools/account-tools.js');
    emailTools(emailServer as any, mockImapService, mockAccountManager, mockSmtpService);
    accountTools(accountServer as any, mockAccountManager, mockImapService, mockSmtpService, mockOAuthService);
  });

  afterEach(async () => {
    delete process.env.IMAP_ATTACHMENT_ROOTS;
    await fsp.rm(dir, { recursive: true, force: true });
    await fsp.rm(outside, { recursive: true, force: true });
  });

  describe('H1: imap_download_attachment savePath confinement', () => {
    const args = { accountId: 'pw', folder: 'INBOX', uid: 1, filename: 'note.txt', extractText: false };

    it('refuses a savePath outside the allowed roots and writes nothing', async () => {
      const target = join(outside, '.zshrc');
      await expect(downloadHandler({ ...args, savePath: target })).rejects.toThrow(/outside the allowed download directories/);
      await expect(fsp.stat(target)).rejects.toThrow();
    });

    it('refuses to mkdir -p an attacker-chosen tree outside the roots', async () => {
      const target = join(outside, 'Library', 'LaunchAgents', 'x.plist');
      await expect(downloadHandler({ ...args, savePath: target })).rejects.toThrow(/outside the allowed download directories/);
      await expect(fsp.stat(join(outside, 'Library'))).rejects.toThrow();
    });

    it('refuses a savePath whose parent is a symlink pointing outside the roots', async () => {
      await fsp.symlink(outside, join(dir, 'escape'));
      await expect(downloadHandler({ ...args, savePath: join(dir, 'escape', 'note.txt') })).rejects.toThrow(/symlink/);
      await expect(fsp.stat(join(outside, 'note.txt'))).rejects.toThrow();
    });

    it('accepts a savePath inside a root, creating subdirectories, but never overwrites', async () => {
      const target = join(dir, 'sub', 'note.txt');
      const result = await downloadHandler({ ...args, savePath: target });
      const payload = JSON.parse(result.content[0].text);
      expect(payload.saved).toBe(true);
      expect(await fsp.readFile(target, 'utf8')).toBe('payload');

      await expect(downloadHandler({ ...args, savePath: target })).rejects.toThrow(/Refusing to overwrite/);
    });

    it('confines the default location to the download dir using the basename of the MIME filename', async () => {
      process.env.IMAP_DOWNLOAD_DIR_UNUSED = '1'; // DOWNLOAD_DIR is module-level; assert on the returned path instead
      mockImapService.getAttachmentContent.mockResolvedValueOnce({ content: Buffer.from('x'), contentType: 'text/plain', filename: '../../evil.txt' });
      const result = await downloadHandler(args);
      const payload = JSON.parse(result.content[0].text);
      expect(payload.path.endsWith('/evil.txt')).toBe(true);
      expect(payload.path).not.toContain('..');
      await fsp.rm(payload.path, { force: true });
    });
  });

  describe('H2: imap_update_account credential redirection', () => {
    it('refuses to change the host of a password account without the password', async () => {
      await expect(updateHandler({ accountId: 'pw', host: 'evil.example', port: 143, tls: false }))
        .rejects.toThrow(/requires re-entering the account password/);
      expect(mockAccountManager.updateAccount).not.toHaveBeenCalled();
    });

    it.each(['port', 'user', 'tls', 'allowStartTLS', 'tlsRejectUnauthorized', 'smtpHost', 'smtpPort', 'smtpSecure', 'smtpUser'])(
      'treats %s as a connection-affecting change', async (field) => {
        const value: any = ['port', 'smtpPort'].includes(field) ? 1 : ['tls', 'allowStartTLS', 'tlsRejectUnauthorized', 'smtpSecure'].includes(field) ? false : 'x';
        await expect(updateHandler({ accountId: 'pw', [field]: value })).rejects.toThrow(/requires re-entering/);
      },
    );

    it('allows the change when the password is supplied in the same call', async () => {
      const result = await updateHandler({ accountId: 'pw', host: 'imap.new.example', password: 'secret' });
      expect(JSON.parse(result.content[0].text).success).toBe(true);
      expect(mockAccountManager.updateAccount).toHaveBeenCalledWith('pw', expect.objectContaining({ host: 'imap.new.example', password: 'secret' }));
    });

    it('allows non-connection edits (name, saveToSent, defaultBcc) without the password', async () => {
      const result = await updateHandler({ accountId: 'pw', saveToSent: false, defaultBcc: 'me@example.com' });
      expect(JSON.parse(result.content[0].text).success).toBe(true);
    });

    it('keeps OAuth accounts on Microsoft hosts with validated TLS', async () => {
      await expect(updateHandler({ accountId: 'oa', host: 'evil.example' })).rejects.toThrow(/not a Microsoft mail endpoint/);
      await expect(updateHandler({ accountId: 'oa', smtpHost: 'smtp.evil.example' })).rejects.toThrow(/not a Microsoft mail endpoint/);
      await expect(updateHandler({ accountId: 'oa', tlsRejectUnauthorized: false })).rejects.toThrow(/always use validated TLS/);
      const ok = await updateHandler({ accountId: 'oa', host: 'outlook.office.com' });
      expect(JSON.parse(ok.content[0].text).success).toBe(true);
    });

    it('imap_add_oauth_account refuses non-Microsoft hosts before starting the device-code flow', async () => {
      await expect(addOAuthHandler({ provider: 'microsoft', email: 'x@outlook.com', clientId: 'cid', host: 'imap.evil.example' }))
        .rejects.toThrow(/not a Microsoft mail endpoint/);
      expect(mockOAuthService.startDeviceCode).not.toHaveBeenCalled();
    });

    it('refuses duplicate account names (env credential overrides are keyed by name)', async () => {
      await expect(addHandler({ name: 'Work', host: 'evil.example', port: 993, user: 'u', password: 'p', tls: true }))
        .rejects.toThrow(/already exists/);
      expect(mockAccountManager.addAccount).not.toHaveBeenCalled();
      await expect(updateHandler({ accountId: 'oa', name: 'Work' })).rejects.toThrow(/already exists/);
      await expect(addOAuthHandler({ provider: 'microsoft', email: 'x@outlook.com', clientId: 'cid', name: 'Work' }))
        .rejects.toThrow(/already exists/);
    });
  });

  describe('M3: strict date parsing', () => {
    it('rejects an unparseable date instead of silently dropping the criterion', async () => {
      await expect(bulkDeleteHandler({ accountId: 'pw', folder: 'INBOX', from: 'a@b.c', before: 'not a date', chunkSize: 100, dryRun: true }))
        .rejects.toThrow(/Invalid date/);
      expect(mockImapService.searchEmails).not.toHaveBeenCalled();
      await expect(searchHandler({ accountId: 'pw', folder: 'INBOX', since: '2026-13-45x', limit: 10, searchAllFolders: false, includeTrash: false, includeSpam: false, includeDrafts: false }))
        .rejects.toThrow(/Invalid date/);
    });

    it('still accepts YYYY-MM-DD', async () => {
      const result = await bulkDeleteHandler({ accountId: 'pw', folder: 'INBOX', from: 'a@b.c', before: '2026-01-31', chunkSize: 100, dryRun: true });
      expect(JSON.parse(result.content[0].text).message).toMatch(/No emails matched/);
    });
  });
});
