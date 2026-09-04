import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { accountTools } from '../src/tools/account-tools.js';

/**
 * imap_add_oauth_account / imap_complete_oauth_login / imap_list_accounts with
 * a scripted OAuth service — no network, no disk.
 */
const handlers: Record<string, Function> = {};

const mockServer = {
  registerTool: vi.fn((name: string, _schema: any, handler: Function) => {
    handlers[name] = handler;
  }),
};

const storedAccounts: any[] = [];

const mockAccountManager = {
  addAccount: vi.fn(async (acc: any) => {
    const stored = { ...acc, id: 'new-id' };
    storedAccounts.push(stored);
    return stored;
  }),
  updateAccount: vi.fn(async (id: string, updates: any) => ({ id, name: 'Old Name', ...updates })),
  updateOAuthTokens: vi.fn(async () => undefined),
  getAccount: vi.fn((id: string) => storedAccounts.find(a => a.id === id)),
  getAllAccounts: vi.fn(() => storedAccounts),
};

const mockImapService = {
  testConnection: vi.fn(async () => ({ success: true, folders: ['INBOX', 'Sent Items'], messageCount: 7 })),
  disconnect: vi.fn(async () => undefined),
};
const mockSmtpService = { disconnect: vi.fn() };

const tokens = {
  accessToken: 'ACCESS-SECRET',
  refreshToken: 'REFRESH-SECRET',
  accessTokenExpiresAt: 1_800_000_000_000,
  scopes: ['https://outlook.office.com/IMAP.AccessAsUser.All', 'https://outlook.office.com/SMTP.Send'],
};

let pollResult: any;
const mockOAuthService = {
  startDeviceCode: vi.fn(async (opts: any) => ({
    flowId: 'flow-1',
    userCode: 'ABCD1234',
    verificationUri: 'https://microsoft.com/devicelogin',
    expiresAt: 1_800_000_000_000,
    message: 'go sign in',
    interval: 5,
    // Echo back what the tool stored so the completion test can return it.
    _context: opts.context,
  })),
  pollDeviceCode: vi.fn(async () => pollResult),
  primeAccessToken: vi.fn(),
  forgetResourceTokens: vi.fn(),
};

const GRAPH_SCOPES = [
  'https://graph.microsoft.com/MailboxSettings.ReadWrite',
  'https://graph.microsoft.com/Mail.ReadBasic',
  'offline_access',
];

const parse = (result: any) => JSON.parse(result.content[0].text);

describe('OAuth account tools', () => {
  beforeEach(() => {
    vi.clearAllMocks();
    storedAccounts.length = 0;
    delete process.env.IMAP_MCP_MS_CLIENT_ID;
    accountTools(
      mockServer as any,
      mockAccountManager as any,
      mockImapService as any,
      mockSmtpService as any,
      mockOAuthService as any,
    );
  });

  afterEach(() => {
    delete process.env.IMAP_MCP_MS_CLIENT_ID;
  });

  describe('imap_add_oauth_account', () => {
    it('starts the device-code flow and returns the user code and URL, never the device code', async () => {
      const result = parse(await handlers.imap_add_oauth_account({
        name: 'Personal Outlook', email: 'me@outlook.com', provider: 'microsoft', clientId: 'client-123',
      }));

      expect(result.status).toBe('awaiting_user');
      expect(result.flowId).toBe('flow-1');
      expect(result.userCode).toBe('ABCD1234');
      expect(result.verificationUri).toBe('https://microsoft.com/devicelogin');
      expect(result.instructions).toContain('ABCD1234');
      expect(result.instructions).toContain('imap_complete_oauth_login');
      expect(JSON.stringify(result)).not.toMatch(/device_code|deviceCode/);

      expect(mockOAuthService.startDeviceCode).toHaveBeenCalledWith(expect.objectContaining({
        clientId: 'client-123',
        tenant: 'consumers',
        context: expect.objectContaining({
          name: 'Personal Outlook',
          email: 'me@outlook.com',
          host: 'outlook.office365.com',
          port: 993,
          smtpHost: 'smtp-mail.outlook.com',
          smtpPort: 587,
        }),
      }));
    });

    it('falls back to IMAP_MCP_MS_CLIENT_ID and errors clearly without any client ID', async () => {
      await expect(handlers.imap_add_oauth_account({ name: 'X', email: 'me@outlook.com', provider: 'microsoft' }))
        .rejects.toThrow(/IMAP_MCP_MS_CLIENT_ID/);

      process.env.IMAP_MCP_MS_CLIENT_ID = 'env-client';
      await handlers.imap_add_oauth_account({ name: 'X', email: 'me@outlook.com', provider: 'microsoft' });
      expect(mockOAuthService.startDeviceCode).toHaveBeenCalledWith(expect.objectContaining({ clientId: 'env-client' }));
    });

    it('honors tenant and host overrides', async () => {
      await handlers.imap_add_oauth_account({
        name: 'Work', email: 'me@contoso.com', provider: 'microsoft', clientId: 'c',
        tenant: 'contoso.onmicrosoft.com', smtpHost: 'smtp.office365.com',
      });
      expect(mockOAuthService.startDeviceCode).toHaveBeenCalledWith(expect.objectContaining({
        tenant: 'contoso.onmicrosoft.com',
        context: expect.objectContaining({ smtpHost: 'smtp.office365.com', tenant: 'contoso.onmicrosoft.com' }),
      }));
    });

    it('rejects an unknown accountId for re-authorization', async () => {
      await expect(handlers.imap_add_oauth_account({
        name: 'X', email: 'me@outlook.com', provider: 'microsoft', clientId: 'c', accountId: 'missing',
      })).rejects.toThrow(/not found/);
    });
  });

  describe('imap_complete_oauth_login', () => {
    it('returns pending with a retry hint while the user has not signed in', async () => {
      pollResult = { status: 'pending', retryAfterSeconds: 5, expiresAt: 1_800_000_000_000 };

      const result = parse(await handlers.imap_complete_oauth_login({ flowId: 'flow-1' }));

      expect(result.status).toBe('pending');
      expect(result.retryAfterSeconds).toBe(5);
      expect(mockAccountManager.addAccount).not.toHaveBeenCalled();
      expect(mockOAuthService.pollDeviceCode).toHaveBeenCalledWith('flow-1', { maxWaitMs: 25_000 });
    });

    it('caps the wait at what the caller asked for', async () => {
      pollResult = { status: 'pending', retryAfterSeconds: 5, expiresAt: 1 };
      await handlers.imap_complete_oauth_login({ flowId: 'flow-1', maxWaitSeconds: 10 });
      expect(mockOAuthService.pollDeviceCode).toHaveBeenCalledWith('flow-1', { maxWaitMs: 10_000 });
    });

    it('stores the account with encrypted-at-rest tokens, tests it, and never returns tokens', async () => {
      pollResult = {
        status: 'complete',
        tokens,
        context: {
          name: 'Personal Outlook', email: 'me@outlook.com', host: 'outlook.office365.com', port: 993,
          smtpHost: 'smtp-mail.outlook.com', smtpPort: 587, clientId: 'client-123', tenant: 'consumers',
        },
      };

      const raw = await handlers.imap_complete_oauth_login({ flowId: 'flow-1' });
      const result = parse(raw);

      expect(mockAccountManager.addAccount).toHaveBeenCalledWith(expect.objectContaining({
        name: 'Personal Outlook',
        user: 'me@outlook.com',
        email: 'me@outlook.com',
        password: '',
        authType: 'oauth2',
        smtp: { host: 'smtp-mail.outlook.com', port: 587, secure: false },
        oauth: expect.objectContaining({
          provider: 'microsoft',
          clientId: 'client-123',
          tenant: 'consumers',
          refreshToken: 'REFRESH-SECRET',
          accessToken: 'ACCESS-SECRET',
        }),
      }));
      expect(mockImapService.testConnection).toHaveBeenCalledTimes(1);

      expect(result.status).toBe('complete');
      expect(result.success).toBe(true);
      expect(result.accountId).toBe('new-id');
      expect(result.authType).toBe('oauth2');
      expect(result.oauth).toEqual({ provider: 'microsoft', tenant: 'consumers', clientId: 'client-123', scopes: tokens.scopes });
      expect(result.connectionTest.folderCount).toBe(2);
      expect(raw.content[0].text).not.toContain('REFRESH-SECRET');
      expect(raw.content[0].text).not.toContain('ACCESS-SECRET');
    });

    it('re-authorizes an existing account in place and drops cached connections', async () => {
      pollResult = {
        status: 'complete',
        tokens,
        context: {
          accountId: 'acc-old', name: 'Old Name', email: 'me@outlook.com', host: 'outlook.office365.com', port: 993,
          smtpHost: 'smtp-mail.outlook.com', smtpPort: 587, clientId: 'client-123', tenant: 'consumers',
        },
      };

      const result = parse(await handlers.imap_complete_oauth_login({ flowId: 'flow-1' }));

      expect(mockAccountManager.addAccount).not.toHaveBeenCalled();
      expect(mockAccountManager.updateAccount).toHaveBeenCalledWith('acc-old', expect.objectContaining({
        authType: 'oauth2',
        password: '',
        oauth: expect.objectContaining({ refreshToken: 'REFRESH-SECRET' }),
      }));
      expect(mockImapService.disconnect).toHaveBeenCalledWith('acc-old');
      expect(mockSmtpService.disconnect).toHaveBeenCalledWith('acc-old');
      expect(result.accountId).toBe('acc-old');
      expect(result.message).toMatch(/re-authorized/);
    });

    it('keeps the account but reports the failure when the connection test fails', async () => {
      pollResult = {
        status: 'complete', tokens,
        context: { name: 'X', email: 'me@outlook.com', host: 'h', port: 993, smtpHost: 's', smtpPort: 587, clientId: 'c', tenant: 'consumers' },
      };
      mockImapService.testConnection.mockResolvedValueOnce({ success: false, error: 'AUTHENTICATE failed' } as any);

      const result = parse(await handlers.imap_complete_oauth_login({ flowId: 'flow-1' }));

      expect(result.status).toBe('complete');
      expect(result.success).toBe(false);
      expect(result.connectionTest.error).toBe('AUTHENTICATE failed');
      expect(result.connectionTest.hint).toMatch(/IMAP.AccessAsUser.All/);
    });

    it('carries an earlier Graph consent over when a mailbox is re-authorized', async () => {
      storedAccounts.push({
        id: 'acc-old', name: 'Old Name', host: 'outlook.office365.com', port: 993, user: 'me@outlook.com', password: '', tls: true,
        authType: 'oauth2',
        oauth: {
          provider: 'microsoft', clientId: 'client-123', tenant: 'consumers', refreshToken: 'r', scopes: tokens.scopes,
          grantedScopes: [...tokens.scopes, 'MailboxSettings.ReadWrite', 'Mail.ReadBasic'],
        },
      });
      pollResult = {
        status: 'complete',
        tokens,
        context: {
          accountId: 'acc-old', name: 'Old Name', email: 'me@outlook.com', host: 'outlook.office365.com', port: 993,
          smtpHost: 'smtp-mail.outlook.com', smtpPort: 587, clientId: 'client-123', tenant: 'consumers',
        },
      };

      const result = parse(await handlers.imap_complete_oauth_login({ flowId: 'flow-1' }));

      expect(mockAccountManager.updateAccount).toHaveBeenCalledWith('acc-old', expect.objectContaining({
        oauth: expect.objectContaining({
          scopes: tokens.scopes,
          grantedScopes: [...tokens.scopes, 'MailboxSettings.ReadWrite', 'Mail.ReadBasic'],
        }),
      }));
      expect(mockOAuthService.forgetResourceTokens).toHaveBeenCalledWith('acc-old');
      expect(result.oauth.graphRulesConsent).toBe(true);
    });

    it('finishes a graph-consent flow by widening the existing account, never creating one or returning tokens', async () => {
      storedAccounts.push({
        id: 'acc-ms', name: 'Outlook', host: 'outlook.office365.com', port: 993, user: 'me@outlook.com', password: '', tls: true,
        authType: 'oauth2',
        oauth: { provider: 'microsoft', clientId: 'client-123', tenant: 'consumers', refreshToken: 'REFRESH-OLD', scopes: tokens.scopes },
      });
      pollResult = {
        status: 'complete',
        tokens: {
          accessToken: 'GRAPH-ACCESS-SECRET',
          refreshToken: 'REFRESH-NEW-SECRET',
          accessTokenExpiresAt: 1_800_000_000_000,
          // Microsoft reports Graph scopes in short form.
          scopes: ['MailboxSettings.ReadWrite', 'Mail.ReadBasic'],
        },
        context: { kind: 'graph-consent', accountId: 'acc-ms', scopes: GRAPH_SCOPES },
      };

      const raw = await handlers.imap_complete_oauth_login({ flowId: 'flow-1' });
      const result = parse(raw);

      expect(mockAccountManager.addAccount).not.toHaveBeenCalled();
      expect(mockAccountManager.updateAccount).not.toHaveBeenCalled();
      expect(mockAccountManager.updateOAuthTokens).toHaveBeenCalledWith('acc-ms', {
        refreshToken: 'REFRESH-NEW-SECRET',
        grantedScopes: [...tokens.scopes, 'MailboxSettings.ReadWrite', 'Mail.ReadBasic'],
      });
      // The Graph access token seeds the in-memory cache; the mail token is untouched.
      expect(mockOAuthService.primeAccessToken).toHaveBeenCalledWith(
        expect.objectContaining({ id: 'acc-ms' }), GRAPH_SCOPES, 'GRAPH-ACCESS-SECRET', 1_800_000_000_000,
      );
      expect(mockImapService.testConnection).not.toHaveBeenCalled();

      expect(result).toMatchObject({
        status: 'complete', success: true, kind: 'graph-consent', accountId: 'acc-ms',
        graphRulesConsent: true, nextStep: 'imap_outlook_list_rules',
      });
      expect(raw.content[0].text).not.toContain('SECRET');
    });

    it('reports a graph-consent flow that Microsoft completed without the Graph scopes', async () => {
      storedAccounts.push({
        id: 'acc-ms', name: 'Outlook', host: 'h', port: 993, user: 'u', password: '', tls: true, authType: 'oauth2',
        oauth: { provider: 'microsoft', clientId: 'c', tenant: 'consumers', refreshToken: 'r', scopes: tokens.scopes },
      });
      pollResult = {
        status: 'complete',
        tokens: { ...tokens, scopes: ['Mail.ReadBasic'] },
        context: { kind: 'graph-consent', accountId: 'acc-ms', scopes: GRAPH_SCOPES },
      };

      const result = parse(await handlers.imap_complete_oauth_login({ flowId: 'flow-1' }));

      expect(result.success).toBe(false);
      expect(result.graphRulesConsent).toBe(false);
      expect(result.message).toMatch(/MailboxSettings.ReadWrite/);
      expect(result.nextStep).toBeUndefined();
    });

    it('fails a graph-consent flow cleanly when the account vanished meanwhile', async () => {
      pollResult = {
        status: 'complete', tokens,
        context: { kind: 'graph-consent', accountId: 'gone', scopes: GRAPH_SCOPES },
      };
      await expect(handlers.imap_complete_oauth_login({ flowId: 'flow-1' })).rejects.toThrow(/no longer exists/);
      expect(mockAccountManager.updateOAuthTokens).not.toHaveBeenCalled();
    });

    it.each(['expired', 'denied', 'error'])('returns a structured %s result instead of throwing', async (status) => {
      pollResult = { status, error: `it went ${status}` };
      const result = parse(await handlers.imap_complete_oauth_login({ flowId: 'flow-1' }));
      expect(result.status).toBe(status);
      expect(result.success).toBe(false);
      expect(result.error).toBe(`it went ${status}`);
      expect(mockAccountManager.addAccount).not.toHaveBeenCalled();
    });
  });

  describe('imap_list_accounts', () => {
    it('reports authType and the public OAuth fields, never tokens or passwords', async () => {
      storedAccounts.push(
        { id: 'p1', name: 'Plain', host: 'imap.example.com', port: 993, user: 'u', password: 'pw-secret', tls: true },
        {
          id: 'o1', name: 'Outlook', host: 'outlook.office365.com', port: 993, user: 'me@outlook.com', password: '', tls: true,
          authType: 'oauth2',
          oauth: { provider: 'microsoft', clientId: 'client-123', tenant: 'consumers', refreshToken: 'REFRESH-SECRET', accessToken: 'ACCESS-SECRET', scopes: ['s'] },
        },
      );

      const raw = await handlers.imap_list_accounts({});
      const result = parse(raw);

      expect(result.accounts[0].authType).toBe('password');
      expect(result.accounts[0].oauth).toBeUndefined();
      expect(result.accounts[1].authType).toBe('oauth2');
      expect(result.accounts[1].oauth).toEqual({ provider: 'microsoft', tenant: 'consumers', clientId: 'client-123', scopes: ['s'] });
      expect(result.accounts[1].oauth.grantedScopes).toBeUndefined();
      expect(raw.content[0].text).not.toContain('REFRESH-SECRET');
      expect(raw.content[0].text).not.toContain('ACCESS-SECRET');
      expect(raw.content[0].text).not.toContain('pw-secret');
    });
  });

  describe('imap_update_account', () => {
    it('rejects a password for an OAuth account and points at the re-authorization flow', async () => {
      storedAccounts.push({
        id: 'o1', name: 'Outlook', host: 'outlook.office365.com', port: 993, user: 'me@outlook.com', password: '', tls: true,
        authType: 'oauth2',
        oauth: { provider: 'microsoft', clientId: 'c', tenant: 'consumers', refreshToken: 'r', scopes: [] },
      });

      await expect(handlers.imap_update_account({ accountId: 'o1', password: 'new' }))
        .rejects.toThrow(/imap_add_oauth_account/);
      expect(mockAccountManager.updateAccount).not.toHaveBeenCalled();

      // Non-credential fields still update.
      await handlers.imap_update_account({ accountId: 'o1', name: 'Renamed' });
      expect(mockAccountManager.updateAccount).toHaveBeenCalledWith('o1', expect.objectContaining({ name: 'Renamed' }));
    });
  });
});
