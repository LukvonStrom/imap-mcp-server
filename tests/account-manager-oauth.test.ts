import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import { promises as fs } from 'fs';

// Same fs mock as tests/account-manager.test.ts: nothing touches the real
// ~/.imap-mcp, and every write is captured so the on-disk shape can be asserted.
vi.mock('fs', async () => {
  const actual = await vi.importActual('fs');
  return {
    ...actual,
    promises: {
      readFile: vi.fn(),
      writeFile: vi.fn(),
      mkdir: vi.fn(),
      chmod: vi.fn(),
    },
    readFileSync: vi.fn(),
    writeFileSync: vi.fn(),
    mkdirSync: vi.fn(),
  };
});

import { AccountManager } from '../src/services/account-manager.js';
import { readFileSync } from 'fs';
import type { OAuthConfig } from '../src/types/index.js';

const REFRESH = 'plaintext-refresh-token';
const ACCESS = 'plaintext-access-token';

const oauth = (): OAuthConfig => ({
  provider: 'microsoft',
  clientId: 'client-123',
  tenant: 'consumers',
  refreshToken: REFRESH,
  accessToken: ACCESS,
  accessTokenExpiresAt: 1_700_000_000_000,
  scopes: ['https://outlook.office.com/IMAP.AccessAsUser.All'],
});

const addOAuthAccount = (manager: AccountManager, name = 'Outlook') => manager.addAccount({
  name,
  host: 'outlook.office365.com',
  port: 993,
  user: 'me@outlook.com',
  password: '',
  tls: true,
  authType: 'oauth2',
  oauth: oauth(),
});

const lastWrite = () => {
  const writes = vi.mocked(fs.writeFile).mock.calls;
  return JSON.parse(writes[writes.length - 1][1] as string);
};

describe('AccountManager OAuth token storage', () => {
  const mockEncryptionKey = '0123456789abcdef0123456789abcdef0123456789abcdef0123456789abcdef';

  beforeEach(() => {
    vi.clearAllMocks();
    vi.mocked(readFileSync).mockReturnValue(mockEncryptionKey);
    vi.mocked(fs.writeFile).mockResolvedValue(undefined);
    vi.mocked(fs.mkdir).mockResolvedValue(undefined);
  });

  afterEach(() => {
    for (const key of Object.keys(process.env)) {
      if (key.startsWith('IMAP_MCP_ACCOUNT_')) delete process.env[key];
    }
  });

  it('encrypts the refresh and access tokens on disk and keeps the rest readable', async () => {
    const manager = new AccountManager();
    await addOAuthAccount(manager);

    const raw = vi.mocked(fs.writeFile).mock.calls.at(-1)![1] as string;
    expect(raw).not.toContain(REFRESH);
    expect(raw).not.toContain(ACCESS);

    const saved = JSON.parse(raw)[0];
    expect(saved.authType).toBe('oauth2');
    expect(saved.oauth.refreshToken).toContain(':');
    expect(saved.oauth.accessToken).toContain(':');
    expect(saved.oauth.clientId).toBe('client-123');
    expect(saved.oauth.tenant).toBe('consumers');
    expect(saved.oauth.scopes).toEqual(['https://outlook.office.com/IMAP.AccessAsUser.All']);
  });

  it('round-trips the tokens through addAccount, getAccount, getAllAccounts, and getAccountByName', async () => {
    const manager = new AccountManager();
    const created = await addOAuthAccount(manager);

    expect(created.oauth?.refreshToken).toBe(REFRESH);
    expect(manager.getAccount(created.id)?.oauth?.refreshToken).toBe(REFRESH);
    expect(manager.getAccount(created.id)?.oauth?.accessToken).toBe(ACCESS);
    expect(manager.getAllAccounts()[0].oauth?.refreshToken).toBe(REFRESH);
    expect(manager.getAccountByName('Outlook')?.oauth?.accessToken).toBe(ACCESS);
  });

  it('updateOAuthTokens rotates only the supplied fields, encrypted', async () => {
    const manager = new AccountManager();
    const created = await addOAuthAccount(manager);

    await manager.updateOAuthTokens(created.id, { accessToken: 'new-access', accessTokenExpiresAt: 42 });

    const saved = lastWrite()[0];
    expect(JSON.stringify(saved)).not.toContain('new-access');
    expect(saved.oauth.accessTokenExpiresAt).toBe(42);

    const account = manager.getAccount(created.id)!;
    expect(account.oauth?.accessToken).toBe('new-access');
    expect(account.oauth?.refreshToken).toBe(REFRESH); // untouched
    expect(account.oauth?.clientId).toBe('client-123');

    await manager.updateOAuthTokens(created.id, { refreshToken: 'rotated' });
    expect(manager.getAccount(created.id)?.oauth?.refreshToken).toBe('rotated');
    expect(JSON.stringify(lastWrite())).not.toContain('rotated');
  });

  it('updateOAuthTokens refuses password accounts', async () => {
    const manager = new AccountManager();
    const created = await manager.addAccount({
      name: 'Plain', host: 'imap.test.com', port: 993, user: 'u', password: 'p', tls: true,
    });
    await expect(manager.updateOAuthTokens(created.id, { accessToken: 'x' })).rejects.toThrow(/not an OAuth account/);
  });

  it('updateAccount encrypts a replaced oauth block', async () => {
    const manager = new AccountManager();
    const created = await manager.addAccount({
      name: 'Plain', host: 'imap.test.com', port: 993, user: 'u', password: 'p', tls: true,
    });

    const updated = await manager.updateAccount(created.id, { authType: 'oauth2', password: '', oauth: oauth() });

    expect(updated.authType).toBe('oauth2');
    expect(updated.oauth?.refreshToken).toBe(REFRESH);
    expect(JSON.stringify(lastWrite())).not.toContain(REFRESH);
  });

  it('lets IMAP_MCP_ACCOUNT_<NAME>_OAUTH_REFRESH_TOKEN supply the refresh token and consumes the variable', async () => {
    process.env.IMAP_MCP_ACCOUNT_OUTLOOK_OAUTH_REFRESH_TOKEN = 'env-refresh';

    const manager = new AccountManager();
    expect(process.env.IMAP_MCP_ACCOUNT_OUTLOOK_OAUTH_REFRESH_TOKEN).toBeUndefined();

    const created = await manager.addAccount({
      name: 'Outlook',
      host: 'outlook.office365.com',
      port: 993,
      user: 'me@outlook.com',
      password: '',
      tls: true,
      authType: 'oauth2',
      oauth: { ...oauth(), refreshToken: '' },
    });

    expect(manager.getAccount(created.id)?.oauth?.refreshToken).toBe('env-refresh');
    expect(JSON.stringify(lastWrite())).not.toContain('env-refresh');
  });

  it('existing password accounts stay untouched (no authType, no oauth)', async () => {
    const manager = new AccountManager();
    const created = await manager.addAccount({
      name: 'Plain', host: 'imap.test.com', port: 993, user: 'u', password: 'p', tls: true,
    });
    const account = manager.getAccount(created.id)!;
    expect(account.authType).toBeUndefined();
    expect(account.oauth).toBeUndefined();
    expect(account.password).toBe('p');
  });
});
