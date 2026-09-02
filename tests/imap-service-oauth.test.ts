import { describe, it, expect, vi, beforeEach } from 'vitest';
import { ImapService } from '../src/services/imap-service.js';
import type { ImapAccount } from '../src/types/index.js';

// Records the options every ImapFlow is constructed with, and lets a test
// script the outcome of each successive connect() call.
const constructorOptions: any[] = [];
const connectOutcomes: Array<Error | null> = [];

vi.mock('imapflow', () => ({
  ImapFlow: class {
    constructor(options: any) {
      constructorOptions.push(options);
    }
    connect() {
      const outcome = connectOutcomes.shift();
      return outcome ? Promise.reject(outcome) : Promise.resolve();
    }
    logout() {
      return Promise.resolve();
    }
    list() {
      return Promise.resolve([{ path: 'INBOX' }]);
    }
    status() {
      return Promise.resolve({ messages: 3 });
    }
    on() {}
  },
}));

const oauthAccount = (): ImapAccount => ({
  id: 'ms',
  name: 'Outlook',
  host: 'outlook.office365.com',
  port: 993,
  user: 'me@outlook.com',
  password: '',
  tls: true,
  authType: 'oauth2',
  oauth: {
    provider: 'microsoft',
    clientId: 'client',
    tenant: 'consumers',
    refreshToken: 'refresh',
    scopes: [],
  },
});

const authFailure = () => Object.assign(new Error('AUTHENTICATE failed'), { authenticationFailed: true });

describe('ImapService with OAuth 2.0 accounts', () => {
  let service: ImapService;
  let oauth: { getValidAccessToken: ReturnType<typeof vi.fn>; forceRefresh: ReturnType<typeof vi.fn> };

  beforeEach(() => {
    constructorOptions.length = 0;
    connectOutcomes.length = 0;
    oauth = {
      getValidAccessToken: vi.fn(async () => 'TOKEN-1'),
      forceRefresh: vi.fn(async () => 'TOKEN-2'),
    };
    service = new ImapService();
    service.setOAuthService(oauth as any);
  });

  it('connect() hands imapflow the access token and no password', async () => {
    await service.connect(oauthAccount());

    expect(constructorOptions).toHaveLength(1);
    expect(constructorOptions[0].auth).toEqual({ user: 'me@outlook.com', accessToken: 'TOKEN-1' });
    expect(constructorOptions[0].auth.pass).toBeUndefined();
    expect(constructorOptions[0].host).toBe('outlook.office365.com');
    expect(constructorOptions[0].secure).toBe(true);
  });

  it('password accounts still log in with pass', async () => {
    await service.connect({ ...oauthAccount(), authType: undefined, oauth: undefined, password: 'pw' });

    expect(constructorOptions[0].auth.pass).toBe('pw');
    expect(constructorOptions[0].auth.accessToken).toBeUndefined();
    expect(oauth.getValidAccessToken).not.toHaveBeenCalled();
  });

  it('retries once with a force-refreshed token when the server rejects the first one', async () => {
    connectOutcomes.push(authFailure(), null);

    await service.connect(oauthAccount());

    expect(oauth.forceRefresh).toHaveBeenCalledTimes(1);
    expect(constructorOptions).toHaveLength(2);
    expect(constructorOptions[0].auth.accessToken).toBe('TOKEN-1');
    expect(constructorOptions[1].auth.accessToken).toBe('TOKEN-2');
  });

  it('surfaces the error when the retry also fails', async () => {
    connectOutcomes.push(authFailure(), authFailure());

    await expect(service.connect(oauthAccount())).rejects.toThrow(/AUTHENTICATE failed/);
    expect(oauth.forceRefresh).toHaveBeenCalledTimes(1);
  });

  it('does not refresh for non-authentication failures', async () => {
    connectOutcomes.push(new Error('ECONNREFUSED'));

    await expect(service.connect(oauthAccount())).rejects.toThrow(/ECONNREFUSED/);
    expect(oauth.forceRefresh).not.toHaveBeenCalled();
  });

  it('testConnection() uses the access token too', async () => {
    const result = await service.testConnection(oauthAccount());

    expect(result.success).toBe(true);
    expect(result.folders).toEqual(['INBOX']);
    expect(constructorOptions[0].auth.accessToken).toBe('TOKEN-1');
    expect(constructorOptions[0].auth.pass).toBeUndefined();
  });

  it('testConnection() reports a missing refresh token by its env variable instead of dialing out', async () => {
    const account = oauthAccount();
    account.oauth!.refreshToken = '';

    const result = await service.testConnection(account);

    expect(result.success).toBe(false);
    expect(result.error).toMatch(/IMAP_MCP_ACCOUNT_OUTLOOK_OAUTH_REFRESH_TOKEN/);
    expect(constructorOptions).toHaveLength(0);
  });
});
