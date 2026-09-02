import { describe, it, expect, vi, beforeEach } from 'vitest';
import type { ImapAccount } from '../src/types/index.js';

// Capture the options nodemailer.createTransport receives; the transporter
// itself only needs verify()/close().
const createTransportCalls: any[] = [];

vi.mock('nodemailer', () => ({
  default: {
    createTransport: vi.fn((options: any) => {
      createTransportCalls.push(options);
      return { verify: vi.fn(async () => true), close: vi.fn(), sendMail: vi.fn() };
    }),
  },
}));

import { SmtpService } from '../src/services/smtp-service.js';

const oauthAccount = (): ImapAccount => ({
  id: 'ms',
  name: 'Outlook',
  host: 'outlook.office365.com',
  port: 993,
  user: 'me@outlook.com',
  password: '',
  tls: true,
  email: 'me@outlook.com',
  authType: 'oauth2',
  oauth: {
    provider: 'microsoft',
    clientId: 'client',
    tenant: 'consumers',
    refreshToken: 'refresh',
    scopes: [],
  },
  smtp: { host: 'smtp-mail.outlook.com', port: 587, secure: false },
});

describe('SmtpService with OAuth 2.0 accounts', () => {
  let smtp: SmtpService;
  let token: string;
  let oauth: { getValidAccessToken: ReturnType<typeof vi.fn> };

  beforeEach(() => {
    createTransportCalls.length = 0;
    token = 'TOKEN-1';
    oauth = { getValidAccessToken: vi.fn(async () => token) };
    smtp = new SmtpService();
    smtp.setOAuthService(oauth as any);
  });

  it('builds an XOAUTH2 transporter with the access token and no password', async () => {
    await smtp.createTransporter(oauthAccount());

    expect(createTransportCalls).toHaveLength(1);
    const opts = createTransportCalls[0];
    expect(opts.auth).toEqual({ type: 'OAuth2', user: 'me@outlook.com', accessToken: 'TOKEN-1' });
    expect(opts.auth.pass).toBeUndefined();
    expect(opts.host).toBe('smtp-mail.outlook.com');
    expect(opts.port).toBe(587);
    expect(opts.secure).toBe(false);
    expect(opts.requireTLS).toBe(true);
  });

  it('falls back to the Outlook SMTP host mapping when the account has no smtp block', async () => {
    const account = oauthAccount();
    delete account.smtp;

    await smtp.createTransporter(account);

    expect(createTransportCalls[0].host).toBe('smtp.office365.com');
    expect(createTransportCalls[0].port).toBe(587);
    expect(createTransportCalls[0].auth.type).toBe('OAuth2');
  });

  it('reuses the cached transporter while the token is unchanged and rebuilds it after a refresh', async () => {
    const account = oauthAccount();
    const first = await smtp.createTransporter(account);
    const again = await smtp.createTransporter(account);
    expect(again).toBe(first);
    expect(createTransportCalls).toHaveLength(1);

    token = 'TOKEN-2';
    const rebuilt = await smtp.createTransporter(account);
    expect(rebuilt).not.toBe(first);
    expect(createTransportCalls).toHaveLength(2);
    expect(createTransportCalls[1].auth.accessToken).toBe('TOKEN-2');
  });

  it('password accounts keep the plain user/pass auth', async () => {
    await smtp.createTransporter({ ...oauthAccount(), authType: undefined, oauth: undefined, password: 'pw' });

    expect(createTransportCalls[0].auth).toEqual({ user: 'me@outlook.com', pass: 'pw' });
    expect(oauth.getValidAccessToken).not.toHaveBeenCalled();
  });

  it('names the refresh-token env variable when the OAuth credential is env-managed but unset', async () => {
    const account = oauthAccount();
    account.oauth!.refreshToken = '';

    await expect(smtp.createTransporter(account)).rejects.toThrow(/IMAP_MCP_ACCOUNT_OUTLOOK_OAUTH_REFRESH_TOKEN/);
    expect(createTransportCalls).toHaveLength(0);
  });
});
