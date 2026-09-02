import { describe, it, expect, beforeAll, afterAll } from 'vitest';
import type { AddressInfo } from 'net';
import type { Server } from 'http';
import { WebUIServer } from '../src/web/server.js';

// getAllAccounts() returns DECRYPTED credentials by design (the manager's job).
// The wizard's HTTP API is unauthenticated and CORS-open, so the API layer must
// never let those credentials cross the wire. Inject a fake manager holding
// obvious secrets and assert the /api/accounts response is scrubbed.
const SECRET = 'imap-plaintext-secret';
const SMTP_SECRET = 'smtp-plaintext-secret';
const REFRESH_SECRET = 'oauth-plaintext-refresh-token';
const ACCESS_SECRET = 'oauth-plaintext-access-token';

const fakeAccountManager = {
  addAccount: async (acc: any) => ({ ...acc, id: 'created' }),
  getAllAccounts: () => [
    {
      id: 'a1',
      name: 'Work',
      host: 'imap.example.com',
      port: 993,
      user: 'user@example.com',
      password: SECRET,
      tls: true,
      smtp: { host: 'smtp.example.com', port: 587, secure: false, password: SMTP_SECRET },
    },
    {
      id: 'o1',
      name: 'Outlook',
      host: 'outlook.office365.com',
      port: 993,
      user: 'me@outlook.com',
      password: '',
      tls: true,
      authType: 'oauth2',
      oauth: {
        provider: 'microsoft',
        clientId: 'client-123',
        tenant: 'consumers',
        refreshToken: REFRESH_SECRET,
        accessToken: ACCESS_SECRET,
        accessTokenExpiresAt: 1,
        scopes: ['https://outlook.office.com/IMAP.AccessAsUser.All'],
      },
    },
  ],
  getAccount: (id: string) =>
    id === 'a1'
      ? {
          id: 'a1',
          name: 'Work',
          host: 'imap.example.com',
          port: 993,
          user: 'user@example.com',
          password: SECRET,
          tls: true,
          smtp: { host: 'smtp.example.com', port: 587, secure: false, password: SMTP_SECRET },
        }
      : undefined,
};

let httpServer: Server;
let baseUrl: string;

beforeAll(async () => {
  const wizard = new WebUIServer(0, {
    accountManager: fakeAccountManager as any,
    imapService: {} as any,
  });
  await new Promise<void>((resolve) => {
    httpServer = wizard.getApp().listen(0, () => resolve());
  });
  const { port } = httpServer.address() as AddressInfo;
  baseUrl = `http://127.0.0.1:${port}`;
});

afterAll(() => {
  httpServer?.close();
});

describe('Web wizard credential exposure', () => {
  it('GET /api/accounts never returns passwords', async () => {
    const res = await fetch(`${baseUrl}/api/accounts`);
    expect(res.ok).toBe(true);
    const raw = await res.text();

    // Belt: the plaintext secrets must not appear anywhere in the payload.
    expect(raw).not.toContain(SECRET);
    expect(raw).not.toContain(SMTP_SECRET);

    // Braces: the fields must be absent, and the rest of the account intact.
    const accounts = JSON.parse(raw);
    expect(Array.isArray(accounts)).toBe(true);
    expect(accounts[0].password).toBeUndefined();
    expect(accounts[0].smtp?.password).toBeUndefined();
    expect(accounts[0].user).toBe('user@example.com');
    expect(accounts[0].smtp?.host).toBe('smtp.example.com');
  });

  it('GET /api/accounts strips OAuth tokens but keeps the public OAuth fields', async () => {
    const res = await fetch(`${baseUrl}/api/accounts`);
    const raw = await res.text();

    expect(raw).not.toContain(REFRESH_SECRET);
    expect(raw).not.toContain(ACCESS_SECRET);

    const outlook = JSON.parse(raw)[1];
    expect(outlook.authType).toBe('oauth2');
    expect(outlook.oauth.refreshToken).toBeUndefined();
    expect(outlook.oauth.accessToken).toBeUndefined();
    expect(outlook.oauth.clientId).toBe('client-123');
    expect(outlook.oauth.tenant).toBe('consumers');
    expect(outlook.oauth.scopes).toEqual(['https://outlook.office.com/IMAP.AccessAsUser.All']);
  });

  it('POST /api/accounts refuses a password login for an OAuth-only provider', async () => {
    const post = (body: Record<string, unknown>) => fetch(`${baseUrl}/api/accounts`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body),
    });

    const byEmail = await post({ name: 'X', email: 'someone@hotmail.com', password: 'pw' });
    expect(byEmail.status).toBe(400);
    expect((await byEmail.json()).error).toMatch(/imap_add_oauth_account/);

    const byHost = await post({ name: 'X', email: 'someone@contoso.com', password: 'pw', host: 'outlook.office365.com' });
    expect(byHost.status).toBe(400);

    const other = await post({ name: 'X', email: 'someone@example.com', password: 'pw', host: 'imap.example.com' });
    expect(other.status).toBe(200);
    expect((await other.json()).success).toBe(true);
  });

  it('GET /api/accounts/:id also stays password-free', async () => {
    const res = await fetch(`${baseUrl}/api/accounts/a1`);
    const raw = await res.text();
    expect(raw).not.toContain(SECRET);
    expect(raw).not.toContain(SMTP_SECRET);
    const body = JSON.parse(raw);
    expect(body.account.password).toBeUndefined();
    expect(body.account.smtp?.password).toBeUndefined();
  });
});
