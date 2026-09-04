import { describe, it, expect, vi, beforeEach } from 'vitest';
import {
  MicrosoftOAuthService,
  MICROSOFT_OAUTH_SCOPES,
  MICROSOFT_GRAPH_RULES_SCOPES,
  ConsentRequiredError,
  allScopesGranted,
  mergeScopes,
  resolveMicrosoftClientId,
  scopeGranted,
  isValidTenant,
} from '../src/services/oauth-service.js';
import type { ImapAccount } from '../src/types/index.js';

/**
 * Drives the device-code flow and token refresh against a scripted `fetch`.
 * No network: every response is queued up front, and the service's sleep is a
 * no-op so `slow_down` / `authorization_pending` loops finish instantly.
 */
type Scripted = { status: number; body: Record<string, unknown> };

function scriptedFetch(responses: Scripted[]) {
  const calls: Array<{ url: string; params: URLSearchParams }> = [];
  const impl = vi.fn(async (input: string | URL | Request, init?: RequestInit) => {
    const next = responses.shift();
    if (!next) throw new Error('fetch called more times than scripted');
    calls.push({ url: String(input), params: new URLSearchParams(String(init?.body ?? '')) });
    return {
      ok: next.status >= 200 && next.status < 300,
      status: next.status,
      json: async () => next.body,
    } as unknown as Response;
  });
  return { impl: impl as unknown as typeof fetch, calls };
}

const deviceCodeOk = {
  status: 200,
  body: {
    device_code: 'DEVICE-SECRET',
    user_code: 'ABCD1234',
    verification_uri: 'https://microsoft.com/devicelogin',
    expires_in: 900,
    interval: 5,
    message: 'go sign in',
  },
};

const tokensOk = {
  status: 200,
  body: {
    access_token: 'ACCESS-1',
    refresh_token: 'REFRESH-1',
    expires_in: 3600,
    scope: 'https://outlook.office.com/IMAP.AccessAsUser.All https://outlook.office.com/SMTP.Send',
  },
};

const pending = { status: 400, body: { error: 'authorization_pending' } };

function service(responses: Scripted[], opts: { now?: () => number } = {}) {
  const { impl, calls } = scriptedFetch(responses);
  const sleeps: number[] = [];
  let clock = 1_000_000;
  const svc = new MicrosoftOAuthService(undefined, {
    fetch: impl,
    sleep: async (ms) => { sleeps.push(ms); clock += ms; },
    now: opts.now ?? (() => clock),
  });
  return { svc, calls, sleeps, advance: (ms: number) => { clock += ms; } };
}

const oauthAccount = (overrides: Partial<ImapAccount['oauth']> = {}): ImapAccount => ({
  id: 'acc-ms',
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
    refreshToken: 'REFRESH-OLD',
    scopes: MICROSOFT_OAUTH_SCOPES,
    ...overrides,
  },
});

describe('MicrosoftOAuthService.startDeviceCode', () => {
  it('posts to the tenant devicecode endpoint with the IMAP/SMTP scopes and returns the user code, never the device code', async () => {
    const { svc, calls } = service([deviceCodeOk]);

    const start = await svc.startDeviceCode({ clientId: 'client-123', tenant: 'consumers' });

    expect(calls[0].url).toBe('https://login.microsoftonline.com/consumers/oauth2/v2.0/devicecode');
    expect(calls[0].params.get('client_id')).toBe('client-123');
    expect(calls[0].params.get('scope')).toBe(MICROSOFT_OAUTH_SCOPES.join(' '));

    expect(start.userCode).toBe('ABCD1234');
    expect(start.verificationUri).toBe('https://microsoft.com/devicelogin');
    expect(start.interval).toBe(5);
    expect(start.flowId).toMatch(/^[0-9a-f]{32}$/);
    expect(JSON.stringify(start)).not.toContain('DEVICE-SECRET');
    expect(svc.hasPendingFlow(start.flowId)).toBe(true);
  });

  it('defaults the tenant to consumers', async () => {
    const { svc, calls } = service([deviceCodeOk]);
    await svc.startDeviceCode({ clientId: 'client-123' });
    expect(calls[0].url).toContain('/consumers/');
  });

  it('rejects a tenant that could escape the URL path', async () => {
    const { svc } = service([]);
    await expect(svc.startDeviceCode({ clientId: 'c', tenant: '../evil' })).rejects.toThrow(/Invalid Entra tenant/);
    expect(isValidTenant('contoso.onmicrosoft.com')).toBe(true);
    expect(isValidTenant('9188040d-6c67-4c5b-b112-36a304b66dad')).toBe(true);
    expect(isValidTenant('a/b')).toBe(false);
  });

  it('surfaces the Microsoft error description when the request fails', async () => {
    const { svc } = service([{ status: 400, body: { error: 'unauthorized_client', error_description: 'AADSTS7000218: public client flows disabled' } }]);
    await expect(svc.startDeviceCode({ clientId: 'client-123' })).rejects.toThrow(/AADSTS7000218/);
  });
});

describe('MicrosoftOAuthService.pollDeviceCode', () => {
  it('reports pending while authorization_pending and completes once tokens arrive', async () => {
    const { svc, calls } = service([deviceCodeOk, pending, pending, tokensOk]);
    const { flowId } = await svc.startDeviceCode({ clientId: 'client-123', context: { name: 'ctx' } });

    // First call: two pendings, then the deadline is reached → pending.
    const first = await svc.pollDeviceCode(flowId, { maxWaitMs: 8_000 });
    expect(first.status).toBe('pending');
    if (first.status === 'pending') expect(first.retryAfterSeconds).toBe(5);

    const second = await svc.pollDeviceCode<{ name: string }>(flowId, { maxWaitMs: 8_000 });
    expect(second.status).toBe('complete');
    if (second.status === 'complete') {
      expect(second.tokens.accessToken).toBe('ACCESS-1');
      expect(second.tokens.refreshToken).toBe('REFRESH-1');
      expect(second.tokens.scopes).toEqual([
        'https://outlook.office.com/IMAP.AccessAsUser.All',
        'https://outlook.office.com/SMTP.Send',
      ]);
      expect(second.context).toEqual({ name: 'ctx' });
    }

    // Token polls carry the device code and the device_code grant.
    const tokenCalls = calls.slice(1);
    expect(tokenCalls.every(c => c.url.endsWith('/consumers/oauth2/v2.0/token'))).toBe(true);
    expect(tokenCalls[0].params.get('grant_type')).toBe('urn:ietf:params:oauth:grant-type:device_code');
    expect(tokenCalls[0].params.get('device_code')).toBe('DEVICE-SECRET');

    // Terminal result removes the flow.
    expect(svc.hasPendingFlow(flowId)).toBe(false);
  });

  it('backs off by 5 seconds on slow_down', async () => {
    const { svc, sleeps } = service([deviceCodeOk, { status: 400, body: { error: 'slow_down' } }, pending, tokensOk]);
    const { flowId } = await svc.startDeviceCode({ clientId: 'client-123' });

    const result = await svc.pollDeviceCode(flowId, { maxWaitMs: 60_000 });
    expect(result.status).toBe('complete');
    // 5s interval, then slow_down bumps it to 10s.
    expect(sleeps).toEqual([10_000, 10_000]);
  });

  it('returns expired when Microsoft says expired_token', async () => {
    const { svc } = service([deviceCodeOk, { status: 400, body: { error: 'expired_token' } }]);
    const { flowId } = await svc.startDeviceCode({ clientId: 'client-123' });
    const result = await svc.pollDeviceCode(flowId);
    expect(result.status).toBe('expired');
    expect(svc.hasPendingFlow(flowId)).toBe(false);
  });

  it('returns expired without polling once the local expiry has passed', async () => {
    const { svc, calls, advance } = service([deviceCodeOk]);
    const { flowId } = await svc.startDeviceCode({ clientId: 'client-123' });
    advance(901_000);
    const result = await svc.pollDeviceCode(flowId);
    expect(result.status).toBe('expired');
    expect(calls).toHaveLength(1);
  });

  it('returns denied when the user declines', async () => {
    const { svc } = service([deviceCodeOk, { status: 400, body: { error: 'authorization_declined' } }]);
    const { flowId } = await svc.startDeviceCode({ clientId: 'client-123' });
    expect((await svc.pollDeviceCode(flowId)).status).toBe('denied');
  });

  it('returns an error for an unknown flowId', async () => {
    const { svc } = service([]);
    const result = await svc.pollDeviceCode('nope');
    expect(result.status).toBe('error');
    if (result.status === 'error') expect(result.error).toMatch(/Unknown flowId/);
  });
});

describe('MicrosoftOAuthService.refreshAccessToken', () => {
  it('exchanges the refresh token and returns the rotated one', async () => {
    const { svc, calls } = service([{ status: 200, body: { access_token: 'ACCESS-2', refresh_token: 'REFRESH-2', expires_in: 3600 } }]);

    const result = await svc.refreshAccessToken(oauthAccount());

    expect(calls[0].url).toBe('https://login.microsoftonline.com/consumers/oauth2/v2.0/token');
    expect(calls[0].params.get('grant_type')).toBe('refresh_token');
    expect(calls[0].params.get('refresh_token')).toBe('REFRESH-OLD');
    expect(calls[0].params.get('client_id')).toBe('client-123');
    expect(result.accessToken).toBe('ACCESS-2');
    expect(result.refreshToken).toBe('REFRESH-2');
    expect(result.accessTokenExpiresAt).toBeGreaterThan(0);
  });

  it('tells the user to re-authorize on invalid_grant, naming the tool and account', async () => {
    const { svc } = service([{ status: 400, body: { error: 'invalid_grant', error_description: 'AADSTS70000: token expired' } }]);
    await expect(svc.refreshAccessToken(oauthAccount())).rejects.toThrow(/imap_add_oauth_account.*acc-ms/);
  });

  it('refuses to refresh without a refresh token', async () => {
    const { svc } = service([]);
    await expect(svc.refreshAccessToken(oauthAccount({ refreshToken: '' }))).rejects.toThrow(/no OAuth refresh token/);
  });
});

describe('MicrosoftOAuthService.getValidAccessToken', () => {
  let updateOAuthTokens: ReturnType<typeof vi.fn>;
  let fakeManager: { updateOAuthTokens: ReturnType<typeof vi.fn> };

  beforeEach(() => {
    updateOAuthTokens = vi.fn(async () => undefined);
    fakeManager = { updateOAuthTokens };
  });

  it('returns the cached token when it has more than five minutes left', async () => {
    const { svc, calls } = service([]);
    svc.setAccountManager(fakeManager as any);
    const account = oauthAccount({ accessToken: 'CACHED', accessTokenExpiresAt: 1_000_000 + 10 * 60 * 1000 });

    expect(await svc.getValidAccessToken(account)).toBe('CACHED');
    expect(calls).toHaveLength(0);
    expect(updateOAuthTokens).not.toHaveBeenCalled();
  });

  it('refreshes, mutates the account, and persists when the cached token is about to expire', async () => {
    const { svc } = service([{ status: 200, body: { access_token: 'ACCESS-3', refresh_token: 'REFRESH-3', expires_in: 3600 } }]);
    svc.setAccountManager(fakeManager as any);
    const account = oauthAccount({ accessToken: 'STALE', accessTokenExpiresAt: 1_000_000 + 60 * 1000 });

    expect(await svc.getValidAccessToken(account)).toBe('ACCESS-3');
    expect(account.oauth?.accessToken).toBe('ACCESS-3');
    expect(account.oauth?.refreshToken).toBe('REFRESH-3');
    expect(updateOAuthTokens).toHaveBeenCalledWith('acc-ms', expect.objectContaining({
      accessToken: 'ACCESS-3',
      refreshToken: 'REFRESH-3',
    }));
  });

  it('shares one refresh between concurrent callers', async () => {
    const { svc, calls } = service([{ status: 200, body: { access_token: 'ACCESS-4', expires_in: 3600 } }]);
    const account = oauthAccount();
    const [a, b] = await Promise.all([svc.getValidAccessToken(account), svc.getValidAccessToken(account)]);
    expect(a).toBe('ACCESS-4');
    expect(b).toBe('ACCESS-4');
    expect(calls).toHaveLength(1);
  });

  it('rejects password accounts', async () => {
    const { svc } = service([]);
    await expect(svc.getValidAccessToken({ ...oauthAccount(), authType: 'password' })).rejects.toThrow(/not configured for OAuth/);
  });
});

describe('resolveMicrosoftClientId', () => {
  it('prefers the explicit argument, then the env var, and explains when neither exists', () => {
    expect(resolveMicrosoftClientId('explicit', { IMAP_MCP_MS_CLIENT_ID: 'env' })).toBe('explicit');
    expect(resolveMicrosoftClientId(undefined, { IMAP_MCP_MS_CLIENT_ID: 'env' })).toBe('env');
    expect(() => resolveMicrosoftClientId(undefined, {})).toThrow(/IMAP_MCP_MS_CLIENT_ID/);
  });
});

describe('MicrosoftOAuthService — Graph (second-resource) scopes', () => {
  let updateOAuthTokens: ReturnType<typeof vi.fn>;

  beforeEach(() => {
    updateOAuthTokens = vi.fn(async () => undefined);
  });

  it('mints a Graph token with the Graph scopes, keeps it in memory only, and persists a rotated refresh token', async () => {
    const { svc, calls } = service([
      { status: 200, body: { access_token: 'GRAPH-1', refresh_token: 'REFRESH-ROTATED', expires_in: 3600, scope: 'MailboxSettings.ReadWrite Mail.ReadBasic' } },
    ]);
    svc.setAccountManager({ updateOAuthTokens } as any);
    const account = oauthAccount({ accessToken: 'MAIL-CACHED', accessTokenExpiresAt: 1_000_000 + 60 * 60 * 1000 });

    const token = await svc.getValidAccessToken(account, MICROSOFT_GRAPH_RULES_SCOPES);

    expect(token).toBe('GRAPH-1');
    expect(calls[0].params.get('grant_type')).toBe('refresh_token');
    expect(calls[0].params.get('scope')).toBe(MICROSOFT_GRAPH_RULES_SCOPES.join(' '));
    // The persisted mail access token is untouched; the refresh token rotates for both.
    expect(account.oauth?.accessToken).toBe('MAIL-CACHED');
    expect(account.oauth?.refreshToken).toBe('REFRESH-ROTATED');
    expect(updateOAuthTokens).toHaveBeenCalledTimes(1);
    expect(updateOAuthTokens).toHaveBeenCalledWith('acc-ms', { refreshToken: 'REFRESH-ROTATED' });

    // Second call for the same scope set is served from the in-memory cache.
    expect(await svc.getValidAccessToken(account, MICROSOFT_GRAPH_RULES_SCOPES)).toBe('GRAPH-1');
    expect(calls).toHaveLength(1);
    // The mail token cache is independent.
    expect(await svc.getValidAccessToken(account)).toBe('MAIL-CACHED');
    expect(calls).toHaveLength(1);
  });

  it('does not write the store when a Graph refresh returns no rotated refresh token', async () => {
    const { svc } = service([{ status: 200, body: { access_token: 'GRAPH-2', expires_in: 3600 } }]);
    svc.setAccountManager({ updateOAuthTokens } as any);
    await svc.getValidAccessToken(oauthAccount(), MICROSOFT_GRAPH_RULES_SCOPES);
    expect(updateOAuthTokens).not.toHaveBeenCalled();
  });

  it('refreshes a Graph token that is about to expire, sharing one refresh between concurrent callers', async () => {
    const { svc, calls, advance } = service([
      { status: 200, body: { access_token: 'GRAPH-A', expires_in: 3600 } },
      { status: 200, body: { access_token: 'GRAPH-B', expires_in: 3600 } },
    ]);
    const account = oauthAccount();
    await svc.getValidAccessToken(account, MICROSOFT_GRAPH_RULES_SCOPES);
    advance(3600 * 1000 - 60 * 1000); // one minute left → inside the five-minute margin
    const [a, b] = await Promise.all([
      svc.getValidAccessToken(account, MICROSOFT_GRAPH_RULES_SCOPES),
      svc.getValidAccessToken(account, MICROSOFT_GRAPH_RULES_SCOPES),
    ]);
    expect(a).toBe('GRAPH-B');
    expect(b).toBe('GRAPH-B');
    expect(calls).toHaveLength(2);
  });

  it('throws ConsentRequiredError (not the re-authorize-the-mailbox error) when Graph consent is missing', async () => {
    const { svc } = service([
      { status: 400, body: { error: 'invalid_grant', error_description: 'AADSTS65001: The user or administrator has not consented to use the application' } },
    ]);
    const err = await svc.getValidAccessToken(oauthAccount(), MICROSOFT_GRAPH_RULES_SCOPES).catch(e => e);
    expect(err).toBeInstanceOf(ConsentRequiredError);
    expect(err.code).toBe('consent_required');
    expect(err.accountId).toBe('acc-ms');
    expect(err.scopes).toEqual(MICROSOFT_GRAPH_RULES_SCOPES);
    expect(err.message).toMatch(/AADSTS65001/);
    expect(err.message).not.toMatch(/imap_add_oauth_account/);
  });

  it('still maps invalid_grant on the mail scopes to the mailbox re-authorization error', async () => {
    const { svc } = service([{ status: 400, body: { error: 'invalid_grant', error_description: 'AADSTS70000' } }]);
    const err = await svc.getValidAccessToken(oauthAccount()).catch(e => e);
    expect(err).not.toBeInstanceOf(ConsentRequiredError);
    expect(err.message).toMatch(/imap_add_oauth_account/);
  });

  it('primeAccessToken seeds the Graph cache but never the mail token; forgetResourceTokens clears it', async () => {
    const { svc, calls } = service([{ status: 200, body: { access_token: 'GRAPH-FRESH', expires_in: 3600 } }]);
    const account = oauthAccount({ accessToken: 'MAIL', accessTokenExpiresAt: 1_000_000 + 3_600_000 });

    svc.primeAccessToken(account, MICROSOFT_GRAPH_RULES_SCOPES, 'GRAPH-PRIMED', 1_000_000 + 3_600_000);
    svc.primeAccessToken(account, MICROSOFT_OAUTH_SCOPES, 'SHOULD-BE-IGNORED', 1_000_000 + 3_600_000);

    expect(await svc.getValidAccessToken(account, MICROSOFT_GRAPH_RULES_SCOPES)).toBe('GRAPH-PRIMED');
    expect(await svc.getValidAccessToken(account)).toBe('MAIL');
    expect(calls).toHaveLength(0);

    svc.forgetResourceTokens('acc-ms');
    expect(await svc.getValidAccessToken(account, MICROSOFT_GRAPH_RULES_SCOPES)).toBe('GRAPH-FRESH');
    expect(calls).toHaveLength(1);
  });

  it('starts a device-code flow for the Graph scopes when asked', async () => {
    const { svc, calls } = service([deviceCodeOk]);
    await svc.startDeviceCode({ clientId: 'client-123', scopes: MICROSOFT_GRAPH_RULES_SCOPES, context: { kind: 'graph-consent' } });
    expect(calls[0].params.get('scope')).toBe(MICROSOFT_GRAPH_RULES_SCOPES.join(' '));
  });
});

describe('scope helpers', () => {
  it('matches the full URI or the short name Microsoft reports for Graph scopes, case-insensitively', () => {
    expect(scopeGranted(['MailboxSettings.ReadWrite', 'Mail.ReadBasic'], 'https://graph.microsoft.com/MailboxSettings.ReadWrite')).toBe(true);
    expect(scopeGranted(['https://graph.microsoft.com/mailboxsettings.readwrite'], 'https://graph.microsoft.com/MailboxSettings.ReadWrite')).toBe(true);
    expect(scopeGranted(['https://outlook.office.com/IMAP.AccessAsUser.All'], 'https://graph.microsoft.com/Mail.ReadBasic')).toBe(false);
    expect(scopeGranted(undefined, 'x')).toBe(false);
  });

  it('allScopesGranted ignores offline_access and needs every other scope', () => {
    expect(allScopesGranted(['MailboxSettings.ReadWrite', 'Mail.ReadBasic'], MICROSOFT_GRAPH_RULES_SCOPES)).toBe(true);
    expect(allScopesGranted(['MailboxSettings.ReadWrite'], MICROSOFT_GRAPH_RULES_SCOPES)).toBe(false);
    expect(allScopesGranted(MICROSOFT_OAUTH_SCOPES, MICROSOFT_GRAPH_RULES_SCOPES)).toBe(false);
  });

  it('mergeScopes unions case-insensitively, first spelling wins', () => {
    expect(mergeScopes(['A', 'b'], ['a', 'C'], undefined)).toEqual(['A', 'b', 'C']);
  });
});
