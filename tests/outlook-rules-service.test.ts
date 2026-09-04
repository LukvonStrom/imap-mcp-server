import { describe, it, expect, vi } from 'vitest';
import {
  OutlookRulesService,
  GraphApiError,
  splitFolderPath,
  wellKnownFolderName,
} from '../src/services/outlook-rules-service.js';
import { ConsentRequiredError, MICROSOFT_GRAPH_RULES_SCOPES } from '../src/services/oauth-service.js';
import type { ImapAccount } from '../src/types/index.js';

/**
 * The Graph client against a scripted `fetch`: every response is matched by
 * method + path (query string included), so a test states exactly which
 * requests it expects. Tokens come from a fake OAuth service.
 */
type Scripted = { status: number; body?: unknown; headers?: Record<string, string> };
type Call = { method: string; path: string; body?: any; auth?: string };

function graph(script: Record<string, Scripted[]>, opts: { token?: string; refreshed?: string } = {}) {
  const calls: Call[] = [];
  const fetchImpl = vi.fn(async (input: string | URL | Request, init?: RequestInit) => {
    const url = String(input);
    expect(url.startsWith('https://graph.microsoft.com/v1.0')).toBe(true);
    const path = url.slice('https://graph.microsoft.com/v1.0'.length);
    const method = init?.method ?? 'GET';
    const headers = init?.headers as Record<string, string>;
    calls.push({ method, path, body: init?.body ? JSON.parse(String(init.body)) : undefined, auth: headers?.Authorization });
    const key = `${method} ${path}`;
    const next = script[key]?.shift();
    if (!next) throw new Error(`unscripted request: ${key}`);
    const text = next.body === undefined ? '' : JSON.stringify(next.body);
    return {
      ok: next.status >= 200 && next.status < 300,
      status: next.status,
      statusText: '',
      headers: { get: (name: string) => next.headers?.[name] ?? null },
      text: async () => text,
    } as unknown as Response;
  });

  const oauth = {
    getValidAccessToken: vi.fn(async (_a: ImapAccount, scopes: string[]) => {
      expect(scopes).toEqual(MICROSOFT_GRAPH_RULES_SCOPES);
      return opts.token ?? 'GRAPH-TOKEN';
    }),
    forceRefresh: vi.fn(async () => opts.refreshed ?? 'GRAPH-TOKEN-2'),
  };
  const sleeps: number[] = [];
  const svc = new OutlookRulesService(oauth as any, {
    fetch: fetchImpl as unknown as typeof fetch,
    sleep: async (ms) => { sleeps.push(ms); },
  });
  return { svc, calls, oauth, sleeps };
}

const account: ImapAccount = {
  id: 'acc-ms', name: 'Outlook', host: 'outlook.office365.com', port: 993, user: 'me@outlook.com', password: '', tls: true,
  authType: 'oauth2',
  oauth: { provider: 'microsoft', clientId: 'c', tenant: 'consumers', refreshToken: 'r', scopes: [] },
};

const FOLDER_QUERY = '?$top=250&$select=id,displayName,parentFolderId,childFolderCount';
const roots = [
  { id: 'ID-INBOX', displayName: 'Inbox', childFolderCount: 1 },
  { id: 'ID-NEWS', displayName: 'Newsletters', childFolderCount: 0 },
  { id: 'ID-JUNK', displayName: 'Junk Email', childFolderCount: 0 },
];
const inboxChildren = [{ id: 'ID-PAYPAL', displayName: 'Paypal', parentFolderId: 'ID-INBOX', childFolderCount: 0 }];

describe('OutlookRulesService — rules', () => {
  it('lists rules with a Graph-scoped Bearer token', async () => {
    const { svc, calls, oauth } = graph({
      'GET /me/mailFolders/inbox/messageRules': [{ status: 200, body: { value: [{ id: 'R1', displayName: 'a', sequence: 1, isEnabled: true }] } }],
    });
    const rules = await svc.listRules(account);
    expect(rules).toEqual([{ id: 'R1', displayName: 'a', sequence: 1, isEnabled: true }]);
    expect(calls[0].auth).toBe('Bearer GRAPH-TOKEN');
    expect(oauth.getValidAccessToken).toHaveBeenCalledWith(account, MICROSOFT_GRAPH_RULES_SCOPES);
  });

  it('creates, patches, and deletes by id (URL-encoded), sending JSON bodies', async () => {
    const { svc, calls } = graph({
      'POST /me/mailFolders/inbox/messageRules': [{ status: 201, body: { id: 'AQAAAJ5dZqA=', displayName: 'x', sequence: 1, isEnabled: true } }],
      'PATCH /me/mailFolders/inbox/messageRules/AQAAAJ5dZqA%3D': [{ status: 200, body: { id: 'AQAAAJ5dZqA=', displayName: 'y', sequence: 1, isEnabled: false } }],
      'DELETE /me/mailFolders/inbox/messageRules/AQAAAJ5dZqA%3D': [{ status: 204 }],
    });
    const created = await svc.createRule(account, { displayName: 'x', sequence: 1, isEnabled: true, conditions: { senderContains: ['@a'] }, actions: { markAsRead: true } });
    expect(created.id).toBe('AQAAAJ5dZqA=');
    expect(calls[0].body).toEqual({ displayName: 'x', sequence: 1, isEnabled: true, conditions: { senderContains: ['@a'] }, actions: { markAsRead: true } });

    const updated = await svc.updateRule(account, 'AQAAAJ5dZqA=', { displayName: 'y', isEnabled: false });
    expect(updated.isEnabled).toBe(false);
    expect(calls[1].body).toEqual({ displayName: 'y', isEnabled: false });

    await svc.deleteRule(account, 'AQAAAJ5dZqA=');
    expect(calls[2].method).toBe('DELETE');
  });

  it('retries once with a forced refresh on 401, then maps a repeated 401/403 to ConsentRequiredError', async () => {
    const { svc, calls, oauth } = graph({
      'GET /me/mailFolders/inbox/messageRules': [
        { status: 401, body: { error: { code: 'InvalidAuthenticationToken', message: 'expired' } } },
        { status: 200, body: { value: [] } },
        { status: 403, body: { error: { code: 'ErrorAccessDenied', message: 'Access is denied' } } },
      ],
    });
    expect(await svc.listRules(account)).toEqual([]);
    expect(oauth.forceRefresh).toHaveBeenCalledWith(account, MICROSOFT_GRAPH_RULES_SCOPES);
    expect(calls[1].auth).toBe('Bearer GRAPH-TOKEN-2');

    const err = await svc.listRules(account).catch(e => e);
    expect(err).toBeInstanceOf(ConsentRequiredError);
    expect(err.message).toMatch(/ErrorAccessDenied/);
    expect(err.message).toMatch(/MailboxSettings.ReadWrite/);
  });

  it('honours Retry-After once on 429 and then gives up with the Graph error', async () => {
    const { svc, sleeps } = graph({
      'GET /me/mailFolders/inbox/messageRules': [
        { status: 429, headers: { 'Retry-After': '3' }, body: { error: { code: 'TooManyRequests', message: 'slow down' } } },
        { status: 200, body: { value: [] } },
        { status: 429, headers: { 'Retry-After': '1' } },
        { status: 429, headers: { 'Retry-After': '1' }, body: { error: { code: 'TooManyRequests', message: 'still' } } },
      ],
    });
    expect(await svc.listRules(account)).toEqual([]);
    expect(sleeps).toEqual([3000]);

    const err = await svc.listRules(account).catch(e => e);
    expect(err).toBeInstanceOf(GraphApiError);
    expect(err.status).toBe(429);
    expect(err.code).toBe('TooManyRequests');
  });

  it('surfaces Graph error code and message for other failures', async () => {
    const { svc } = graph({
      'GET /me/mailFolders/inbox/messageRules/nope': [{ status: 404, body: { error: { code: 'ErrorItemNotFound', message: 'The specified object was not found in the store.' } } }],
    });
    const err = await svc.getRule(account, 'nope').catch(e => e);
    expect(err).toBeInstanceOf(GraphApiError);
    expect(err.code).toBe('ErrorItemNotFound');
    expect(err.message).toMatch(/404 ErrorItemNotFound: The specified object/);
  });
});

describe('OutlookRulesService — folders', () => {
  it('resolves a well-known first segment directly and walks child folders per segment, case-insensitively', async () => {
    const { svc, calls } = graph({
      'GET /me/mailFolders/inbox': [{ status: 200, body: { id: 'ID-INBOX', displayName: 'Inbox', childFolderCount: 1 } }],
      [`GET /me/mailFolders/ID-INBOX/childFolders${FOLDER_QUERY}`]: [{ status: 200, body: { value: inboxChildren } }],
    });
    const { folder, created } = await svc.resolveFolder(account, 'inbox/PAYPAL');
    expect(folder).toMatchObject({ id: 'ID-PAYPAL', path: 'Inbox/Paypal' });
    expect(created).toEqual([]);
    expect(calls.map(c => c.path)).toEqual(['/me/mailFolders/inbox', `/me/mailFolders/ID-INBOX/childFolders${FOLDER_QUERY}`]);
  });

  it('maps user spellings to Graph well-known names', () => {
    expect(wellKnownFolderName('Junk Email')).toBe('junkemail');
    expect(wellKnownFolderName('spam')).toBe('junkemail');
    expect(wellKnownFolderName('Deleted Items')).toBe('deleteditems');
    expect(wellKnownFolderName('Trash')).toBe('deleteditems');
    expect(wellKnownFolderName('Sent Items')).toBe('sentitems');
    expect(wellKnownFolderName('Archive')).toBe('archive');
    expect(wellKnownFolderName('Newsletters')).toBeUndefined();
    expect(splitFolderPath('/Inbox//Paypal/ ')).toEqual(['Inbox', 'Paypal']);
  });

  it('looks a top-level user folder up among the root folders (siblings of Inbox)', async () => {
    const { svc } = graph({
      [`GET /me/mailFolders${FOLDER_QUERY}`]: [{ status: 200, body: { value: roots } }],
    });
    const { folder } = await svc.resolveFolder(account, 'newsletters');
    expect(folder).toMatchObject({ id: 'ID-NEWS', path: 'Newsletters' });
  });

  it('errors with guidance for an unknown folder unless create is requested, then creates missing segments', async () => {
    const { svc, calls } = graph({
      [`GET /me/mailFolders${FOLDER_QUERY}`]: [{ status: 200, body: { value: roots } }, { status: 200, body: { value: roots } }],
      'POST /me/mailFolders': [{ status: 201, body: { id: 'ID-SHOP', displayName: 'Shopping', childFolderCount: 0 } }],
      [`GET /me/mailFolders/ID-SHOP/childFolders${FOLDER_QUERY}`]: [{ status: 200, body: { value: [] } }],
      'POST /me/mailFolders/ID-SHOP/childFolders': [{ status: 201, body: { id: 'ID-AMZ', displayName: 'Amazon', parentFolderId: 'ID-SHOP', childFolderCount: 0 } }],
    });
    await expect(svc.resolveFolder(account, 'Shopping/Amazon')).rejects.toThrow(/Folder "Shopping" not found at the top level.*createFolder: true/);

    const { folder, created } = await svc.resolveFolder(account, 'Shopping/Amazon', { create: true });
    expect(folder).toMatchObject({ id: 'ID-AMZ', path: 'Shopping/Amazon' });
    expect(created).toEqual(['Shopping', 'Shopping/Amazon']);
    expect(calls.find(c => c.method === 'POST' && c.path === '/me/mailFolders')?.body).toEqual({ displayName: 'Shopping' });
    expect(calls.find(c => c.method === 'POST' && c.path === '/me/mailFolders/ID-SHOP/childFolders')?.body).toEqual({ displayName: 'Amazon' });
  });

  it('listAllFolders walks the hierarchy once and builds display paths, following nextLink pages', async () => {
    const { svc, calls } = graph({
      [`GET /me/mailFolders${FOLDER_QUERY}`]: [{
        status: 200,
        body: { value: roots.slice(0, 2), '@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/mailFolders?$skip=2' },
      }],
      'GET /me/mailFolders?$skip=2': [{ status: 200, body: { value: roots.slice(2) } }],
      [`GET /me/mailFolders/ID-INBOX/childFolders${FOLDER_QUERY}`]: [{ status: 200, body: { value: inboxChildren } }],
    });
    const folders = await svc.listAllFolders(account);
    expect(folders.map(f => f.path)).toEqual(['Inbox', 'Inbox/Paypal', 'Newsletters', 'Junk Email']);
    expect(calls).toHaveLength(3);
  });

  it('rejects an empty path', async () => {
    const { svc } = graph({});
    await expect(svc.resolveFolder(account, ' / ')).rejects.toThrow(/Folder path is empty/);
  });
});
