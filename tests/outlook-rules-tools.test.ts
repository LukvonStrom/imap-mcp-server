import { describe, it, expect, vi, beforeEach } from 'vitest';
import { outlookRulesTools, buildActions, buildPredicates, hasGraphConsent } from '../src/tools/outlook-rules-tools.js';
import { GraphApiError } from '../src/services/outlook-rules-service.js';
import { ConsentRequiredError, MICROSOFT_GRAPH_RULES_SCOPES } from '../src/services/oauth-service.js';
import type { ImapAccount } from '../src/types/index.js';

/**
 * imap_outlook_* against a scripted rules service and OAuth service — no
 * network. Covers the account guard, the consent gate, the refusal rules, the
 * Graph payload shape, and that no token ever reaches the output.
 */
const handlers: Record<string, Function> = {};
const mockServer = {
  registerTool: vi.fn((name: string, _schema: any, handler: Function) => { handlers[name] = handler; }),
};

const MAIL_SCOPES = ['https://outlook.office.com/IMAP.AccessAsUser.All', 'https://outlook.office.com/SMTP.Send'];

function account(overrides: Partial<ImapAccount> = {}, oauth: Partial<NonNullable<ImapAccount['oauth']>> = {}): ImapAccount {
  return {
    id: 'acc-ms', name: 'Outlook', host: 'outlook.office365.com', port: 993, user: 'me@outlook.com', email: 'me@outlook.com',
    password: '', tls: true, authType: 'oauth2',
    oauth: {
      provider: 'microsoft', clientId: 'client-123', tenant: 'consumers',
      refreshToken: 'REFRESH-SECRET', accessToken: 'ACCESS-SECRET', accessTokenExpiresAt: 1,
      scopes: MAIL_SCOPES,
      grantedScopes: [...MAIL_SCOPES, 'MailboxSettings.ReadWrite', 'Mail.ReadBasic'],
      ...oauth,
    },
    ...overrides,
  };
}

let accounts: ImapAccount[] = [];
const mockAccountManager = {
  resolveAccountId: vi.fn((id?: string, name?: string) => {
    const match = accounts.find(a => (id && a.id === id) || (name && a.name === name)) ?? (accounts.length === 1 ? accounts[0] : undefined);
    if (!match) throw new Error(`Account ${id ?? name} not found`);
    return match.id;
  }),
  getAccount: vi.fn((id: string) => accounts.find(a => a.id === id)),
};

const mockOAuthService = {
  startDeviceCode: vi.fn(async () => ({
    flowId: 'flow-graph', userCode: 'WXYZ9876', verificationUri: 'https://microsoft.com/devicelogin',
    expiresAt: 1_800_000_000_000, message: 'go', interval: 5,
  })),
};

const existingRule = {
  id: 'R1', displayName: 'LinkedIn', sequence: 1, isEnabled: true,
  conditions: { senderContains: ['@linkedin.com'] },
  actions: { moveToFolder: 'ID-NEWS', stopProcessingRules: true },
  exceptions: null,
};

const mockRulesService = {
  listRules: vi.fn(async () => [existingRule]),
  getRule: vi.fn(async () => existingRule),
  createRule: vi.fn(async (_a: ImapAccount, body: any) => ({ id: 'R-NEW', ...body })),
  updateRule: vi.fn(async (_a: ImapAccount, id: string, patch: any) => ({ ...existingRule, id, ...patch })),
  deleteRule: vi.fn(async () => undefined),
  resolveFolder: vi.fn(async (_a: ImapAccount, path: string, opts: any) => {
    if (path.toLowerCase() === 'newsletters') return { folder: { id: 'ID-NEWS', displayName: 'Newsletters', path: 'Newsletters', childFolderCount: 0 }, created: [] };
    if (opts?.create) return { folder: { id: 'ID-CREATED', displayName: path, path, childFolderCount: 0 }, created: [path] };
    throw new Error(`Folder "${path}" not found at the top level. Check the exact name with imap_list_folders, or pass createFolder: true to create it.`);
  }),
  listAllFolders: vi.fn(async () => [
    { id: 'ID-INBOX', displayName: 'Inbox', path: 'Inbox', childFolderCount: 0 },
    { id: 'ID-NEWS', displayName: 'Newsletters', path: 'Newsletters', childFolderCount: 0 },
  ]),
};

const parse = (result: any) => JSON.parse(result.content[0].text);

describe('Outlook rules tools', () => {
  beforeEach(() => {
    vi.clearAllMocks();
    accounts = [account()];
    outlookRulesTools(mockServer as any, mockAccountManager as any, mockOAuthService as any, mockRulesService as any);
  });

  it('registers the five tools', () => {
    expect(Object.keys(handlers).sort()).toEqual([
      'imap_outlook_authorize_rules', 'imap_outlook_create_rule', 'imap_outlook_delete_rule',
      'imap_outlook_list_rules', 'imap_outlook_update_rule',
    ]);
  });

  describe('account guard', () => {
    it.each(['imap_outlook_list_rules', 'imap_outlook_authorize_rules', 'imap_outlook_delete_rule'])(
      '%s rejects a password account with a clear error', async (tool) => {
        accounts = [{ id: 'p1', name: 'Plain', host: 'imap.example.com', port: 993, user: 'u', password: 'pw', tls: true }];
        await expect(handlers[tool]({ accountId: 'p1', ruleId: 'x' })).rejects.toThrow(/not a Microsoft OAuth 2.0 account.*authType "password"/);
        expect(mockRulesService.listRules).not.toHaveBeenCalled();
      },
    );

    it('resolves the account by name and by single-account default', async () => {
      await handlers.imap_outlook_list_rules({ accountName: 'Outlook' });
      await handlers.imap_outlook_list_rules({});
      expect(mockRulesService.listRules).toHaveBeenCalledTimes(2);
    });
  });

  describe('consent gate', () => {
    it('returns graph-consent-required (no Graph call) when grantedScopes lack the Graph scopes', async () => {
      accounts = [account({}, { grantedScopes: MAIL_SCOPES })];
      const result = parse(await handlers.imap_outlook_list_rules({ accountId: 'acc-ms' }));
      expect(result).toMatchObject({
        success: false, error: 'graph-consent-required', accountId: 'acc-ms', nextStep: 'imap_outlook_authorize_rules',
        requiredScopes: ['https://graph.microsoft.com/MailboxSettings.ReadWrite', 'https://graph.microsoft.com/Mail.ReadBasic'],
      });
      expect(mockRulesService.listRules).not.toHaveBeenCalled();
    });

    it('treats a legacy account without grantedScopes as consented only for what scopes lists', () => {
      expect(hasGraphConsent(account({}, { grantedScopes: undefined }))).toBe(false);
      expect(hasGraphConsent(account({}, { grantedScopes: undefined, scopes: [...MAIL_SCOPES, 'MailboxSettings.ReadWrite', 'Mail.ReadBasic'] }))).toBe(true);
    });

    it('maps a ConsentRequiredError from the service to the same structured error', async () => {
      mockRulesService.listRules.mockRejectedValueOnce(new ConsentRequiredError('AADSTS65001 consent missing', MICROSOFT_GRAPH_RULES_SCOPES, 'acc-ms'));
      const result = parse(await handlers.imap_outlook_list_rules({ accountId: 'acc-ms' }));
      expect(result.error).toBe('graph-consent-required');
      expect(result.message).toMatch(/AADSTS65001/);
    });

    it('returns Graph API failures as structured errors', async () => {
      mockRulesService.listRules.mockRejectedValueOnce(new GraphApiError('Microsoft Graph GET x failed: 404 MailboxNotEnabledForRESTAPI: nope', 404, 'MailboxNotEnabledForRESTAPI', 'GET', 'x'));
      const result = parse(await handlers.imap_outlook_list_rules({ accountId: 'acc-ms' }));
      expect(result).toMatchObject({ success: false, error: 'graph-error', status: 404, code: 'MailboxNotEnabledForRESTAPI' });
      expect(result.hint).toMatch(/not reachable through Microsoft Graph/);
    });
  });

  describe('imap_outlook_authorize_rules', () => {
    it('starts a Graph-scoped device-code flow bound to the existing account and never returns tokens', async () => {
      accounts = [account({}, { grantedScopes: MAIL_SCOPES })];
      const raw = await handlers.imap_outlook_authorize_rules({ accountId: 'acc-ms' });
      const result = parse(raw);

      expect(mockOAuthService.startDeviceCode).toHaveBeenCalledWith({
        clientId: 'client-123',
        tenant: 'consumers',
        scopes: MICROSOFT_GRAPH_RULES_SCOPES,
        context: { kind: 'graph-consent', accountId: 'acc-ms', scopes: MICROSOFT_GRAPH_RULES_SCOPES },
      });
      expect(result).toMatchObject({ status: 'awaiting_user', accountId: 'acc-ms', flowId: 'flow-graph', userCode: 'WXYZ9876', verificationUri: 'https://microsoft.com/devicelogin' });
      expect(result.alreadyGranted).toBeUndefined();
      expect(result.instructions).toContain('WXYZ9876');
      expect(result.instructions).toContain('imap_complete_oauth_login');
      expect(raw.content[0].text).not.toContain('SECRET');
    });

    it('flags when consent already exists but still allows re-consenting', async () => {
      const result = parse(await handlers.imap_outlook_authorize_rules({ accountId: 'acc-ms' }));
      expect(result.alreadyGranted).toBe(true);
      expect(result.status).toBe('awaiting_user');
    });
  });

  describe('imap_outlook_list_rules', () => {
    it('returns rules sorted by sequence with resolved folder paths and no tokens', async () => {
      mockRulesService.listRules.mockResolvedValueOnce([
        { id: 'R2', displayName: 'Second', sequence: 2, isEnabled: false, conditions: { subjectContains: ['x'] }, actions: { markAsRead: true }, exceptions: null },
        existingRule,
      ]);
      const raw = await handlers.imap_outlook_list_rules({ accountId: 'acc-ms' });
      const result = parse(raw);

      expect(result.success).toBe(true);
      expect(result.count).toBe(2);
      expect(result.rules.map((r: any) => r.id)).toEqual(['R1', 'R2']);
      expect(result.rules[0]).toEqual({
        id: 'R1', displayName: 'LinkedIn', sequence: 1, isEnabled: true,
        conditions: { senderContains: ['@linkedin.com'] },
        actions: { moveToFolder: 'ID-NEWS', moveToFolderPath: 'Newsletters', stopProcessingRules: true },
        exceptions: {},
      });
      expect(mockRulesService.listAllFolders).toHaveBeenCalledTimes(1);
      expect(raw.content[0].text).not.toMatch(/SECRET|refreshToken|accessToken/);
    });

    it('skips the folder walk when no rule moves mail, and reports when the walk fails', async () => {
      mockRulesService.listRules.mockResolvedValueOnce([{ id: 'R3', displayName: 'x', sequence: 1, isEnabled: true, actions: { markAsRead: true } }]);
      let result = parse(await handlers.imap_outlook_list_rules({ accountId: 'acc-ms' }));
      expect(mockRulesService.listAllFolders).not.toHaveBeenCalled();
      expect(result.folderPathsUnavailable).toBeUndefined();

      mockRulesService.listAllFolders.mockRejectedValueOnce(new Error('Mail.ReadBasic missing'));
      result = parse(await handlers.imap_outlook_list_rules({ accountId: 'acc-ms' }));
      expect(result.folderPathsUnavailable).toBe('Mail.ReadBasic missing');
      expect(result.rules[0].actions.moveToFolderPath).toBeUndefined();
    });
  });

  describe('imap_outlook_create_rule', () => {
    const base = { accountId: 'acc-ms', displayName: 'LinkedIn → Newsletters', createFolder: false, confirmDelete: false, isEnabled: true };

    it('refuses a rule without any condition', async () => {
      await expect(handlers.imap_outlook_create_rule({ ...base, action: 'markRead' })).rejects.toThrow(/without any condition/);
      await expect(handlers.imap_outlook_create_rule({ ...base, action: 'markRead', senderContains: [] })).rejects.toThrow(/without any condition/);
      expect(mockRulesService.createRule).not.toHaveBeenCalled();
    });

    it('refuses action delete without confirmDelete', async () => {
      await expect(handlers.imap_outlook_create_rule({ ...base, action: 'delete', senderContains: ['@spam.example'] })).rejects.toThrow(/confirmDelete: true/);
      expect(mockRulesService.createRule).not.toHaveBeenCalled();
    });

    it('refuses a move without a folder', async () => {
      await expect(handlers.imap_outlook_create_rule({ ...base, action: 'move', senderContains: ['@a'] })).rejects.toThrow(/needs moveToFolder/);
    });

    it('builds the Graph payload: substring conditions, fromAddresses wrapper, folder id, exceptions, next sequence', async () => {
      const result = parse(await handlers.imap_outlook_create_rule({
        ...base,
        senderContains: '@linkedin.com',
        fromAddresses: ['jobs@linkedin.com', 'JOBS@linkedin.com'],
        subjectContains: '["digest","weekly"]', // a client that stringified the array
        exceptSubjectContains: ['receipt'],
        action: 'moveAndMarkRead',
        moveToFolder: 'Newsletters',
        markImportance: 'low',
      }));

      expect(mockRulesService.resolveFolder).toHaveBeenCalledWith(expect.objectContaining({ id: 'acc-ms' }), 'Newsletters', { create: false });
      expect(mockRulesService.createRule).toHaveBeenCalledWith(expect.objectContaining({ id: 'acc-ms' }), {
        displayName: 'LinkedIn → Newsletters',
        sequence: 2,
        isEnabled: true,
        conditions: {
          senderContains: ['@linkedin.com'],
          fromAddresses: [{ emailAddress: { address: 'jobs@linkedin.com' } }],
          subjectContains: ['digest', 'weekly'],
        },
        exceptions: { subjectContains: ['receipt'] },
        actions: { moveToFolder: 'ID-NEWS', markAsRead: true, markImportance: 'low', stopProcessingRules: true },
      });
      expect(result.success).toBe(true);
      expect(result.rule.id).toBe('R-NEW');
      expect(result.rule.actions.moveToFolderPath).toBe('Newsletters');
      expect(result.message).toMatch(/moves to "Newsletters"/);
    });

    it('honours an explicit sequence without listing rules, and creates the folder on request', async () => {
      const result = parse(await handlers.imap_outlook_create_rule({
        ...base, senderContains: ['@uber.com'], action: 'move', moveToFolder: 'Inbox/Uber', createFolder: true, sequence: 7, stopProcessingRules: false,
      }));
      expect(mockRulesService.listRules).not.toHaveBeenCalled();
      expect(mockRulesService.resolveFolder).toHaveBeenCalledWith(expect.anything(), 'Inbox/Uber', { create: true });
      expect(mockRulesService.createRule.mock.calls[0][1]).toMatchObject({ sequence: 7, actions: { moveToFolder: 'ID-CREATED', stopProcessingRules: false } });
      expect(result.createdFolders).toEqual(['Inbox/Uber']);
    });

    it('surfaces an unknown folder as an error rather than creating it', async () => {
      await expect(handlers.imap_outlook_create_rule({ ...base, senderContains: ['@a'], action: 'move', moveToFolder: 'Typo' }))
        .rejects.toThrow(/Folder "Typo" not found/);
      expect(mockRulesService.createRule).not.toHaveBeenCalled();
    });

    it('confirmed delete rules send Graph delete: true', async () => {
      await handlers.imap_outlook_create_rule({ ...base, senderContains: ['@spam.example'], action: 'delete', confirmDelete: true, sequence: 1 });
      expect(mockRulesService.createRule.mock.calls[0][1].actions).toEqual({ delete: true, stopProcessingRules: true });
    });
  });

  describe('imap_outlook_update_rule', () => {
    const base = { accountId: 'acc-ms', ruleId: 'R1', createFolder: false, confirmDelete: false };

    it('merges a changed predicate into the stored conditions instead of replacing them', async () => {
      mockRulesService.getRule.mockResolvedValueOnce({ ...existingRule, conditions: { senderContains: ['@linkedin.com'], subjectContains: ['old'] } });
      const result = parse(await handlers.imap_outlook_update_rule({ ...base, subjectContains: ['new'], exceptSenderContains: ['boss@linkedin.com'], isEnabled: false }));
      expect(mockRulesService.updateRule).toHaveBeenCalledWith(expect.anything(), 'R1', {
        isEnabled: false,
        conditions: { senderContains: ['@linkedin.com'], subjectContains: ['new'] },
        exceptions: { senderContains: ['boss@linkedin.com'] },
      });
      expect(result.changed.sort()).toEqual(['conditions', 'exceptions', 'isEnabled']);
    });

    it('an empty list removes that predicate, but the rule must keep at least one condition', async () => {
      mockRulesService.getRule.mockResolvedValueOnce({ ...existingRule, conditions: { senderContains: ['@linkedin.com'], subjectContains: ['old'] } });
      await handlers.imap_outlook_update_rule({ ...base, subjectContains: [] });
      expect(mockRulesService.updateRule.mock.calls[0][2]).toEqual({ conditions: { senderContains: ['@linkedin.com'] } });

      await expect(handlers.imap_outlook_update_rule({ ...base, senderContains: [] })).rejects.toThrow(/without any condition/);
    });

    it('action rebuilds the actions; moveToFolder alone only swaps the target', async () => {
      await handlers.imap_outlook_update_rule({ ...base, action: 'markRead' });
      expect(mockRulesService.updateRule.mock.calls[0][2]).toEqual({ actions: { markAsRead: true, stopProcessingRules: true } });

      await handlers.imap_outlook_update_rule({ ...base, moveToFolder: 'Newsletters', markImportance: 'high' });
      expect(mockRulesService.updateRule.mock.calls[1][2]).toEqual({ actions: { moveToFolder: 'ID-NEWS', markImportance: 'high', stopProcessingRules: true } });
    });

    it('refuses delete without confirmation and an empty update', async () => {
      await expect(handlers.imap_outlook_update_rule({ ...base, action: 'delete' })).rejects.toThrow(/confirmDelete: true/);
      await expect(handlers.imap_outlook_update_rule({ ...base })).rejects.toThrow(/Nothing to update/);
      expect(mockRulesService.updateRule).not.toHaveBeenCalled();
    });
  });

  describe('imap_outlook_delete_rule', () => {
    it('deletes by id and reports it', async () => {
      const result = parse(await handlers.imap_outlook_delete_rule({ accountId: 'acc-ms', ruleId: 'R1' }));
      expect(mockRulesService.deleteRule).toHaveBeenCalledWith(expect.objectContaining({ id: 'acc-ms' }), 'R1');
      expect(result).toMatchObject({ success: true, ruleId: 'R1' });
    });
  });

  describe('payload helpers', () => {
    it('buildPredicates separates conditions from exceptions and records cleared keys', () => {
      const built = buildPredicates({ senderContains: ['a'], fromAddresses: [], exceptSubjectContains: 'receipt' });
      expect(built.conditions).toEqual({ senderContains: ['a'] });
      expect(built.cleared.conditions).toEqual(['fromAddresses']);
      expect(built.exceptions).toEqual({ subjectContains: ['receipt'] });
      expect(built.conditionsTouched).toBe(true);
      expect(built.exceptionsTouched).toBe(true);
      expect(buildPredicates({}).conditionsTouched).toBe(false);
    });

    it('buildActions maps the four actions', () => {
      expect(buildActions({ action: 'move', moveToFolderId: 'F', stopProcessingRules: true })).toEqual({ moveToFolder: 'F', stopProcessingRules: true });
      expect(buildActions({ action: 'delete', stopProcessingRules: false })).toEqual({ delete: true, stopProcessingRules: false });
      expect(buildActions({ action: 'markRead', markImportance: 'low', stopProcessingRules: true })).toEqual({ markAsRead: true, markImportance: 'low', stopProcessingRules: true });
    });
  });
});
