import { describe, it, expect, vi, beforeEach } from 'vitest';
import { outlookRulesTools } from '../src/tools/outlook-rules-tools.js';
import { ConsentRequiredError, MICROSOFT_GRAPH_RULES_SCOPES } from '../src/services/oauth-service.js';
import type { ImapAccount } from '../src/types/index.js';

// Creating a mail folder through Graph needs Mail.ReadWrite, which the rules
// consent does not request. When createFolder is set the tool must create the
// folder over the already-authorised IMAP session and then resolve it via
// Graph, never asking Graph to create it.
const handlers: Record<string, Function> = {};
const mockServer = { registerTool: vi.fn((name: string, _s: any, h: Function) => { handlers[name] = h; }) };

const account = (): ImapAccount => ({
  id: 'acc', name: 'Outlook', host: 'outlook.office365.com', port: 993, user: 'me@outlook.com', password: '', tls: true,
  authType: 'oauth2',
  oauth: { provider: 'microsoft', clientId: 'cid', tenant: 'consumers', refreshToken: 'r', scopes: [], grantedScopes: [...MICROSOFT_GRAPH_RULES_SCOPES] },
} as any);

const mockAccountManager = { resolveAccountId: vi.fn(() => 'acc'), getAccount: vi.fn(() => account()) };
const mockOAuthService = {};
let known = new Set<string>(['newsletters']);
const mockRulesService = {
  listRules: vi.fn(async () => []),
  createRule: vi.fn(async (_a: ImapAccount, body: any) => ({ id: 'R-NEW', ...body })),
  resolveFolder: vi.fn(async (_a: ImapAccount, path: string, opts: any) => {
    if (known.has(path.toLowerCase())) return { folder: { id: `ID-${path}`, displayName: path, path, childFolderCount: 0 }, created: [] };
    if (opts?.create) throw new Error('Graph create should not be used when an IMAP folder creator is available');
    throw new Error(`Folder "${path}" not found`);
  }),
  listAllFolders: vi.fn(async () => []),
};
const mockImap = { createFolder: vi.fn(async (_id: string, path: string) => { known.add(path.toLowerCase()); return { path, created: true, alreadyExisted: false }; }) };

const parse = (r: any) => JSON.parse(r.content[0].text);

describe('imap_outlook_create_rule createFolder fallback over IMAP', () => {
  beforeEach(() => {
    vi.clearAllMocks();
    known = new Set(['newsletters']);
    outlookRulesTools(mockServer as any, mockAccountManager as any, mockOAuthService as any, mockRulesService as any, mockImap as any);
  });

  it('creates a missing target folder over IMAP, then resolves it through Graph', async () => {
    const res = parse(await handlers['imap_outlook_create_rule']({
      accountName: 'Outlook', displayName: 'Reading', senderContains: ['@example.com'], action: 'move', moveToFolder: 'Reading', createFolder: true,
    }));
    expect(res.success).not.toBe(false);
    expect(mockImap.createFolder).toHaveBeenCalledWith('acc', 'Reading');
    expect(mockRulesService.resolveFolder).not.toHaveBeenCalledWith(expect.anything(), 'Reading', expect.objectContaining({ create: true }));
    expect(mockRulesService.createRule.mock.calls[0][1].actions.moveToFolder).toBe('ID-Reading');
  });

  it('does not touch IMAP when the folder already exists', async () => {
    await handlers['imap_outlook_create_rule']({
      accountName: 'Outlook', displayName: 'N', senderContains: ['@example.com'], action: 'move', moveToFolder: 'Newsletters', createFolder: true,
    });
    expect(mockImap.createFolder).not.toHaveBeenCalled();
  });

  it('still surfaces a Graph consent problem instead of masking it with a folder create', async () => {
    mockRulesService.resolveFolder.mockRejectedValueOnce(new ConsentRequiredError('consent', [...MICROSOFT_GRAPH_RULES_SCOPES]));
    const res = parse(await handlers['imap_outlook_create_rule']({
      accountName: 'Outlook', displayName: 'N', senderContains: ['@example.com'], action: 'move', moveToFolder: 'Reading', createFolder: true,
    }));
    expect(res.error).toBe('graph-consent-required');
    expect(mockImap.createFolder).not.toHaveBeenCalled();
  });
});
