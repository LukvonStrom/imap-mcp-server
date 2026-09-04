import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { DESTRUCTIVE, MUTATING, NETWORK_AUTH, READ_ONLY } from './annotations.js';
import { AccountManager } from '../services/account-manager.js';
import {
  ConsentRequiredError,
  DEFAULT_POLL_MAX_WAIT_MS,
  MICROSOFT_GRAPH_RULES_SCOPES,
  MicrosoftOAuthService,
  allScopesGranted,
} from '../services/oauth-service.js';
import {
  GraphApiError,
  OutlookRulesService,
  type GraphFolder,
  type GraphMessageRule,
  type GraphMessageRuleInput,
  type GraphMessageRulePatch,
  type GraphRuleActions,
  type GraphRulePredicates,
} from '../services/outlook-rules-service.js';
import type { ImapAccount } from '../types/index.js';
import { parseSerializedArray } from '../utils/array-input.js';

/**
 * Outlook.com / Microsoft 365 inbox rules through Microsoft Graph.
 *
 * Rules are the one piece of server-side automation Outlook offers and Graph
 * is the only API for them, so these tools are the difference between "the
 * assistant tells you which rule to click together in Outlook's settings"
 * and "the assistant creates it". They need a second consent from the user
 * (the Graph scopes, separate from the IMAP/SMTP ones) — see
 * `imap_outlook_authorize_rules` — and they only work on accounts that
 * authenticate with `authType: "oauth2"` / `provider: "microsoft"`.
 *
 * Outbound host: `https://graph.microsoft.com` (rules, folder lookup) plus
 * `login.microsoftonline.com` for the token — nothing else. Tokens are never
 * returned; rule definitions are returned in full.
 */

const AUTHORIZE_TOOL = 'imap_outlook_authorize_rules';

/**
 * What `imap_outlook_authorize_rules` attaches to its device-code flow and
 * `imap_complete_oauth_login` gets back. No secrets — the flow id is the
 * only handle and the device code stays inside the OAuth service.
 */
export interface PendingGraphConsent {
  kind: 'graph-consent';
  accountId: string;
  scopes: string[];
}

const accountSelector = {
  accountId: z.string().optional().describe('Account ID (from imap_list_accounts). Optional if accountName is given or only one account is configured. Must be an Outlook.com / Microsoft 365 account added with imap_add_oauth_account.'),
  accountName: z.string().optional().describe('Account name instead of accountId. Optional if accountId is given or only one account is configured.'),
};

/** One-or-many string input; MCP clients sometimes stringify arrays. */
const stringList = z.union([z.string(), z.array(z.string())]);
type StringList = z.infer<typeof stringList>;

const RULE_ACTIONS = ['move', 'markRead', 'delete', 'moveAndMarkRead'] as const;
type RuleAction = typeof RULE_ACTIONS[number];

const IMPORTANCE = ['low', 'normal', 'high'] as const;

/** Condition / exception fields shared by create and update. */
const conditionFields = {
  senderContains: stringList.optional().describe('Match when the sender\'s address OR display name contains any of these substrings (case-insensitive). The natural way to match a whole domain: pass "@linkedin.com" to catch everything from linkedin.com. Multiple values are OR-ed.'),
  fromAddresses: stringList.optional().describe('Match when the sender is exactly one of these email addresses (e.g. "noreply@uber.com"). Use senderContains for domains or partial matches.'),
  subjectContains: stringList.optional().describe('Match when the subject contains any of these substrings (case-insensitive).'),
  bodyOrSubjectContains: stringList.optional().describe('Match when the subject OR body contains any of these substrings.'),
  headerContains: stringList.optional().describe('Match when any message header contains one of these substrings — e.g. a List-Id value such as "list.example.com" to catch a mailing list regardless of sender.'),
  exceptSubjectContains: stringList.optional().describe('Do NOT apply the rule when the subject contains any of these substrings (an exception). Example: file Uber mail but keep receipts in the inbox with exceptSubjectContains: ["receipt"].'),
  exceptSenderContains: stringList.optional().describe('Do NOT apply the rule when the sender address/name contains any of these substrings.'),
  exceptFromAddresses: stringList.optional().describe('Do NOT apply the rule when the sender is exactly one of these addresses.'),
};

const actionFields = {
  moveToFolder: z.string().optional().describe('Destination folder as a display path, e.g. "Newsletters", "Inbox/Paypal", or a well-known name ("Archive", "Junk Email", "Deleted Items"). Required for action "move" / "moveAndMarkRead". Top-level Outlook folders are siblings of Inbox, so "Newsletters" is a root folder and "Inbox/Newsletters" is a subfolder of the inbox — check imap_list_folders for the mailbox\'s layout. Matched case-insensitively.'),
  createFolder: z.boolean().default(false).describe('Create moveToFolder (and any missing parent) when it does not exist yet. Default false: an unknown folder is an error so a typo cannot silently create a new folder.'),
  markImportance: z.enum(IMPORTANCE).optional().describe('Additionally mark matching messages with this importance.'),
  stopProcessingRules: z.boolean().optional().describe('Stop evaluating later rules once this one matched (Outlook\'s "stop processing more rules"). Default true on create — the usual choice for filing rules.'),
};

type ConditionInput = { [K in keyof typeof conditionFields]?: StringList };

/** Normalise a one-or-many input to a de-duplicated string array (empty when absent). */
function toList(value: StringList | undefined, field: string): string[] {
  const recovered = parseSerializedArray(value, field);
  const items = Array.isArray(recovered) ? recovered : recovered === undefined ? [] : [recovered];
  const out: string[] = [];
  for (const item of items) {
    const text = String(item).trim();
    if (text && !out.some(x => x.toLowerCase() === text.toLowerCase())) out.push(text);
  }
  return out;
}

/** Graph's recipient wrapper for `fromAddresses`. */
function recipients(addresses: string[]) {
  return addresses.map(address => ({ emailAddress: { address } }));
}

/**
 * Translate the flat tool inputs into Graph `messageRulePredicates` for the
 * conditions and the exceptions. Returns which of the two were touched so an
 * update can leave the untouched one alone. Exported for tests.
 */
export function buildPredicates(input: ConditionInput): {
  conditions: GraphRulePredicates; conditionsTouched: boolean;
  exceptions: GraphRulePredicates; exceptionsTouched: boolean;
  /** Fields explicitly passed as an empty list — meaning "remove this predicate" on update. */
  cleared: { conditions: (keyof GraphRulePredicates)[]; exceptions: (keyof GraphRulePredicates)[] };
} {
  const conditions: GraphRulePredicates = {};
  const exceptions: GraphRulePredicates = {};
  const cleared = { conditions: [] as (keyof GraphRulePredicates)[], exceptions: [] as (keyof GraphRulePredicates)[] };
  let conditionsTouched = false;
  let exceptionsTouched = false;

  const put = (
    target: GraphRulePredicates,
    key: 'senderContains' | 'subjectContains' | 'bodyOrSubjectContains' | 'headerContains' | 'fromAddresses',
    raw: StringList | undefined,
    field: string,
    bucket: 'conditions' | 'exceptions',
  ) => {
    if (raw === undefined) return;
    if (bucket === 'conditions') conditionsTouched = true; else exceptionsTouched = true;
    const list = toList(raw, field);
    if (list.length === 0) {
      cleared[bucket].push(key);
      return;
    }
    if (key === 'fromAddresses') target.fromAddresses = recipients(list);
    else target[key] = list;
  };

  put(conditions, 'senderContains', input.senderContains, 'senderContains', 'conditions');
  put(conditions, 'fromAddresses', input.fromAddresses, 'fromAddresses', 'conditions');
  put(conditions, 'subjectContains', input.subjectContains, 'subjectContains', 'conditions');
  put(conditions, 'bodyOrSubjectContains', input.bodyOrSubjectContains, 'bodyOrSubjectContains', 'conditions');
  put(conditions, 'headerContains', input.headerContains, 'headerContains', 'conditions');
  put(exceptions, 'subjectContains', input.exceptSubjectContains, 'exceptSubjectContains', 'exceptions');
  put(exceptions, 'senderContains', input.exceptSenderContains, 'exceptSenderContains', 'exceptions');
  put(exceptions, 'fromAddresses', input.exceptFromAddresses, 'exceptFromAddresses', 'exceptions');

  return { conditions, conditionsTouched, exceptions, exceptionsTouched, cleared };
}

/** Graph `messageRuleActions` for one of the four tool-level actions. Exported for tests. */
export function buildActions(opts: {
  action: RuleAction;
  moveToFolderId?: string;
  markImportance?: typeof IMPORTANCE[number];
  stopProcessingRules: boolean;
}): GraphRuleActions {
  const actions: GraphRuleActions = {};
  switch (opts.action) {
    case 'move':
      actions.moveToFolder = opts.moveToFolderId;
      break;
    case 'moveAndMarkRead':
      actions.moveToFolder = opts.moveToFolderId;
      actions.markAsRead = true;
      break;
    case 'markRead':
      actions.markAsRead = true;
      break;
    case 'delete':
      actions.delete = true;
      break;
  }
  if (opts.markImportance) actions.markImportance = opts.markImportance;
  actions.stopProcessingRules = opts.stopProcessingRules;
  return actions;
}

function hasAnyPredicate(p: GraphRulePredicates | null | undefined): boolean {
  if (!p) return false;
  return Object.values(p).some(v => v !== undefined && v !== null && (!Array.isArray(v) || v.length > 0));
}

/** Drop `null` members Graph returns for absent objects; shape one rule for output. */
function presentRule(rule: GraphMessageRule, folderPaths?: Map<string, string>) {
  const actions: Record<string, unknown> | undefined = rule.actions ? { ...rule.actions } : undefined;
  if (actions && folderPaths) {
    for (const key of ['moveToFolder', 'copyToFolder'] as const) {
      const id = actions[key];
      if (typeof id === 'string') {
        const path = folderPaths.get(id);
        if (path) actions[`${key}Path`] = path;
      }
    }
  }
  return {
    id: rule.id,
    displayName: rule.displayName,
    sequence: rule.sequence,
    isEnabled: rule.isEnabled,
    ...(rule.hasError ? { hasError: true } : {}),
    ...(rule.isReadOnly ? { isReadOnly: true } : {}),
    conditions: rule.conditions ?? {},
    actions: actions ?? {},
    exceptions: rule.exceptions ?? {},
  };
}

function json(payload: unknown) {
  return { content: [{ type: 'text' as const, text: JSON.stringify(payload, null, 2) }] };
}

/** True when the stored grant covers the Graph rules scopes (legacy accounts: `scopes` is the whole grant). */
export function hasGraphConsent(account: ImapAccount): boolean {
  const granted = account.oauth?.grantedScopes ?? account.oauth?.scopes;
  return allScopesGranted(granted, MICROSOFT_GRAPH_RULES_SCOPES);
}

function consentRequired(account: ImapAccount, detail?: string) {
  return json({
    success: false,
    error: 'graph-consent-required',
    accountId: account.id,
    nextStep: AUTHORIZE_TOOL,
    requiredScopes: MICROSOFT_GRAPH_RULES_SCOPES.filter(s => s !== 'offline_access'),
    message:
      `Account "${account.name}" has not granted the Microsoft Graph permissions needed for inbox rules. ` +
      `Call ${AUTHORIZE_TOOL} with this accountId, have the user sign in with the code it returns, then finish with imap_complete_oauth_login. ` +
      'The Entra app registration must list the delegated Microsoft Graph permissions MailboxSettings.ReadWrite and Mail.ReadBasic (see README "Outlook.com inbox rules (Graph)").' +
      (detail ? ` Details: ${detail}` : ''),
  });
}

/** The one IMAP capability the rules tools borrow: creating a folder without a Graph write scope. */
export type FolderCreator = { createFolder(accountId: string, folderPath: string): Promise<unknown> };

export function outlookRulesTools(
  server: McpServer,
  accountManager: AccountManager,
  oauthService: MicrosoftOAuthService,
  rulesService: OutlookRulesService,
  imapService?: FolderCreator,
): void {
  /** Resolve the selector to an account and insist on a Microsoft OAuth one. */
  function loadAccount(rawAccountId: string | undefined, accountName: string | undefined): ImapAccount {
    const accountId = accountManager.resolveAccountId(rawAccountId, accountName);
    const account = accountManager.getAccount(accountId);
    if (!account) throw new Error(`Account ${accountId} not found. Use imap_list_accounts to see available accounts.`);
    if (account.authType !== 'oauth2' || account.oauth?.provider !== 'microsoft') {
      throw new Error(
        `Account "${account.name}" is not a Microsoft OAuth 2.0 account (authType "${account.authType ?? 'password'}"). ` +
        'Inbox rules are only available for Outlook.com / Microsoft 365 mailboxes added with imap_add_oauth_account; ' +
        'other providers have no rules API here.'
      );
    }
    return account;
  }

  /**
   * Run a Graph operation and turn "no consent" / Graph failures into
   * structured results the model can act on, instead of opaque throws.
   */
  async function withGraph(account: ImapAccount, run: () => Promise<ReturnType<typeof json>>) {
    if (!hasGraphConsent(account)) return consentRequired(account);
    try {
      return await run();
    } catch (err) {
      if (err instanceof ConsentRequiredError) return consentRequired(account, err.message);
      if (err instanceof GraphApiError) {
        return json({
          success: false,
          error: 'graph-error',
          status: err.status,
          code: err.code,
          message: err.message,
          ...(err.code === 'MailboxNotEnabledForRESTAPI'
            ? { hint: 'This mailbox is not reachable through Microsoft Graph (typically an on-premises or hybrid Exchange mailbox). Rules must be managed in Outlook directly.' }
            : {}),
        });
      }
      throw err;
    }
  }

  /** Folder-id → display-path map for labelling move targets; undefined (with a reason) when the walk fails. */
  async function folderPathMap(account: ImapAccount): Promise<{ map?: Map<string, string>; reason?: string }> {
    try {
      const folders = await rulesService.listAllFolders(account);
      return { map: new Map(folders.map(f => [f.id, f.path])) };
    } catch (err) {
      return { reason: err instanceof Error ? err.message : String(err) };
    }
  }

  async function resolveMoveTarget(account: ImapAccount, path: string, create: boolean): Promise<{ folder: GraphFolder; created: string[] }> {
    if (!create || !imapService) {
      return rulesService.resolveFolder(account, path, { create });
    }
    // Creating a folder through Graph needs Mail.ReadWrite, which the rules
    // consent deliberately does not ask for (Mail.ReadBasic only). The IMAP
    // session is already authorised for the same mailbox, so create there and
    // let Graph merely look the new folder up.
    try {
      return await rulesService.resolveFolder(account, path, { create: false });
    } catch (err) {
      if (err instanceof ConsentRequiredError) throw err;
      await imapService.createFolder(account.id, path);
      const resolved = await rulesService.resolveFolder(account, path, { create: false });
      return { folder: resolved.folder, created: [path] };
    }
  }

  // ---------------------------------------------------------------------------
  // Consent
  // ---------------------------------------------------------------------------

  server.registerTool(AUTHORIZE_TOOL, {
    title: 'Authorize Outlook rules access',
    annotations: NETWORK_AUTH,
    description: `Grant this server access to an Outlook.com / Microsoft 365 account's inbox rules (Microsoft Graph). Needed once per account before imap_outlook_list_rules / imap_outlook_create_rule / imap_outlook_update_rule / imap_outlook_delete_rule work — the mailbox sign-in from imap_add_oauth_account only covers IMAP/SMTP, and Microsoft requires a separate consent for Graph. Step 1 of 2: returns a verificationUri and a short userCode — show BOTH to the user and ask them to open the URL, enter the code, sign in to the same mailbox, and accept the "Read and write your mailbox settings" / "Read basic mail" permissions. Then call imap_complete_oauth_login with the returned flowId (step 2); it stores the widened consent on the existing account (no new account is created). Requires the Entra app registration to list the delegated Microsoft Graph permissions MailboxSettings.ReadWrite and Mail.ReadBasic. Call this when a rules tool returns error "graph-consent-required".`,
    inputSchema: { ...accountSelector },
  }, async ({ accountId: rawAccountId, accountName }) => {
    const account = loadAccount(rawAccountId, accountName);
    const oauth = account.oauth!;

    const pending: PendingGraphConsent = {
      kind: 'graph-consent',
      accountId: account.id,
      scopes: [...MICROSOFT_GRAPH_RULES_SCOPES],
    };
    const start = await oauthService.startDeviceCode<PendingGraphConsent>({
      clientId: oauth.clientId,
      tenant: oauth.tenant,
      scopes: MICROSOFT_GRAPH_RULES_SCOPES,
      context: pending,
    });

    return json({
      status: 'awaiting_user',
      accountId: account.id,
      ...(hasGraphConsent(account) ? { alreadyGranted: true } : {}),
      flowId: start.flowId,
      userCode: start.userCode,
      verificationUri: start.verificationUri,
      expiresAt: new Date(start.expiresAt).toISOString(),
      message: start.message,
      requestedScopes: MICROSOFT_GRAPH_RULES_SCOPES.filter(s => s !== 'offline_access'),
      instructions:
        `Show the user this URL and code: open ${start.verificationUri} in a browser, enter the code ` +
        `${start.userCode}, sign in to ${account.email ?? account.user}, and accept the mailbox-settings permission. ` +
        `The code expires at ${new Date(start.expiresAt).toISOString()}. ` +
        `Then call imap_complete_oauth_login with flowId "${start.flowId}". It waits up to ` +
        `${Math.round(DEFAULT_POLL_MAX_WAIT_MS / 1000)} seconds; if it returns status "pending", call it again with the same flowId. ` +
        'On "complete" the existing account is updated in place and the rules tools become available.',
    });
  });

  // ---------------------------------------------------------------------------
  // List
  // ---------------------------------------------------------------------------

  server.registerTool('imap_outlook_list_rules', {
    title: 'List Outlook inbox rules',
    // Read-only, but it reads from graph.microsoft.com rather than the user's
    // own IMAP server, so it is open-world.
    annotations: { ...READ_ONLY, openWorldHint: true },
    description: 'List the server-side inbox rules of an Outlook.com / Microsoft 365 account via Microsoft Graph: id, displayName, sequence (evaluation order), isEnabled, conditions, actions, exceptions. Move/copy targets are folder ids; the tool adds moveToFolderPath / copyToFolderPath with the display path when the folder hierarchy can be read. Use it before creating a rule (to avoid duplicates and pick a sequence) and to find the ruleId for imap_outlook_update_rule / imap_outlook_delete_rule. Requires the one-time Graph consent from imap_outlook_authorize_rules; returns error "graph-consent-required" otherwise.',
    inputSchema: { ...accountSelector },
  }, async ({ accountId: rawAccountId, accountName }) => {
    const account = loadAccount(rawAccountId, accountName);
    return withGraph(account, async () => {
      const rules = await rulesService.listRules(account);
      const needsFolders = rules.some(r => r.actions?.moveToFolder || r.actions?.copyToFolder);
      const { map, reason } = needsFolders ? await folderPathMap(account) : {};
      const sorted = [...rules].sort((a, b) => (a.sequence ?? 0) - (b.sequence ?? 0));
      return json({
        success: true,
        accountId: account.id,
        count: sorted.length,
        rules: sorted.map(r => presentRule(r, map)),
        ...(needsFolders && !map ? { folderPathsUnavailable: reason } : {}),
      });
    });
  });

  // ---------------------------------------------------------------------------
  // Create
  // ---------------------------------------------------------------------------

  server.registerTool('imap_outlook_create_rule', {
    title: 'Create Outlook inbox rule',
    // Creating a rule is reversible (delete it) but each call adds another,
    // and it talks to graph.microsoft.com.
    annotations: { ...MUTATING, openWorldHint: true, idempotentHint: false },
    description: 'Create a server-side inbox rule on an Outlook.com / Microsoft 365 account via Microsoft Graph — the rule then runs on Microsoft\'s servers for every incoming message, with no client involved. Typical uses: file newsletters into a folder (senderContains "@linkedin.com" → move to "Newsletters"), file a service\'s mail except receipts (senderContains "@uber.com", exceptSubjectContains "receipt"), mark list mail read, or delete known junk. At least ONE condition is required (a rule without conditions would apply to all mail and is refused). Every condition value is a case-insensitive substring; multiple values in one field are OR-ed, different fields are AND-ed. "delete" moves matches to Deleted Items and must be confirmed with confirmDelete: true. Run imap_outlook_list_rules first to avoid duplicates. Requires the one-time Graph consent from imap_outlook_authorize_rules. Returns the created rule.',
    inputSchema: {
      ...accountSelector,
      displayName: z.string().min(1).describe('Name of the rule as shown in Outlook\'s settings, e.g. "LinkedIn → Newsletters".'),
      ...conditionFields,
      action: z.enum(RULE_ACTIONS).describe('What to do with matching messages: "move" (to moveToFolder), "markRead", "moveAndMarkRead", or "delete" (to Deleted Items; requires confirmDelete: true).'),
      ...actionFields,
      confirmDelete: z.boolean().default(false).describe('Must be true for action "delete" — states that the user explicitly wants matching mail deleted rather than filed.'),
      isEnabled: z.boolean().default(true).describe('Whether the rule is active immediately (default true).'),
      sequence: z.coerce.number().int().min(1).optional().describe('Evaluation order among the account\'s rules (1 runs first). Defaults to after the existing rules.'),
    },
  }, async (input) => {
    const account = loadAccount(input.accountId, input.accountName);

    const { conditions, exceptions } = buildPredicates(input);
    if (!hasAnyPredicate(conditions)) {
      throw new Error(
        'Refusing to create a rule without any condition — it would apply to every incoming message. ' +
        'Pass at least one of senderContains, fromAddresses, subjectContains, bodyOrSubjectContains, or headerContains.'
      );
    }
    if (input.action === 'delete' && !input.confirmDelete) {
      throw new Error(
        'action "delete" moves every matching message to Deleted Items automatically. Confirm with the user that this is wanted, then call again with confirmDelete: true — or use action "move" to file the mail into a folder instead.'
      );
    }
    const isMove = input.action === 'move' || input.action === 'moveAndMarkRead';
    if (isMove && !input.moveToFolder) {
      throw new Error(`action "${input.action}" needs moveToFolder (a display path such as "Newsletters" or "Inbox/Paypal").`);
    }

    return withGraph(account, async () => {
      let target: { folder: GraphFolder; created: string[] } | undefined;
      if (isMove) target = await resolveMoveTarget(account, input.moveToFolder!, input.createFolder);

      let sequence = input.sequence;
      if (sequence === undefined) {
        const existing = await rulesService.listRules(account);
        sequence = existing.reduce((max, r) => Math.max(max, r.sequence ?? 0), 0) + 1;
      }

      const rule: GraphMessageRuleInput = {
        displayName: input.displayName,
        sequence,
        isEnabled: input.isEnabled,
        conditions,
        ...(hasAnyPredicate(exceptions) ? { exceptions } : {}),
        actions: buildActions({
          action: input.action,
          moveToFolderId: target?.folder.id,
          markImportance: input.markImportance,
          stopProcessingRules: input.stopProcessingRules ?? true,
        }),
      };

      const created = await rulesService.createRule(account, rule);
      const paths = target ? new Map([[target.folder.id, target.folder.path]]) : undefined;
      return json({
        success: true,
        accountId: account.id,
        rule: presentRule(created, paths),
        ...(target?.created.length ? { createdFolders: target.created } : {}),
        message: `Rule "${created.displayName}" created` + (target ? ` (moves to "${target.folder.path}")` : ''),
      });
    });
  });

  // ---------------------------------------------------------------------------
  // Update
  // ---------------------------------------------------------------------------

  server.registerTool('imap_outlook_update_rule', {
    title: 'Update Outlook inbox rule',
    annotations: { ...MUTATING, openWorldHint: true, idempotentHint: false },
    description: 'Change an existing Outlook.com / Microsoft 365 inbox rule via Microsoft Graph. Get the ruleId from imap_outlook_list_rules. Only the fields you pass change: a condition/exception field replaces that one predicate (pass an empty list to remove it; other predicates stay); `action` rebuilds the actions; moveToFolder / markImportance / stopProcessingRules on their own adjust just that action; displayName, isEnabled (enable/disable without deleting), and sequence are set directly. The resulting rule must keep at least one condition. Switching to action "delete" requires confirmDelete: true. Returns the updated rule.',
    inputSchema: {
      ...accountSelector,
      ruleId: z.string().min(1).describe('The rule id from imap_outlook_list_rules.'),
      displayName: z.string().min(1).optional().describe('New name for the rule.'),
      ...conditionFields,
      action: z.enum(RULE_ACTIONS).optional().describe('Replace the rule\'s actions with "move", "markRead", "moveAndMarkRead", or "delete" (requires confirmDelete: true).'),
      ...actionFields,
      confirmDelete: z.boolean().default(false).describe('Must be true when switching the rule to action "delete".'),
      isEnabled: z.boolean().optional().describe('Enable (true) or disable (false) the rule.'),
      sequence: z.coerce.number().int().min(1).optional().describe('New evaluation order (1 runs first).'),
    },
  }, async (input) => {
    const account = loadAccount(input.accountId, input.accountName);

    if (input.action === 'delete' && !input.confirmDelete) {
      throw new Error('Switching a rule to action "delete" needs confirmDelete: true — confirm with the user first.');
    }
    const isMoveAction = input.action === 'move' || input.action === 'moveAndMarkRead';
    if (isMoveAction && !input.moveToFolder) {
      throw new Error(`action "${input.action}" needs moveToFolder (a display path such as "Newsletters" or "Inbox/Paypal").`);
    }

    return withGraph(account, async () => {
      const existing = await rulesService.getRule(account, input.ruleId);
      const patch: GraphMessageRulePatch = {};

      if (input.displayName !== undefined) patch.displayName = input.displayName;
      if (input.isEnabled !== undefined) patch.isEnabled = input.isEnabled;
      if (input.sequence !== undefined) patch.sequence = input.sequence;

      // Graph replaces nested objects wholesale on PATCH, so merge with what
      // is stored rather than sending only the changed predicate.
      const built = buildPredicates(input);
      if (built.conditionsTouched) {
        const merged: GraphRulePredicates = { ...(existing.conditions ?? {}), ...built.conditions };
        for (const key of built.cleared.conditions) delete merged[key];
        if (!hasAnyPredicate(merged)) {
          throw new Error('The update would leave the rule without any condition, so it would apply to every message. Keep at least one condition, or delete the rule instead.');
        }
        patch.conditions = merged;
      }
      if (built.exceptionsTouched) {
        const merged: GraphRulePredicates = { ...(existing.exceptions ?? {}), ...built.exceptions };
        for (const key of built.cleared.exceptions) delete merged[key];
        patch.exceptions = merged;
      }

      let target: { folder: GraphFolder; created: string[] } | undefined;
      if (input.moveToFolder !== undefined) {
        target = await resolveMoveTarget(account, input.moveToFolder, input.createFolder);
      }

      if (input.action !== undefined) {
        patch.actions = buildActions({
          action: input.action,
          moveToFolderId: target?.folder.id,
          markImportance: input.markImportance ?? existing.actions?.markImportance,
          stopProcessingRules: input.stopProcessingRules ?? existing.actions?.stopProcessingRules ?? true,
        });
      } else if (target || input.markImportance !== undefined || input.stopProcessingRules !== undefined) {
        const actions: GraphRuleActions = { ...(existing.actions ?? {}) };
        if (target) actions.moveToFolder = target.folder.id;
        if (input.markImportance !== undefined) actions.markImportance = input.markImportance;
        if (input.stopProcessingRules !== undefined) actions.stopProcessingRules = input.stopProcessingRules;
        patch.actions = actions;
      }

      if (Object.keys(patch).length === 0) {
        throw new Error('Nothing to update — pass at least one field besides ruleId.');
      }

      const updated = await rulesService.updateRule(account, input.ruleId, patch);
      const paths = target ? new Map([[target.folder.id, target.folder.path]]) : undefined;
      return json({
        success: true,
        accountId: account.id,
        rule: presentRule(updated, paths),
        changed: Object.keys(patch),
        ...(target?.created.length ? { createdFolders: target.created } : {}),
        message: `Rule "${updated.displayName}" updated`,
      });
    });
  });

  // ---------------------------------------------------------------------------
  // Delete
  // ---------------------------------------------------------------------------

  server.registerTool('imap_outlook_delete_rule', {
    title: 'Delete Outlook inbox rule',
    annotations: { ...DESTRUCTIVE, openWorldHint: true },
    description: 'Permanently delete an Outlook.com / Microsoft 365 inbox rule via Microsoft Graph. Get the ruleId from imap_outlook_list_rules and confirm with the user first — the rule definition cannot be recovered (mail it already filed is untouched). To pause a rule instead, use imap_outlook_update_rule with isEnabled: false.',
    inputSchema: {
      ...accountSelector,
      ruleId: z.string().min(1).describe('The rule id from imap_outlook_list_rules.'),
    },
  }, async ({ accountId: rawAccountId, accountName, ruleId }) => {
    const account = loadAccount(rawAccountId, accountName);
    return withGraph(account, async () => {
      await rulesService.deleteRule(account, ruleId);
      return json({ success: true, accountId: account.id, ruleId, message: `Rule ${ruleId} deleted` });
    });
  });
}
