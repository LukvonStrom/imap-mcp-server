import type { ImapAccount } from '../types/index.js';
import {
  ConsentRequiredError,
  MICROSOFT_GRAPH_RULES_SCOPES,
  MicrosoftOAuthService,
} from './oauth-service.js';

/**
 * Thin Microsoft Graph client for Outlook.com / Microsoft 365 **inbox rules**
 * (`/me/mailFolders/inbox/messageRules`) and the folder lookups a rule's
 * `moveToFolder` needs.
 *
 * Rules are the only server-side automation Outlook offers and Graph is the
 * only API for them — IMAP cannot see or change them. Every call carries a
 * Bearer token minted for the Graph resource (`MICROSOFT_GRAPH_RULES_SCOPES`)
 * by `MicrosoftOAuthService.getValidAccessToken(account, scopes)`; the token
 * is never logged or returned.
 *
 * Outbound host: `https://graph.microsoft.com` only (besides the identity
 * host the OAuth service talks to). Documented in README.md, SECURITY.md and
 * AGENTS.md alongside `login.microsoftonline.com`.
 */

export const MICROSOFT_GRAPH_HOST = 'https://graph.microsoft.com';
const GRAPH_BASE = `${MICROSOFT_GRAPH_HOST}/v1.0`;

/** Page size for folder listings; Outlook mailboxes rarely have more folders per level. */
const FOLDER_PAGE_SIZE = 250;

/**
 * Well-known folder names Graph resolves without an id, keyed by the
 * spellings a user (or an IMAP folder list) is likely to use. Lower-case.
 */
const WELL_KNOWN_FOLDERS: Record<string, string> = {
  inbox: 'inbox',
  junk: 'junkemail',
  'junk email': 'junkemail',
  junkemail: 'junkemail',
  spam: 'junkemail',
  'deleted items': 'deleteditems',
  deleteditems: 'deleteditems',
  deleted: 'deleteditems',
  trash: 'deleteditems',
  archive: 'archive',
  drafts: 'drafts',
  'sent items': 'sentitems',
  sentitems: 'sentitems',
  sent: 'sentitems',
};

/** Graph's `emailAddress` wrapper used in `fromAddresses` / `sentToAddresses`. */
export interface GraphRecipient {
  emailAddress: { address: string; name?: string };
}

/** Subset of Graph `messageRulePredicates` this server reads and writes. */
export interface GraphRulePredicates {
  senderContains?: string[];
  fromAddresses?: GraphRecipient[];
  subjectContains?: string[];
  bodyOrSubjectContains?: string[];
  bodyContains?: string[];
  headerContains?: string[];
  recipientContains?: string[];
  sentToAddresses?: GraphRecipient[];
  importance?: 'low' | 'normal' | 'high';
  hasAttachments?: boolean;
  isAutomaticForward?: boolean;
  isAutomaticReply?: boolean;
  isMeetingRequest?: boolean;
  sentOnlyToMe?: boolean;
  sentToMe?: boolean;
  sentToOrCcMe?: boolean;
  [key: string]: unknown;
}

/** Subset of Graph `messageRuleActions` this server reads and writes. */
export interface GraphRuleActions {
  moveToFolder?: string;
  copyToFolder?: string;
  markAsRead?: boolean;
  delete?: boolean;
  permanentDelete?: boolean;
  markImportance?: 'low' | 'normal' | 'high';
  stopProcessingRules?: boolean;
  assignCategories?: string[];
  forwardTo?: GraphRecipient[];
  forwardAsAttachmentTo?: GraphRecipient[];
  redirectTo?: GraphRecipient[];
  [key: string]: unknown;
}

/** A Graph `messageRule` as returned by the API. */
export interface GraphMessageRule {
  id: string;
  displayName: string;
  sequence: number;
  isEnabled: boolean;
  hasError?: boolean;
  isReadOnly?: boolean;
  conditions?: GraphRulePredicates | null;
  actions?: GraphRuleActions | null;
  exceptions?: GraphRulePredicates | null;
}

/** Body for creating a rule (POST) — everything but the id. */
export type GraphMessageRuleInput = Omit<GraphMessageRule, 'id' | 'hasError' | 'isReadOnly'>;

/** Body for PATCH — any subset. Nested objects replace the stored one wholesale. */
export type GraphMessageRulePatch = Partial<GraphMessageRuleInput>;

/** One mail folder with its display path (`Inbox/Paypal`), for id ↔ path mapping. */
export interface GraphFolder {
  id: string;
  displayName: string;
  /** Slash-joined display names from the root, e.g. `Inbox/Paypal`. */
  path: string;
  parentFolderId?: string;
  childFolderCount: number;
}

interface GraphFolderRaw {
  id: string;
  displayName: string;
  parentFolderId?: string;
  childFolderCount?: number;
}

interface GraphErrorBody {
  error?: { code?: string; message?: string };
}

/**
 * A non-2xx Graph response. `code` is Graph's own error code
 * (`ErrorAccessDenied`, `ErrorItemNotFound`, `MailboxNotEnabledForRESTAPI`, …).
 */
export class GraphApiError extends Error {
  constructor(
    message: string,
    readonly status: number,
    readonly code: string,
    readonly method: string,
    readonly path: string,
  ) {
    super(message);
    this.name = 'GraphApiError';
  }
}

export interface OutlookRulesServiceOptions {
  /** Injectable HTTP client for tests; defaults to the global `fetch`. */
  fetch?: typeof fetch;
  /** Injectable timer for the single 429 retry. */
  sleep?: (ms: number) => Promise<void>;
  /** Longest a `Retry-After` is honoured before giving up (ms). */
  maxRetryAfterMs?: number;
}

/** Split a folder path on `/`, dropping empty segments (`/Inbox//Paypal/` → `Inbox`, `Paypal`). */
export function splitFolderPath(path: string): string[] {
  return path.split('/').map(s => s.trim()).filter(Boolean);
}

/** Map a user-facing folder name to Graph's well-known folder name, if it is one. */
export function wellKnownFolderName(segment: string): string | undefined {
  return WELL_KNOWN_FOLDERS[segment.trim().toLowerCase()];
}

export class OutlookRulesService {
  private readonly fetchImpl: typeof fetch;
  private readonly sleep: (ms: number) => Promise<void>;
  private readonly maxRetryAfterMs: number;

  constructor(private readonly oauthService: MicrosoftOAuthService, options: OutlookRulesServiceOptions = {}) {
    this.fetchImpl = options.fetch ?? ((input, init) => fetch(input, init));
    this.sleep = options.sleep ?? ((ms) => new Promise(resolve => setTimeout(resolve, ms)));
    this.maxRetryAfterMs = options.maxRetryAfterMs ?? 30_000;
  }

  // ---------------------------------------------------------------------------
  // Rules
  // ---------------------------------------------------------------------------

  async listRules(account: ImapAccount): Promise<GraphMessageRule[]> {
    const body = await this.request<{ value?: GraphMessageRule[] }>(account, 'GET', '/me/mailFolders/inbox/messageRules');
    return body?.value ?? [];
  }

  async getRule(account: ImapAccount, ruleId: string): Promise<GraphMessageRule> {
    const rule = await this.request<GraphMessageRule>(account, 'GET', `/me/mailFolders/inbox/messageRules/${encodeURIComponent(ruleId)}`);
    if (!rule) throw new GraphApiError(`Rule ${ruleId} not found`, 404, 'ErrorItemNotFound', 'GET', ruleId);
    return rule;
  }

  async createRule(account: ImapAccount, rule: GraphMessageRuleInput): Promise<GraphMessageRule> {
    const created = await this.request<GraphMessageRule>(account, 'POST', '/me/mailFolders/inbox/messageRules', rule);
    if (!created) throw new GraphApiError('Graph returned no rule body on create', 0, 'EmptyResponse', 'POST', '/me/mailFolders/inbox/messageRules');
    return created;
  }

  async updateRule(account: ImapAccount, ruleId: string, patch: GraphMessageRulePatch): Promise<GraphMessageRule> {
    const path = `/me/mailFolders/inbox/messageRules/${encodeURIComponent(ruleId)}`;
    const updated = await this.request<GraphMessageRule>(account, 'PATCH', path, patch);
    // Graph answers PATCH with the updated rule; fall back to a GET if a
    // deployment ever returns 204.
    return updated ?? this.getRule(account, ruleId);
  }

  async deleteRule(account: ImapAccount, ruleId: string): Promise<void> {
    await this.request<void>(account, 'DELETE', `/me/mailFolders/inbox/messageRules/${encodeURIComponent(ruleId)}`);
  }

  // ---------------------------------------------------------------------------
  // Folders
  // ---------------------------------------------------------------------------

  /**
   * Resolve a display path such as `Newsletters` or `Inbox/Paypal` to a
   * folder id. The first segment may be a well-known name (`inbox`, `junk`,
   * `archive`, …); every segment is matched against display names
   * case-insensitively. Top-level user folders in Outlook are siblings of
   * Inbox, so `Newsletters` is looked up at the root, not under Inbox.
   *
   * With `create: true`, missing trailing segments are created (a missing
   * *first* segment becomes a root folder). Returns the folder and whether
   * anything was created.
   */
  async resolveFolder(
    account: ImapAccount,
    path: string,
    opts: { create?: boolean } = {},
  ): Promise<{ folder: GraphFolder; created: string[] }> {
    const segments = splitFolderPath(path);
    if (segments.length === 0) {
      throw new Error('Folder path is empty. Use a display path such as "Newsletters" or "Inbox/Paypal".');
    }

    const created: string[] = [];
    let current: GraphFolder | undefined;

    for (let i = 0; i < segments.length; i++) {
      const segment = segments[i];
      let next: GraphFolder | undefined;

      if (i === 0) {
        const wellKnown = wellKnownFolderName(segment);
        if (wellKnown) {
          const raw = await this.request<GraphFolderRaw>(account, 'GET', `/me/mailFolders/${wellKnown}`);
          if (raw) next = this.toFolder(raw, undefined);
        }
        if (!next) {
          const roots = await this.listChildren(account, undefined);
          next = roots.find(f => f.displayName.toLowerCase() === segment.toLowerCase());
        }
      } else {
        const children = await this.listChildren(account, current);
        next = children.find(f => f.displayName.toLowerCase() === segment.toLowerCase());
      }

      if (!next) {
        if (!opts.create) {
          const where = current ? `under "${current.path}"` : 'at the top level';
          throw new Error(
            `Folder "${segment}" not found ${where} (looking for "${segments.join('/')}"). ` +
            'Check the exact name with imap_list_folders, or pass createFolder: true to create it.'
          );
        }
        next = await this.createFolder(account, current, segment);
        created.push(next.path);
      }
      current = next;
    }

    return { folder: current!, created };
  }

  /**
   * Every folder in the mailbox with its display path, from one hierarchy
   * walk (root listing plus one child listing per folder that has children).
   * Used to label `moveToFolder` ids when listing rules.
   */
  async listAllFolders(account: ImapAccount): Promise<GraphFolder[]> {
    const out: GraphFolder[] = [];
    const walk = async (parent: GraphFolder | undefined) => {
      const children = await this.listChildren(account, parent);
      for (const child of children) {
        out.push(child);
        if (child.childFolderCount > 0) await walk(child);
      }
    };
    await walk(undefined);
    return out;
  }

  /** Create `name` under `parent` (or at the root when `parent` is undefined). */
  async createFolder(account: ImapAccount, parent: GraphFolder | undefined, name: string): Promise<GraphFolder> {
    const path = parent ? `/me/mailFolders/${encodeURIComponent(parent.id)}/childFolders` : '/me/mailFolders';
    const raw = await this.request<GraphFolderRaw>(account, 'POST', path, { displayName: name });
    if (!raw) throw new GraphApiError(`Graph returned no folder body when creating "${name}"`, 0, 'EmptyResponse', 'POST', path);
    return this.toFolder(raw, parent);
  }

  private async listChildren(account: ImapAccount, parent: GraphFolder | undefined): Promise<GraphFolder[]> {
    const base = parent
      ? `/me/mailFolders/${encodeURIComponent(parent.id)}/childFolders`
      : '/me/mailFolders';
    const out: GraphFolder[] = [];
    let path: string | undefined = `${base}?$top=${FOLDER_PAGE_SIZE}&$select=id,displayName,parentFolderId,childFolderCount`;
    // Follow @odata.nextLink for the rare mailbox with more than a page per level.
    while (path) {
      const page: { value?: GraphFolderRaw[]; '@odata.nextLink'?: string } | undefined =
        await this.request(account, 'GET', path);
      for (const raw of page?.value ?? []) out.push(this.toFolder(raw, parent));
      const next: string | undefined = page?.['@odata.nextLink'];
      path = next && next.startsWith(GRAPH_BASE) ? next.slice(GRAPH_BASE.length) : undefined;
    }
    return out;
  }

  private toFolder(raw: GraphFolderRaw, parent: GraphFolder | undefined): GraphFolder {
    return {
      id: raw.id,
      displayName: raw.displayName,
      path: parent ? `${parent.path}/${raw.displayName}` : raw.displayName,
      parentFolderId: raw.parentFolderId,
      childFolderCount: raw.childFolderCount ?? 0,
    };
  }

  // ---------------------------------------------------------------------------
  // HTTP
  // ---------------------------------------------------------------------------

  /**
   * One Graph request with a Graph-scoped Bearer token. A 401 is retried
   * once with a forced refresh (a cached token may have been revoked); a
   * 429 is retried once after `Retry-After`. Everything else non-2xx becomes
   * a `GraphApiError`; 401/403 that survive the retry become
   * `ConsentRequiredError` so the tools can point at the authorize step.
   */
  private async request<T>(account: ImapAccount, method: string, path: string, body?: unknown): Promise<T | undefined> {
    let token = await this.oauthService.getValidAccessToken(account, MICROSOFT_GRAPH_RULES_SCOPES);
    let retriedAuth = false;
    let retriedRate = false;

    for (;;) {
      const res = await this.fetchImpl(`${GRAPH_BASE}${path}`, {
        method,
        headers: {
          Authorization: `Bearer ${token}`,
          Accept: 'application/json',
          ...(body !== undefined ? { 'Content-Type': 'application/json' } : {}),
        },
        ...(body !== undefined ? { body: JSON.stringify(body) } : {}),
      });

      if (res.ok) {
        if (res.status === 204) return undefined;
        const text = await res.text();
        if (!text) return undefined;
        return JSON.parse(text) as T;
      }

      if (res.status === 401 && !retriedAuth) {
        retriedAuth = true;
        token = await this.oauthService.forceRefresh(account, MICROSOFT_GRAPH_RULES_SCOPES);
        continue;
      }

      if (res.status === 429 && !retriedRate) {
        retriedRate = true;
        const header = res.headers?.get?.('Retry-After');
        const seconds = header ? Number(header) : NaN;
        const waitMs = Number.isFinite(seconds) && seconds >= 0 ? seconds * 1000 : 2_000;
        if (waitMs <= this.maxRetryAfterMs) {
          await this.sleep(waitMs);
          continue;
        }
      }

      let parsed: GraphErrorBody = {};
      try {
        parsed = JSON.parse(await res.text()) as GraphErrorBody;
      } catch {
        // Non-JSON error body — fall through with the status only.
      }
      const code = parsed.error?.code || `HTTP_${res.status}`;
      const message = parsed.error?.message || res.statusText || `HTTP ${res.status}`;

      if (res.status === 401 || res.status === 403) {
        throw new ConsentRequiredError(
          `Microsoft Graph refused ${method} ${path} for account "${account.name}" (${res.status} ${code}: ${message}). ` +
          'The token lacks the Graph permissions for inbox rules — the user has to consent to MailboxSettings.ReadWrite and Mail.ReadBasic once, ' +
          'and the Entra app registration must list those delegated Microsoft Graph permissions.',
          MICROSOFT_GRAPH_RULES_SCOPES,
          account.id,
        );
      }
      throw new GraphApiError(`Microsoft Graph ${method} ${path} failed: ${res.status} ${code}: ${message}`, res.status, code, method, path);
    }
  }
}
