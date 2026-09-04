import crypto from 'crypto';
import type { ImapAccount, OAuthConfig } from '../types/index.js';
import type { AccountManager } from './account-manager.js';

/**
 * OAuth 2.0 device-code flow and token refresh against Microsoft Entra ID, for
 * XOAUTH2 logins to Outlook.com / Hotmail / Live and Microsoft 365 mailboxes.
 *
 * Microsoft no longer accepts passwords or app passwords for IMAP/SMTP, so an
 * access token is the only way in. The device-code flow fits an MCP server
 * well: no redirect URI, no local HTTP listener — the user opens a URL, types a
 * short code, and the server polls for the result.
 *
 * Outbound traffic: `https://login.microsoftonline.com` only (the device-code
 * and token endpoints). Tokens are never logged; the device_code never leaves
 * this process.
 *
 * Access tokens are resource-bound: Microsoft's v2 endpoint mints a token for
 * exactly one resource per request, so a token that works against IMAP/SMTP
 * (`https://outlook.office.com/...`) is useless against Microsoft Graph and
 * vice versa. The *refresh* token of a public client is multi-resource,
 * though — once the user has consented to a second resource's scopes, the
 * same refresh token can be redeemed for either. `getValidAccessToken` takes
 * the scope set and caches per (account, scope set): mail tokens persist on
 * the account as before, every other set lives in memory only.
 */

export const MICROSOFT_LOGIN_HOST = 'https://login.microsoftonline.com';

/** Scopes that let an access token drive IMAP and SMTP plus obtain a refresh token. */
export const MICROSOFT_OAUTH_SCOPES = [
  'https://outlook.office.com/IMAP.AccessAsUser.All',
  'https://outlook.office.com/SMTP.Send',
  'offline_access',
];

/**
 * Microsoft Graph scopes for managing inbox rules: `MailboxSettings.ReadWrite`
 * covers `/me/mailFolders/inbox/messageRules`; `Mail.ReadBasic` is the least
 * privilege that lists `/me/mailFolders` (needed to turn a folder path into
 * the id a rule's `moveToFolder` wants). `offline_access` keeps the refresh
 * token covering this resource too.
 */
export const MICROSOFT_GRAPH_RULES_SCOPES = [
  'https://graph.microsoft.com/MailboxSettings.ReadWrite',
  'https://graph.microsoft.com/Mail.ReadBasic',
  'offline_access',
];

/**
 * Whether `granted` (as returned by Microsoft in a token response) covers
 * `wanted`. Microsoft reports Exchange scopes with their resource URI
 * (`https://outlook.office.com/IMAP.AccessAsUser.All`) but Graph scopes in
 * short form (`Mail.ReadBasic`), so the comparison accepts either spelling:
 * the full URI, or the last path segment, case-insensitively.
 */
export function scopeGranted(granted: readonly string[] | undefined, wanted: string): boolean {
  if (!granted?.length) return false;
  const short = (scope: string) => scope.slice(scope.lastIndexOf('/') + 1).toLowerCase();
  const wantedFull = wanted.toLowerCase();
  const wantedShort = short(wanted);
  return granted.some(g => g.toLowerCase() === wantedFull || short(g) === wantedShort);
}

/** Union of two scope lists, first occurrence wins, de-duplicated case-insensitively. */
export function mergeScopes(...lists: (readonly string[] | undefined)[]): string[] {
  const out: string[] = [];
  for (const list of lists) {
    for (const scope of list ?? []) {
      if (!out.some(s => s.toLowerCase() === scope.toLowerCase())) out.push(scope);
    }
  }
  return out;
}

/** True when every non-`offline_access` scope in `wanted` has been granted. */
export function allScopesGranted(granted: readonly string[] | undefined, wanted: readonly string[]): boolean {
  return wanted.filter(s => s !== 'offline_access').every(s => scopeGranted(granted, s));
}

/**
 * Thrown when a refresh for a scope set fails because the user has not
 * consented to it (Entra `AADSTS65001` / `consent_required`, or a refresh
 * token that does not cover the resource). Callers map it to a structured
 * "run the authorize tool" response instead of a generic failure.
 */
export class ConsentRequiredError extends Error {
  readonly code = 'consent_required' as const;
  constructor(message: string, readonly scopes: readonly string[], readonly accountId: string) {
    super(message);
    this.name = 'ConsentRequiredError';
  }
}

/** Personal Microsoft accounts (Outlook.com / Hotmail / Live / MSN). */
export const DEFAULT_MICROSOFT_TENANT = 'consumers';

/** Env var supplying the Entra application (client) ID when a tool call omits it. */
export const MS_CLIENT_ID_ENV = 'IMAP_MCP_MS_CLIENT_ID';

/** Refresh when the cached access token has less than this long to live. */
const REFRESH_MARGIN_MS = 5 * 60 * 1000;

/** Cap on how long a single `pollDeviceCode` call blocks (MCP clients time out). */
export const DEFAULT_POLL_MAX_WAIT_MS = 25_000;

export interface DeviceCodeStart {
  /** Opaque handle for `pollDeviceCode`. Random; carries no secret. */
  flowId: string;
  /** Code the user types at `verificationUri`. */
  userCode: string;
  verificationUri: string;
  /** Epoch ms after which the code is no longer valid. */
  expiresAt: number;
  /** Microsoft's human-readable instruction line. */
  message: string;
  /** Minimum seconds between token-endpoint polls. */
  interval: number;
}

export interface OAuthTokens {
  accessToken: string;
  refreshToken: string;
  /** Epoch ms when `accessToken` expires. */
  accessTokenExpiresAt: number;
  scopes: string[];
}

export type DeviceCodePollResult<TContext = unknown> =
  | { status: 'pending'; retryAfterSeconds: number; expiresAt: number }
  | { status: 'complete'; tokens: OAuthTokens; context: TContext | undefined }
  | { status: 'expired' | 'denied' | 'error'; error: string };

export interface RefreshResult {
  accessToken: string;
  accessTokenExpiresAt: number;
  /** Present when Microsoft rotated the refresh token. Persist it. */
  refreshToken?: string;
}

interface PendingFlow {
  deviceCode: string;
  clientId: string;
  tenant: string;
  interval: number;
  expiresAt: number;
  context: unknown;
}

interface TokenResponse {
  access_token?: string;
  refresh_token?: string;
  expires_in?: number;
  scope?: string;
  error?: string;
  error_description?: string;
}

interface DeviceCodeResponse {
  device_code?: string;
  user_code?: string;
  verification_uri?: string;
  expires_in?: number;
  interval?: number;
  message?: string;
  error?: string;
  error_description?: string;
}

export interface OAuthServiceOptions {
  /** Injectable for tests; defaults to a real timer. */
  sleep?: (ms: number) => Promise<void>;
  /** Injectable clock for tests. */
  now?: () => number;
  /** Injectable HTTP client for tests; defaults to the global `fetch`. */
  fetch?: typeof fetch;
}

/** Entra tenants are `consumers` / `common` / `organizations`, a GUID, or a verified domain. */
export function isValidTenant(tenant: string): boolean {
  return /^[A-Za-z0-9][A-Za-z0-9.-]{0,254}$/.test(tenant);
}

/**
 * Resolve the client ID for a device-code flow: explicit argument first, then
 * the `IMAP_MCP_MS_CLIENT_ID` environment variable. Throws with setup guidance
 * when neither is present.
 */
export function resolveMicrosoftClientId(explicit?: string, env: NodeJS.ProcessEnv = process.env): string {
  const clientId = (explicit ?? env[MS_CLIENT_ID_ENV] ?? '').trim();
  if (clientId) return clientId;
  throw new Error(
    'No Microsoft client ID available. Pass clientId, or set the ' +
    `${MS_CLIENT_ID_ENV} environment variable for the MCP server. The client ID is the ` +
    '"Application (client) ID" of an Entra app registration with "Allow public client flows" ' +
    'enabled and the delegated Office 365 Exchange Online permissions IMAP.AccessAsUser.All ' +
    'and SMTP.Send — see the README section "Outlook.com / Microsoft 365 (OAuth 2.0)".'
  );
}

export class MicrosoftOAuthService {
  private flows: Map<string, PendingFlow> = new Map();
  /** In-flight refreshes keyed by account id + scope set, so IMAP and SMTP share one refresh. */
  private refreshes: Map<string, Promise<string>> = new Map();
  /**
   * Access tokens for non-mail scope sets (e.g. Graph), keyed like
   * `refreshes`. Memory only: the persisted `oauth.accessToken` stays the
   * mail token, and a Graph token is cheap to mint again after a restart.
   */
  private resourceTokens: Map<string, { accessToken: string; expiresAt: number }> = new Map();
  private accountManager?: AccountManager;
  private readonly sleep: (ms: number) => Promise<void>;
  private readonly now: () => number;
  private readonly fetchImpl: typeof fetch;

  constructor(accountManager?: AccountManager, options: OAuthServiceOptions = {}) {
    this.accountManager = accountManager;
    this.sleep = options.sleep ?? ((ms) => new Promise(resolve => setTimeout(resolve, ms)));
    this.now = options.now ?? (() => Date.now());
    this.fetchImpl = options.fetch ?? ((input, init) => fetch(input, init));
  }

  setAccountManager(accountManager: AccountManager): void {
    this.accountManager = accountManager;
  }

  private endpoint(tenant: string, path: 'devicecode' | 'token'): string {
    if (!isValidTenant(tenant)) {
      throw new Error(`Invalid Entra tenant "${tenant}". Use consumers, common, organizations, a tenant GUID, or a verified domain.`);
    }
    return `${MICROSOFT_LOGIN_HOST}/${tenant}/oauth2/v2.0/${path}`;
  }

  private async postForm<T>(url: string, params: Record<string, string>): Promise<{ ok: boolean; status: number; body: T }> {
    const res = await this.fetchImpl(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded', Accept: 'application/json' },
      body: new URLSearchParams(params).toString(),
    });
    let body: T;
    try {
      body = await res.json() as T;
    } catch {
      body = {} as T;
    }
    return { ok: res.ok, status: res.status, body };
  }

  /**
   * Begin a device-code authorization. The returned `flowId` is a random
   * handle; the `device_code` itself stays in memory and is never returned.
   * `context` is stored alongside and handed back when the flow completes, so
   * a caller can remember what account the tokens are for.
   */
  async startDeviceCode<TContext = unknown>(opts: {
    clientId: string;
    tenant?: string;
    scopes?: string[];
    context?: TContext;
  }): Promise<DeviceCodeStart> {
    const tenant = opts.tenant || DEFAULT_MICROSOFT_TENANT;
    const scopes = opts.scopes ?? MICROSOFT_OAUTH_SCOPES;

    const { ok, status, body } = await this.postForm<DeviceCodeResponse>(this.endpoint(tenant, 'devicecode'), {
      client_id: opts.clientId,
      scope: scopes.join(' '),
    });

    if (!ok || !body.device_code || !body.user_code || !body.verification_uri) {
      const detail = body.error_description || body.error || `HTTP ${status}`;
      throw new Error(
        `Microsoft device-code request failed: ${detail}. Check that the client ID belongs to an ` +
        `Entra app registration with "Allow public client flows" enabled and that tenant "${tenant}" ` +
        'matches its supported account types.'
      );
    }

    const flowId = crypto.randomBytes(16).toString('hex');
    const interval = Math.max(1, body.interval ?? 5);
    const expiresAt = this.now() + (body.expires_in ?? 900) * 1000;

    this.flows.set(flowId, {
      deviceCode: body.device_code,
      clientId: opts.clientId,
      tenant,
      interval,
      expiresAt,
      context: opts.context,
    });

    return {
      flowId,
      userCode: body.user_code,
      verificationUri: body.verification_uri,
      expiresAt,
      message: body.message ?? `To sign in, open ${body.verification_uri} and enter the code ${body.user_code}.`,
      interval,
    };
  }

  /** Whether a flow started with `startDeviceCode` is still awaiting the user. */
  hasPendingFlow(flowId: string): boolean {
    return this.flows.has(flowId);
  }

  /**
   * Poll the token endpoint for a pending device-code flow. Honors the
   * server-provided interval and `slow_down`. Returns `pending` when the user
   * has not finished within `maxWaitMs` — call again. Terminal results remove
   * the flow.
   */
  async pollDeviceCode<TContext = unknown>(
    flowId: string,
    opts: { maxWaitMs?: number } = {},
  ): Promise<DeviceCodePollResult<TContext>> {
    const flow = this.flows.get(flowId);
    if (!flow) {
      return {
        status: 'error',
        error: `Unknown flowId "${flowId}". The flow may have completed, expired, or been started by a previous server process. Start again with imap_add_oauth_account.`,
      };
    }

    const maxWaitMs = opts.maxWaitMs ?? DEFAULT_POLL_MAX_WAIT_MS;
    const deadline = this.now() + maxWaitMs;
    const tokenUrl = this.endpoint(flow.tenant, 'token');

    for (;;) {
      if (this.now() >= flow.expiresAt) {
        this.flows.delete(flowId);
        return { status: 'expired', error: 'The device code expired before the sign-in was completed. Start again with imap_add_oauth_account.' };
      }

      const { ok, status, body } = await this.postForm<TokenResponse>(tokenUrl, {
        grant_type: 'urn:ietf:params:oauth:grant-type:device_code',
        client_id: flow.clientId,
        device_code: flow.deviceCode,
      });

      if (ok && body.access_token) {
        this.flows.delete(flowId);
        if (!body.refresh_token) {
          return {
            status: 'error',
            error: 'Microsoft returned an access token but no refresh token. Make sure the offline_access scope is granted (it is requested by default) and try again.',
          };
        }
        return {
          status: 'complete',
          tokens: {
            accessToken: body.access_token,
            refreshToken: body.refresh_token,
            accessTokenExpiresAt: this.now() + (body.expires_in ?? 3600) * 1000,
            scopes: body.scope ? body.scope.split(' ').filter(Boolean) : MICROSOFT_OAUTH_SCOPES,
          },
          context: flow.context as TContext | undefined,
        };
      }

      switch (body.error) {
        case 'authorization_pending':
          break;
        case 'slow_down':
          // RFC 8628 §3.5: increase the interval by 5 seconds.
          flow.interval += 5;
          break;
        case 'expired_token':
          this.flows.delete(flowId);
          return { status: 'expired', error: 'The device code expired before the sign-in was completed. Start again with imap_add_oauth_account.' };
        case 'authorization_declined':
          this.flows.delete(flowId);
          return { status: 'denied', error: 'The user declined the authorization request.' };
        default: {
          this.flows.delete(flowId);
          const detail = body.error_description || body.error || `HTTP ${status}`;
          return { status: 'error', error: `Microsoft token request failed: ${detail}` };
        }
      }

      const waitMs = flow.interval * 1000;
      if (this.now() + waitMs > deadline) {
        return { status: 'pending', retryAfterSeconds: flow.interval, expiresAt: flow.expiresAt };
      }
      await this.sleep(waitMs);
    }
  }

  /** The scopes an account's mail tokens are minted with. */
  private mailScopes(oauth: OAuthConfig): string[] {
    return oauth.scopes?.length ? oauth.scopes : MICROSOFT_OAUTH_SCOPES;
  }

  /** Cache / in-flight key for one (account, scope set). Order-insensitive. */
  private scopeKey(accountId: string, scopes: readonly string[]): string {
    return `${accountId}|${[...new Set(scopes.map(s => s.toLowerCase()))].sort().join(' ')}`;
  }

  /**
   * Exchange the account's refresh token for a new access token. Microsoft
   * may rotate the refresh token; when it does, the new one is returned too.
   *
   * `scopes` selects the resource the token is for and defaults to the mail
   * scopes. Any other set (Graph) is a second resource the user must have
   * consented to; a rejection there surfaces as `ConsentRequiredError` so
   * the caller can point at the authorize tool rather than at a full
   * re-authorization of the mailbox.
   */
  async refreshAccessToken(account: ImapAccount, scopes?: readonly string[]): Promise<RefreshResult> {
    const oauth = this.requireOAuth(account);
    if (!oauth.refreshToken) {
      throw new Error(
        `Account "${account.name}" has no OAuth refresh token. Re-authorize it with imap_add_oauth_account (accountId: ${account.id}).`
      );
    }
    const requested = scopes ?? this.mailScopes(oauth);
    const isMail = this.scopeKey(account.id, requested) === this.scopeKey(account.id, this.mailScopes(oauth));

    const { ok, status, body } = await this.postForm<TokenResponse>(this.endpoint(oauth.tenant, 'token'), {
      grant_type: 'refresh_token',
      client_id: oauth.clientId,
      refresh_token: oauth.refreshToken,
      scope: requested.join(' '),
    });

    if (!ok || !body.access_token) {
      const detail = body.error_description || body.error || `HTTP ${status}`;
      const consentCodes = ['invalid_grant', 'interaction_required', 'consent_required', 'invalid_scope'];
      if (!isMail && consentCodes.includes(body.error ?? '')) {
        const list = requested.filter(s => s !== 'offline_access').join(', ');
        throw new ConsentRequiredError(
          `Microsoft would not issue a token for ${list} on account "${account.name}" (${body.error}). ` +
          'The user has not consented to these scopes yet (or the consent was revoked); the mailbox sign-in itself is unaffected. ' +
          `Details: ${detail}`,
          requested,
          account.id,
        );
      }
      if (body.error === 'invalid_grant' || body.error === 'interaction_required') {
        throw new Error(
          `The OAuth refresh token for account "${account.name}" was rejected (${body.error}). ` +
          'It has expired, been revoked, or the user must sign in again. ' +
          `Re-authorize with imap_add_oauth_account (accountId: ${account.id}), then imap_complete_oauth_login. ` +
          `Details: ${detail}`
        );
      }
      throw new Error(`Failed to refresh the OAuth access token for account "${account.name}": ${detail}`);
    }

    return {
      accessToken: body.access_token,
      accessTokenExpiresAt: this.now() + (body.expires_in ?? 3600) * 1000,
      ...(body.refresh_token ? { refreshToken: body.refresh_token } : {}),
    };
  }

  /**
   * Return an access token that is good for at least five more minutes,
   * refreshing (and persisting the result through the AccountManager) when
   * the cached one is missing or about to expire. Mutates `account.oauth` so
   * a caller holding the object sees the fresh token. Concurrent callers for
   * the same account share a single refresh.
   */
  async getValidAccessToken(account: ImapAccount, scopes?: readonly string[]): Promise<string> {
    const oauth = this.requireOAuth(account);
    const requested = scopes ?? this.mailScopes(oauth);
    const key = this.scopeKey(account.id, requested);
    const isMail = key === this.scopeKey(account.id, this.mailScopes(oauth));

    if (isMail) {
      if (
        oauth.accessToken &&
        oauth.accessTokenExpiresAt !== undefined &&
        oauth.accessTokenExpiresAt - this.now() > REFRESH_MARGIN_MS
      ) {
        return oauth.accessToken;
      }
    } else {
      const cached = this.resourceTokens.get(key);
      if (cached && cached.expiresAt - this.now() > REFRESH_MARGIN_MS) {
        return cached.accessToken;
      }
    }
    return this.forceRefresh(account, requested);
  }

  /**
   * Seed the in-memory cache for a non-mail scope set with a token that was
   * just minted elsewhere (the device-code flow for Graph consent returns
   * one), so the first Graph call after consenting needs no refresh. Mail
   * scopes are ignored here — those go through the account store.
   */
  primeAccessToken(account: ImapAccount, scopes: readonly string[], accessToken: string, expiresAt: number): void {
    const oauth = this.requireOAuth(account);
    const key = this.scopeKey(account.id, scopes);
    if (key === this.scopeKey(account.id, this.mailScopes(oauth))) return;
    this.resourceTokens.set(key, { accessToken, expiresAt });
  }

  /** Drop every cached non-mail token for an account (after re-authorization or removal). */
  forgetResourceTokens(accountId: string): void {
    for (const key of [...this.resourceTokens.keys()]) {
      if (key.startsWith(`${accountId}|`)) this.resourceTokens.delete(key);
    }
  }

  /**
   * Refresh regardless of the cached token's age — used after the server
   * rejected a token that looked valid (revoked, or clock skew).
   */
  async forceRefresh(account: ImapAccount, scopes?: readonly string[]): Promise<string> {
    const oauth = this.requireOAuth(account);
    const requested = scopes ?? this.mailScopes(oauth);
    const key = this.scopeKey(account.id, requested);
    const isMail = key === this.scopeKey(account.id, this.mailScopes(oauth));

    const inFlight = this.refreshes.get(key);
    if (inFlight) return inFlight;

    const run = (async () => {
      const result = await this.refreshAccessToken(account, requested);
      if (isMail) {
        oauth.accessToken = result.accessToken;
        oauth.accessTokenExpiresAt = result.accessTokenExpiresAt;
      } else {
        this.resourceTokens.set(key, { accessToken: result.accessToken, expiresAt: result.accessTokenExpiresAt });
      }
      if (result.refreshToken) {
        oauth.refreshToken = result.refreshToken;
      }

      // Mail tokens are persisted as before. For any other resource only a
      // rotated refresh token is worth writing — the access token is memory-only.
      if (this.accountManager && (isMail || result.refreshToken)) {
        try {
          await this.accountManager.updateOAuthTokens(account.id, {
            ...(isMail ? { accessToken: result.accessToken, accessTokenExpiresAt: result.accessTokenExpiresAt } : {}),
            ...(result.refreshToken ? { refreshToken: result.refreshToken } : {}),
          });
        } catch (err) {
          // Not fatal for this connection — the in-memory token works — but a
          // rotated refresh token that failed to persist means the next server
          // start will need a re-authorization. Log the account, never the token.
          console.error(
            `[oauth] Could not persist refreshed tokens for account ${account.id}: ${err instanceof Error ? err.message : String(err)}`
          );
        }
      }

      return result.accessToken;
    })();

    this.refreshes.set(key, run);
    try {
      return await run;
    } finally {
      this.refreshes.delete(key);
    }
  }

  private requireOAuth(account: ImapAccount): OAuthConfig {
    if (account.authType !== 'oauth2' || !account.oauth) {
      throw new Error(`Account "${account.name}" is not configured for OAuth 2.0.`);
    }
    if (account.oauth.provider !== 'microsoft') {
      throw new Error(`Unsupported OAuth provider "${account.oauth.provider}" on account "${account.name}".`);
    }
    return account.oauth;
  }
}
