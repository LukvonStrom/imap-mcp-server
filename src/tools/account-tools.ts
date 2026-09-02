import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { AccountManager } from '../services/account-manager.js';
import { ImapService } from '../services/imap-service.js';
import { SmtpService } from '../services/smtp-service.js';
import {
  MicrosoftOAuthService,
  DEFAULT_MICROSOFT_TENANT,
  DEFAULT_POLL_MAX_WAIT_MS,
  MS_CLIENT_ID_ENV,
  isValidTenant,
  resolveMicrosoftClientId,
} from '../services/oauth-service.js';
import { getProviderById } from '../providers/email-providers.js';
import type { ImapAccount, OAuthConfig } from '../types/index.js';
import { z } from 'zod';

/**
 * What imap_add_oauth_account remembers between starting the device-code flow
 * and imap_complete_oauth_login storing the account. Held in memory by the
 * OAuth service, keyed by the flowId. Contains no secrets.
 */
interface PendingOAuthAccount {
  accountId?: string;
  name: string;
  email: string;
  host: string;
  port: number;
  smtpHost: string;
  smtpPort: number;
  clientId: string;
  tenant: string;
  sentFolder?: string;
  defaultBcc?: string | string[];
}

/** The non-secret part of an account's OAuth config, safe to return from tools. */
function publicOAuth(oauth: OAuthConfig | undefined) {
  if (!oauth) return undefined;
  return { provider: oauth.provider, tenant: oauth.tenant, clientId: oauth.clientId, scopes: oauth.scopes };
}

export function accountTools(
  server: McpServer,
  accountManager: AccountManager,
  imapService: ImapService,
  smtpService: SmtpService,
  oauthService: MicrosoftOAuthService = new MicrosoftOAuthService(accountManager)
): void {
  // Add account tool
  server.registerTool('imap_add_account', {
    description: 'Add a new IMAP account configuration',
    inputSchema: {
      name: z.string().describe('Friendly name for the account'),
      host: z.string().describe('IMAP server hostname'),
      port: z.coerce.number().default(993).describe('IMAP server port (default: 993)'),
      user: z.string().describe('Username for authentication'),
      password: z.string().describe('Password for authentication'),
      tls: z.boolean().default(true).describe('Use TLS/SSL (default: true)'),
      allowStartTLS: z.boolean().optional().describe('When tls is false, imapflow still opportunistically upgrades via STARTTLS if the server advertises it, validating the cert against `host` regardless of `tls`. Set to false to disable that upgrade and stay on the plain connection — needed for providers (e.g. some DreamHost mail hosting) that advertise STARTTLS on a hostname covered only by a shared/wildcard cert. Defaults to true. WARNING: combined with tls:false this sends the password and all mail in cleartext — only use on a trusted or local network, and try fixing tls/host first.'),
      tlsRejectUnauthorized: z.boolean().optional().describe('Reject unauthorized TLS certificates (self-signed, expired, etc.). Set to false to allow connections to servers with invalid certificates. Default: true. WARNING: disabling validation allows man-in-the-middle interception of the password and all mail; only use for a trusted internal/self-hosted server, never for a public provider.'),
      email: z.string().optional().describe('Email address (From: header). Defaults to user if omitted'),
      smtpHost: z.string().optional().describe('SMTP server hostname. Defaults to IMAP host with imap.→smtp. rewrite'),
      smtpPort: z.coerce.number().optional().describe('SMTP server port (465 for SMTPS, 587 for STARTTLS). Defaults to 587'),
      smtpSecure: z.boolean().optional().describe('Use implicit TLS (SMTPS). Ignored for port 587/25 which always use STARTTLS, and for port 465 which always uses implicit TLS'),
      sentFolder: z.string().optional().describe('Explicit Sent-folder name for saving sent-mail copies (e.g. "Gesendet"). Only needed when auto-detection fails — the server must lack a \\Sent SPECIAL-USE folder. Check names with imap_list_folders'),
      defaultBcc: z.union([z.string(), z.array(z.string())]).optional().describe('Optional BCC address(es) applied automatically to every outbound send, reply, forward, and draft for this account. Merged with any per-call bcc'),
    }
  }, async ({ name, host, port, user, password, tls, allowStartTLS, tlsRejectUnauthorized, email, smtpHost, smtpPort, smtpSecure, sentFolder, defaultBcc }) => {
    const smtp = (smtpHost || smtpPort !== undefined || smtpSecure !== undefined)
      ? {
          host: smtpHost || host,
          port: smtpPort ?? 587,
          secure: smtpSecure ?? false,
        }
      : undefined;

    const account = await accountManager.addAccount({
      name,
      host,
      port,
      user,
      password,
      tls,
      ...(allowStartTLS !== undefined ? { allowStartTLS } : {}),
      ...(tlsRejectUnauthorized !== undefined ? { tlsRejectUnauthorized } : {}),
      ...(email ? { email } : {}),
      ...(smtp ? { smtp } : {}),
      ...(sentFolder ? { sentFolder } : {}),
      ...(defaultBcc !== undefined && defaultBcc !== '' && !(Array.isArray(defaultBcc) && defaultBcc.length === 0)
        ? { defaultBcc }
        : {}),
    });

    return {
      content: [{
        type: 'text',
        text: JSON.stringify({
          success: true,
          accountId: account.id,
          message: `Account "${name}" added successfully`,
        }, null, 2)
      }]
    };
  });

  // Update account tool — lets callers fix SMTP config (and other fields) on existing accounts
  server.registerTool('imap_update_account', {
    description: 'Update an existing IMAP account. Useful for fixing SMTP settings without removing and re-adding the account.',
    inputSchema: {
      accountId: z.string().describe('ID of the account to update'),
      name: z.string().optional().describe('New friendly name'),
      host: z.string().optional().describe('IMAP host'),
      port: z.coerce.number().optional().describe('IMAP port'),
      user: z.string().optional().describe('IMAP username'),
      password: z.string().optional().describe('New password. Rejected for OAuth 2.0 accounts (authType "oauth2") — re-authorize those with imap_add_oauth_account instead'),
      tls: z.boolean().optional().describe('Use TLS for IMAP'),
      allowStartTLS: z.boolean().optional().describe('When tls is false, set to false to also disable imapflow\'s opportunistic STARTTLS upgrade (see imap_add_account). Defaults to true. WARNING: with tls:false this means a fully cleartext session (password included) — only on a trusted network.'),
      tlsRejectUnauthorized: z.boolean().optional().describe('Reject unauthorized TLS certificates (self-signed, expired, etc.). Set to false to allow connections to servers with invalid certificates. WARNING: disabling validation allows man-in-the-middle interception of the password and all mail; only use for a trusted internal/self-hosted server, never for a public provider.'),
      email: z.string().optional().describe('Email address (From: header)'),
      smtpHost: z.string().optional().describe('SMTP hostname'),
      smtpPort: z.coerce.number().optional().describe('SMTP port (465 for SMTPS, 587 for STARTTLS)'),
      smtpSecure: z.boolean().optional().describe('Use implicit TLS (SMTPS). Port 587/25 always use STARTTLS regardless'),
      smtpUser: z.string().optional().describe('SMTP username (if different from IMAP user)'),
      smtpPassword: z.string().optional().describe('SMTP password (if different from IMAP password)'),
      saveToSent: z.boolean().optional().describe('Save sent emails to the Sent folder'),
      sentFolder: z.string().optional().describe('Explicit Sent-folder name for saving sent-mail copies (e.g. "Gesendet"). Overrides auto-detection; pass an empty string to clear the override and re-enable auto-detection. Check names with imap_list_folders'),
      defaultBcc: z.union([z.string(), z.array(z.string())]).optional().describe('Optional BCC address(es) applied automatically to every outbound message for this account. Pass an empty string to clear'),
    }
  }, async ({ accountId, name, host, port, user, password, tls, allowStartTLS, tlsRejectUnauthorized, email, smtpHost, smtpPort, smtpSecure, smtpUser, smtpPassword, saveToSent, sentFolder, defaultBcc }) => {
    const existing = accountManager.getAccount(accountId);
    if (!existing) {
      throw new Error(`Account ${accountId} not found`);
    }

    if (existing.authType === 'oauth2' && (password !== undefined || smtpPassword !== undefined)) {
      throw new Error(
        `Account "${existing.name}" authenticates with OAuth 2.0 and has no password; Microsoft does not accept ` +
        'passwords or app passwords for IMAP/SMTP. To re-authorize it, run imap_add_oauth_account with ' +
        `accountId "${accountId}" and finish with imap_complete_oauth_login.`
      );
    }

    const updates: any = {};
    if (name !== undefined) updates.name = name;
    if (host !== undefined) updates.host = host;
    if (port !== undefined) updates.port = port;
    if (user !== undefined) updates.user = user;
    if (password !== undefined) updates.password = password;
    if (tls !== undefined) updates.tls = tls;
    if (allowStartTLS !== undefined) updates.allowStartTLS = allowStartTLS;
    if (tlsRejectUnauthorized !== undefined) updates.tlsRejectUnauthorized = tlsRejectUnauthorized;
    if (email !== undefined) updates.email = email;
    if (saveToSent !== undefined) updates.saveToSent = saveToSent;
    // Empty string clears the override (falls back to auto-detection).
    if (sentFolder !== undefined) updates.sentFolder = sentFolder === '' ? undefined : sentFolder;
    // Empty string (or empty array) clears the default BCC.
    if (defaultBcc !== undefined) {
      if (defaultBcc === '' || (Array.isArray(defaultBcc) && defaultBcc.length === 0)) {
        updates.defaultBcc = undefined;
      } else {
        updates.defaultBcc = defaultBcc;
      }
    }

    const smtpTouched = [smtpHost, smtpPort, smtpSecure, smtpUser, smtpPassword].some(v => v !== undefined);
    if (smtpTouched) {
      const current = existing.smtp;
      updates.smtp = {
        host: smtpHost ?? current?.host ?? existing.host,
        port: smtpPort ?? current?.port ?? 587,
        secure: smtpSecure ?? current?.secure ?? false,
        ...(smtpUser !== undefined ? { user: smtpUser } : current?.user ? { user: current.user } : {}),
        ...(smtpPassword !== undefined ? { password: smtpPassword } : {}),
      };
    }

    // Invalidate any cached SMTP transporter so the next send picks up new config
    if (smtpTouched) {
      smtpService.disconnect(accountId);
    }

    const updated = await accountManager.updateAccount(accountId, updates);

    return {
      content: [{
        type: 'text',
        text: JSON.stringify({
          success: true,
          accountId: updated.id,
          message: `Account "${updated.name}" updated`,
          smtp: updated.smtp ? { host: updated.smtp.host, port: updated.smtp.port, secure: updated.smtp.secure } : undefined,
        }, null, 2)
      }]
    };
  });

  // List accounts tool
  server.registerTool('imap_list_accounts', {
    description: 'List all configured IMAP accounts. Each entry reports authType ("password" or "oauth2"); OAuth accounts also show their provider, tenant, and client ID. Tokens and passwords are never returned.',
    inputSchema: {}
  }, async () => {
    const accounts = accountManager.getAllAccounts();

    return {
      content: [{
        type: 'text',
        text: JSON.stringify({
          accounts: accounts.map(acc => ({
            id: acc.id,
            name: acc.name,
            host: acc.host,
            port: acc.port,
            user: acc.user,
            tls: acc.tls,
            ...(acc.allowStartTLS === false ? { allowStartTLS: false } : {}),
            ...(acc.tlsRejectUnauthorized === false ? { tlsRejectUnauthorized: false } : {}),
            authType: acc.authType ?? 'password',
            ...(acc.authType === 'oauth2' && acc.oauth ? { oauth: publicOAuth(acc.oauth) } : {}),
          })),
        }, null, 2)
      }]
    };
  });

  // ---------------------------------------------------------------------------
  // OAuth 2.0 (Microsoft device-code flow)
  // ---------------------------------------------------------------------------

  server.registerTool('imap_add_oauth_account', {
    description: 'Start adding an Outlook.com / Hotmail / Live / Microsoft 365 mailbox using OAuth 2.0 (device-code flow). Microsoft no longer accepts passwords or app passwords for IMAP/SMTP, so use this instead of imap_add_account for those providers. Step 1 of 2: this returns a verificationUri and a short userCode — show BOTH to the user and ask them to open the URL in a browser, enter the code, and sign in. Then call imap_complete_oauth_login with the returned flowId (step 2); it waits for the sign-in and stores the account. Requires the Application (client) ID of an Entra app registration with "Allow public client flows" enabled and delegated Office 365 Exchange Online permissions IMAP.AccessAsUser.All and SMTP.Send (see README "Outlook.com / Microsoft 365 (OAuth 2.0)"). Pass accountId to re-authorize an existing account (e.g. after "refresh token was rejected").',
    inputSchema: {
      name: z.string().optional().describe('Friendly name for the account (e.g. "Personal Outlook"). Defaults to the email address. When accountId is given, supplying a name renames that account; omit it to keep the current name'),
      email: z.string().describe('The mailbox address (e.g. user@outlook.com). Used as the IMAP/SMTP username and as the From: address'),
      provider: z.enum(['microsoft']).default('microsoft').describe('OAuth provider. Only "microsoft" (Outlook.com, Hotmail, Live, MSN, Microsoft 365) is supported'),
      clientId: z.string().optional().describe(`Application (client) ID of the Entra app registration. Optional when the server has the ${MS_CLIENT_ID_ENV} environment variable set; required otherwise`),
      tenant: z.string().optional().describe('Entra tenant to sign in against. "consumers" (default) for personal Outlook.com / Hotmail / Live accounts; "organizations" or the tenant GUID / verified domain (e.g. contoso.onmicrosoft.com) for Microsoft 365 work accounts; "common" when the app registration allows both'),
      accountId: z.string().optional().describe('ID of an existing account to re-authorize or convert to OAuth 2.0. The stored tokens are replaced when the flow completes; other settings are kept unless overridden here'),
      host: z.string().optional().describe('IMAP hostname override. Defaults to outlook.office365.com (port 993, TLS) — only change for a non-standard Exchange endpoint'),
      port: z.coerce.number().optional().describe('IMAP port override (default 993)'),
      smtpHost: z.string().optional().describe('SMTP hostname override. Defaults to smtp-mail.outlook.com (port 587, STARTTLS); smtp.office365.com also works for Microsoft 365'),
      smtpPort: z.coerce.number().optional().describe('SMTP port override (default 587, STARTTLS)'),
      sentFolder: z.string().optional().describe('Explicit Sent-folder name. Normally unnecessary — Outlook advertises "Sent Items" via SPECIAL-USE and it is auto-detected'),
      defaultBcc: z.union([z.string(), z.array(z.string())]).optional().describe('Optional BCC address(es) applied automatically to every outbound message for this account'),
    }
  }, async ({ name, email, provider, clientId, tenant, accountId, host, port, smtpHost, smtpPort, sentFolder, defaultBcc }) => {
    if (provider !== 'microsoft') {
      throw new Error(`Unsupported OAuth provider "${provider}"`);
    }

    let existing: ImapAccount | undefined;
    if (accountId) {
      existing = accountManager.getAccount(accountId);
      if (!existing) {
        throw new Error(`Account ${accountId} not found. Use imap_list_accounts to see available accounts.`);
      }
    }

    // Provider defaults: the Outlook entry carries the Microsoft IMAP/SMTP endpoints.
    const outlook = getProviderById('outlook');
    const resolvedClientId = resolveMicrosoftClientId(clientId ?? existing?.oauth?.clientId);
    const resolvedTenant = tenant ?? existing?.oauth?.tenant ?? outlook?.oauthTenant ?? DEFAULT_MICROSOFT_TENANT;
    if (!isValidTenant(resolvedTenant)) {
      throw new Error(`Invalid tenant "${resolvedTenant}". Use consumers, common, organizations, a tenant GUID, or a verified domain.`);
    }

    const pending: PendingOAuthAccount = {
      accountId,
      name: name || existing?.name || email,
      email,
      host: host ?? existing?.host ?? outlook?.imapHost ?? 'outlook.office365.com',
      port: port ?? existing?.port ?? outlook?.imapPort ?? 993,
      smtpHost: smtpHost ?? existing?.smtp?.host ?? outlook?.smtpHost ?? 'smtp-mail.outlook.com',
      smtpPort: smtpPort ?? existing?.smtp?.port ?? outlook?.smtpPort ?? 587,
      clientId: resolvedClientId,
      tenant: resolvedTenant,
      ...(sentFolder ? { sentFolder } : {}),
      ...(defaultBcc !== undefined && defaultBcc !== '' && !(Array.isArray(defaultBcc) && defaultBcc.length === 0)
        ? { defaultBcc }
        : {}),
    };

    const start = await oauthService.startDeviceCode<PendingOAuthAccount>({
      clientId: resolvedClientId,
      tenant: resolvedTenant,
      context: pending,
    });

    return {
      content: [{
        type: 'text',
        text: JSON.stringify({
          status: 'awaiting_user',
          flowId: start.flowId,
          userCode: start.userCode,
          verificationUri: start.verificationUri,
          expiresAt: new Date(start.expiresAt).toISOString(),
          message: start.message,
          instructions:
            `Show the user this URL and code: open ${start.verificationUri} in a browser, enter the code ` +
            `${start.userCode}, and sign in to ${email}. The code expires at ${new Date(start.expiresAt).toISOString()}. ` +
            `Then call imap_complete_oauth_login with flowId "${start.flowId}". It waits up to ` +
            `${Math.round(DEFAULT_POLL_MAX_WAIT_MS / 1000)} seconds; if it returns status "pending", call it again with the same flowId.`,
        }, null, 2)
      }]
    };
  });

  server.registerTool('imap_complete_oauth_login', {
    description: 'Step 2 of the OAuth 2.0 sign-in started by imap_add_oauth_account. Waits (up to ~25 seconds) for the user to finish signing in at the verification URL, then stores the account with its encrypted tokens and runs a connection test. Returns status "pending" if the user has not finished yet — call again with the same flowId until it returns "complete", "expired", "denied", or "error". Never returns tokens.',
    inputSchema: {
      flowId: z.string().describe('The flowId returned by imap_add_oauth_account'),
      maxWaitSeconds: z.coerce.number().min(1).max(25).optional().describe('How long this call may wait for the sign-in before returning "pending" (1–25 seconds, default 25). Lower it if your MCP client times out'),
    }
  }, async ({ flowId, maxWaitSeconds }) => {
    const result = await oauthService.pollDeviceCode<PendingOAuthAccount>(flowId, {
      maxWaitMs: maxWaitSeconds !== undefined ? maxWaitSeconds * 1000 : DEFAULT_POLL_MAX_WAIT_MS,
    });

    if (result.status === 'pending') {
      return {
        content: [{
          type: 'text',
          text: JSON.stringify({
            status: 'pending',
            retryAfterSeconds: result.retryAfterSeconds,
            expiresAt: new Date(result.expiresAt).toISOString(),
            message: 'The user has not finished signing in yet. Make sure they opened the verification URL and entered the code, then call imap_complete_oauth_login again with the same flowId.',
          }, null, 2)
        }]
      };
    }

    if (result.status !== 'complete') {
      return {
        content: [{
          type: 'text',
          text: JSON.stringify({
            status: result.status,
            success: false,
            error: result.error,
            message: 'Start over with imap_add_oauth_account to get a new code.',
          }, null, 2)
        }]
      };
    }

    const pending = result.context;
    if (!pending) {
      throw new Error('OAuth flow completed but no account details were attached to it. Start over with imap_add_oauth_account.');
    }

    const oauth: OAuthConfig = {
      provider: 'microsoft',
      clientId: pending.clientId,
      tenant: pending.tenant,
      refreshToken: result.tokens.refreshToken,
      accessToken: result.tokens.accessToken,
      accessTokenExpiresAt: result.tokens.accessTokenExpiresAt,
      scopes: result.tokens.scopes,
    };

    let account: ImapAccount;
    if (pending.accountId) {
      // Re-authorization / conversion: replace the credentials, keep the rest.
      // The password is cleared — XOAUTH2 never uses it, and Microsoft would
      // not accept it anyway.
      await imapService.disconnect(pending.accountId);
      smtpService.disconnect(pending.accountId);
      account = await accountManager.updateAccount(pending.accountId, {
        name: pending.name,
        host: pending.host,
        port: pending.port,
        user: pending.email,
        email: pending.email,
        password: '',
        tls: true,
        authType: 'oauth2',
        oauth,
        smtp: { host: pending.smtpHost, port: pending.smtpPort, secure: false },
        ...(pending.sentFolder ? { sentFolder: pending.sentFolder } : {}),
        ...(pending.defaultBcc !== undefined ? { defaultBcc: pending.defaultBcc } : {}),
      });
    } else {
      account = await accountManager.addAccount({
        name: pending.name,
        host: pending.host,
        port: pending.port,
        user: pending.email,
        email: pending.email,
        password: '',
        tls: true,
        authType: 'oauth2',
        oauth,
        smtp: { host: pending.smtpHost, port: pending.smtpPort, secure: false },
        ...(pending.sentFolder ? { sentFolder: pending.sentFolder } : {}),
        ...(pending.defaultBcc !== undefined ? { defaultBcc: pending.defaultBcc } : {}),
      });
    }

    const test = await imapService.testConnection(account);

    return {
      content: [{
        type: 'text',
        text: JSON.stringify({
          status: 'complete',
          success: test.success,
          accountId: account.id,
          name: account.name,
          email: pending.email,
          host: account.host,
          port: account.port,
          smtp: { host: pending.smtpHost, port: pending.smtpPort },
          authType: 'oauth2',
          oauth: publicOAuth(oauth),
          connectionTest: test.success
            ? { success: true, folderCount: test.folders?.length ?? 0, messageCount: test.messageCount }
            : {
                success: false,
                error: test.error,
                hint: 'The sign-in succeeded but IMAP rejected the token. Check that the app registration has the delegated "Office 365 Exchange Online" permissions IMAP.AccessAsUser.All and SMTP.Send, and that IMAP is enabled for the mailbox. The account was saved; re-run imap_add_oauth_account with this accountId after fixing it.',
              },
          message: pending.accountId
            ? `Account "${account.name}" re-authorized with OAuth 2.0`
            : `Account "${account.name}" added with OAuth 2.0`,
        }, null, 2)
      }]
    };
  });

  // Remove account tool
  server.registerTool('imap_remove_account', {
    description: 'Remove an IMAP account configuration',
    inputSchema: {
      accountId: z.string().describe('ID of the account to remove'),
    }
  }, async ({ accountId }) => {
    await imapService.disconnect(accountId);
    await accountManager.removeAccount(accountId);
    
    return {
      content: [{
        type: 'text',
        text: JSON.stringify({
          success: true,
          message: `Account ${accountId} removed successfully`,
        }, null, 2)
      }]
    };
  });

  // Connect to account tool
  server.registerTool('imap_connect', {
    description: 'Connect to an IMAP account',
    inputSchema: {
      accountId: z.string().optional().describe('Account ID to connect to'),
      accountName: z.string().optional().describe('Account name to connect to'),
    }
  }, async ({ accountId, accountName }) => {
    let account;
    
    if (accountId) {
      account = accountManager.getAccount(accountId);
    } else if (accountName) {
      account = accountManager.getAccountByName(accountName);
    } else {
      throw new Error('Either accountId or accountName must be provided');
    }
    
    if (!account) {
      throw new Error('Account not found');
    }
    
    await imapService.connect(account);
    
    return {
      content: [{
        type: 'text',
        text: JSON.stringify({
          success: true,
          message: `Connected to account "${account.name}"`,
          accountId: account.id,
        }, null, 2)
      }]
    };
  });

  // Disconnect from account tool
  server.registerTool('imap_disconnect', {
    description: 'Disconnect from an IMAP account',
    inputSchema: {
      accountId: z.string().describe('Account ID to disconnect from'),
    }
  }, async ({ accountId }) => {
    await imapService.disconnect(accountId);

    return {
      content: [{
        type: 'text',
        text: JSON.stringify({
          success: true,
          message: `Disconnected from account ${accountId}`,
        }, null, 2)
      }]
    };
  });

  // Test account connection tool (without re-entering password)
  server.registerTool('imap_test_account', {
    description: 'Test an existing account connection without re-entering credentials. Validates IMAP connectivity and returns folder count and message count.',
    inputSchema: {
      accountId: z.string().describe('Account ID to test'),
    }
  }, async ({ accountId }) => {
    const account = accountManager.getAccount(accountId);
    if (!account) {
      throw new Error(`Account ${accountId} not found`);
    }

    const result = await imapService.testConnection(account);

    return {
      content: [{
        type: 'text',
        text: JSON.stringify({
          accountId,
          accountName: account.name,
          host: account.host,
          ...result,
        }, null, 2)
      }]
    };
  });
}