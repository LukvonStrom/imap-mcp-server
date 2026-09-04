# AGENTS.md

Cross-agent project guide for **imap-mcp-server** — a Model Context Protocol (MCP)
server that gives AI assistants (Claude, ChatGPT, Cursor, …) access to IMAP/SMTP
mailboxes. This file is the shared source of truth for any agent or contributor
working in this repository.

## Architecture

- **Entry point** — `src/index.ts` boots an `McpServer` (MCP SDK) over **stdio**
  and registers all tools via `src/tools/index.ts`.
- **Services** (`src/services/`):
  - `ImapService` — IMAP protocol via **`imapflow`**, with connection pooling,
    folder operations, search, fetch, move/delete, append (Sent/Drafts).
  - `SmtpService` — outbound mail via **`nodemailer`**; composes raw MIME and sends.
  - `AccountManager` — account CRUD with **AES-256-CBC** encrypted credential
    storage at `~/.imap-mcp/accounts.json` (key at `~/.imap-mcp/.key`).
    Credentials can be overridden at read time via environment variables keyed
    by the account's normalized name (uppercase, non-alphanumeric → `_`):
    `IMAP_MCP_ACCOUNT_<NAME>_IMAP_USERNAME` / `_IMAP_PASSWORD` and
    `IMAP_MCP_ACCOUNT_<NAME>_SMTP_USERNAME` / `_SMTP_PASSWORD`. Overrides are
    in-memory only (never persisted) and apply only to existing accounts. The
    variables are consumed at startup (constructor): captured into an
    AES-256-encrypted in-memory cache and deleted from `process.env` so the
    plaintext secret does not linger in the environment. An empty credential is
    the marker for "env-managed"; `assertCredentialsResolved`
    (`src/utils/env-credentials.ts`) is called from `ImapService.connect` and
    `SmtpService.createTransporter` and fails with the missing variable's name
    instead of dialing out blank. Keep `envVarName()` in sync with its copy in
    `public/js/app.js` (the wizard is a static asset and cannot import it) —
    `tests/env-credentials.test.ts` asserts the two agree. OAuth accounts
    (`authType: 'oauth2'`) store `oauth.refreshToken` / `oauth.accessToken`
    encrypted the same way; their env override is
    `IMAP_MCP_ACCOUNT_<NAME>_OAUTH_REFRESH_TOKEN`, and `assertCredentialsResolved`
    requires a refresh token instead of a password for them.
  - `MicrosoftOAuthService` (`src/services/oauth-service.ts`) — Entra
    device-code flow and refresh-token exchange for Outlook.com / Microsoft 365
    (XOAUTH2). Keeps pending flows in memory keyed by a random `flowId` (the
    `device_code` never leaves the process), caches access tokens with a
    five-minute refresh margin, and persists rotated tokens through
    `AccountManager.updateOAuthTokens`. Talks only to
    `https://login.microsoftonline.com`. `ImapService` / `SmtpService` obtain
    tokens from it via `getValidAccessToken` (IMAP retries once with
    `forceRefresh` on an `authenticationFailed` error). Access tokens are
    **resource-bound** (one Microsoft resource per token), so
    `getValidAccessToken(account, scopes)` takes the scope set: the mail
    scopes (default) are cached on the account as before; any other set —
    `MICROSOFT_GRAPH_RULES_SCOPES` for inbox rules — is cached in memory only,
    keyed by (account, scope set). The refresh token is shared across
    resources; a refresh that fails for a non-mail set with
    `invalid_grant` / `AADSTS65001` throws `ConsentRequiredError` (the user
    has not consented to that resource) rather than the "re-authorize the
    mailbox" error. `oauth.grantedScopes` records every scope the user has
    consented to (`scopes` stays the mail set).
  - `OutlookRulesService` (`src/services/outlook-rules-service.ts`) — thin
    Microsoft Graph client (global `fetch`, Bearer token from the OAuth
    service) for Outlook.com / Microsoft 365 **inbox rules**
    (`/me/mailFolders/inbox/messageRules`) and the folder lookups a
    `moveToFolder` needs (well-known names `inbox`/`junkemail`/… map
    directly; other paths walk `/me/mailFolders` → `childFolders` per
    segment, case-insensitive). Retries a 401 once after `forceRefresh`,
    honours one `Retry-After` on 429, maps 401/403 to
    `ConsentRequiredError` and everything else to `GraphApiError`
    (`status`, Graph `code`). Talks only to `https://graph.microsoft.com`.
  - `SpamService` — disposable/known-spam domain detection.
- **Tools** (`src/tools/`), grouped by area:
  - `account-tools.ts` — add / update / list / remove / connect / disconnect / test,
    plus the OAuth pair `imap_add_oauth_account` / `imap_complete_oauth_login`
    (the latter also finishes the `graph-consent` flow started by
    `imap_outlook_authorize_rules`: it widens the existing account's refresh
    token and `grantedScopes` instead of creating an account).
  - `outlook-rules-tools.ts` — `imap_outlook_authorize_rules` (Graph consent,
    device-code) and `imap_outlook_list_rules` / `_create_rule` /
    `_update_rule` / `_delete_rule`. All require `authType: 'oauth2'` +
    `provider: 'microsoft'` and return `{ error: 'graph-consent-required',
    nextStep: 'imap_outlook_authorize_rules' }` while the Graph scopes are
    missing. Create refuses a rule without conditions and `action: 'delete'`
    without `confirmDelete: true`.
  - `email-tools.ts` — search, get, latest, send, reply, forward, save draft,
    mark read/unread, delete, bulk delete, move, attachments, upload, threads.
  - `folder-tools.ts` — list, status, create, unread counts.
  - `spam-tools.ts` — spam analysis, domain stats, allow/deny lists.
  - `export-tools.ts` — `imap_export_messages` (metadata export + rule candidates).
  - `sweep-tools.ts` — `imap_sweep`: age-based move / mark-read / delete by
    sender (dry-run by default; `delete` needs `confirmDelete`; Trash/Junk
    sources refused unless `allowSpecialFolders`). Uses `ImapService.moveUids`
    for chunked `UID MOVE`.
- **Web setup wizard** — `src/web/server.ts` (Express) serves `public/` for
  account onboarding (`npm run setup` / `imap-setup`).
- **Types** — `src/types/index.ts`.
- All tools return **JSON-formatted text** content; errors are returned as
  structured JSON where practical rather than thrown for caller-facing failures.

## Build / Test commands

```bash
npm install          # install dependencies
npm run build        # bundle to dist/ via esbuild (build.mjs)
npm test             # run the vitest suite (run mode)
npm run test:watch   # vitest in watch mode
npm run lint         # tsc --noEmit type-check
npm run dev          # run the server from source (tsx watch)
npm run setup        # launch the web setup wizard
```

Always run `npm run build` **and** `npm test` before committing changes that
touch `src/`. Keep the suite green (currently 577 tests).

> Note: `npm run lint` (`tsc --noEmit`) is memory-hungry on this project — the
> MCP SDK's `registerTool` generics are deep enough to surface a pre-existing
> `TS2589` and can OOM on low-RAM machines. Run it with a larger heap if needed
> (`node --max-old-space-size=8192 ./node_modules/typescript/bin/tsc --noEmit`).

## Security rules (must follow)

1. **Never log secrets.** Passwords, encryption keys, `accounts.json` contents,
   raw auth tokens, and full message bodies must not be written to stdout/stderr
   or to disk outside the user's mailbox/download directories. When adding logs,
   log identifiers (account id, folder, uid), not credentials.
2. **No destructive mail operations without explicit guard logic.** Deletes,
   bulk deletes, and moves must be driven by explicit caller input. Bulk/criteria
   deletion must keep its `dryRun` path and require concrete criteria — never
   delete a whole folder by default, and never widen a delete beyond what the
   caller specified.
3. **Tool schema changes require docs + tests.** Do not rename existing tools or
   change their input/output shape without (a) updating the tool `description`,
   (b) updating `README.md`, and (c) adding/adjusting tests. Prefer additive,
   backward-compatible changes (new optional fields) over breaking ones.
4. **Credentials stay local.** Do not add telemetry, analytics, crash reporting,
   or any third-party network calls. The only outbound connections are to the
   user's own IMAP/SMTP servers — plus, for OAuth accounts only,
   `login.microsoftonline.com` for the token exchange/refresh, and
   `graph.microsoft.com` for the `imap_outlook_*` inbox-rules tools (Graph is
   the only API for Outlook rules; only rule/folder metadata crosses that
   connection, never mail bodies or the mailbox credential). Any new outbound
   host needs the same treatment: documented in README.md and SECURITY.md.
5. **Validate and sanitize file paths** for attachment upload/download (already
   done via `path.basename`); keep writes confined to the configured directories.

## Conventions

- TypeScript, ESM (`"type": "module"`), **Node ≥ 22.12** (declared in
  `package.json` `engines.node`; CI runs 22.x and 24.x).
  - npm checks a dependency's `engines` against the Node doing the *install*,
    not against the floor we declare — so a dependency needing a newer Node
    installs silently and only breaks on a user's older runtime. This is how
    #108 happened. `tests/node-engines.test.ts` walks the runtime dependency
    tree and fails if any package needs more than we advertise. When it fires,
    either raise the floor (CI matrix, AGENTS.md and README with it) or pin the
    package back via npm `overrides`.
- Tool names are stable public API: `imap_*`. Do not rename without a strong
  reason and a migration note.
- Zod schemas describe every tool input; every field gets a `.describe()` that
  tells an LLM **when and how** to use it.
- **Every tool declares a `title` and MCP `annotations`** (`readOnlyHint`,
  `destructiveHint`, `idempotentHint`, `openWorldHint`) — clients use these to
  decide which calls to auto-approve and which to confirm. Pick a preset from
  `src/tools/annotations.ts` (`READ_ONLY`, `CONNECTION`,
  `READ_ONLY_LOCAL_OUTPUT`, `MUTATING`, `LOCAL_CONFIG`, `CREATES`,
  `DESTRUCTIVE`, `SENDS_MAIL`, `NETWORK_AUTH`; each documents its reasoning)
  and spread an override only when one hint genuinely differs. Rules of thumb:
  anything that deletes or sends mail is `destructiveHint: true`; only
  delivery to third parties, the OAuth identity host, and the Graph
  inbox-rules tools are `openWorldHint: true` (the user's own IMAP/SMTP
  server is closed-world; `graph.microsoft.com` is another host, so even
  the read-only `imap_outlook_list_rules` is open-world);
  `readOnlyHint: true` tools must also be listed in `READ_ONLY_TOOLS`
  (`src/tools/index.ts`) — `tests/tool-annotations.test.ts` enforces all of
  this. Tools still return JSON as text content; `outputSchema` /
  `structuredContent` is a future, deliberately separate step because it
  changes the response shape.
- Match the surrounding code style; keep error handling and connection cleanup
  consistent with existing tools.
