# IMAP MCP Server

A powerful Model Context Protocol (MCP) server that provides seamless IMAP email integration with secure account management and connection pooling.

## Features

- 🔐 **Secure Account Management**: Encrypted credential storage with AES-256 encryption
- 🚀 **Connection Pooling**: Efficient IMAP connection management
- 📧 **Comprehensive Email Operations**: Search, read, move, mark, delete, and bulk delete emails
- ✉️ **Email Sending**: Send, reply, and forward emails via SMTP
- 📁 **Folder Management**: List folders, check status, get unread counts
- 🔄 **Multiple Account Support**: Manage multiple IMAP accounts simultaneously
- 🛡️ **Type-Safe**: Built with TypeScript for reliability
- 🌐 **Web-Based Setup Wizard**: Easy account configuration with provider presets
- 📱 **15+ Email Providers**: Pre-configured settings for Gmail, Outlook, Yahoo, and more
- 🔗 **Auto SMTP Configuration**: Automatic SMTP settings based on IMAP provider
- 🪪 **Microsoft OAuth 2.0**: Device-code sign-in for Outlook.com / Hotmail / Live and Microsoft 365 (XOAUTH2) — Microsoft no longer accepts passwords for IMAP/SMTP

## Installation

> **Requires Node.js 22.12 or newer.** Node 18 and 20 have both reached
> end-of-life, and several of this package's dependencies no longer support
> them. Check yours with `node --version`.

### Run via npx (No Installation Required)

Once published to npm, you can run the server directly without cloning or building anything — `npx` downloads the prebuilt package and runs it:

```bash
npx -y imap-mcp-server
```

This is the easiest way to use the server in an MCP client (see [Configuration](#configuration) for ready-to-paste `npx` configs).

### Quick Install (Recommended)

#### macOS/Linux:
```bash
curl -fsSL https://raw.githubusercontent.com/nikolausm/imap-mcp-server/main/install.sh | bash
```

#### Windows (PowerShell as Administrator):
```powershell
iwr -useb https://raw.githubusercontent.com/nikolausm/imap-mcp-server/main/install.ps1 | iex
```

### Manual Installation

1. Clone the repository:
```bash
git clone https://github.com/nikolausm/imap-mcp-server.git
cd imap-mcp-server
```

2. Install dependencies:
```bash
npm install
```

3. Build the project:
```bash
npm run build
```

## Account Setup

Accounts are stored encrypted in `~/.imap-mcp/accounts.json`. This file is **shared by all run modes** — whether you start the server via `npx`, a global install, or a local clone, they all read the same accounts. So you only need to set up your accounts once.

### Setting Up Accounts in npx Mode

If you run the server via `npx` (no clone), you have two ways to add accounts:

**Option A — Run the setup wizard directly via npx (no install needed):**

```bash
npx -p imap-mcp-server imap-setup
```

This launches the same web-based wizard described below and writes to `~/.imap-mcp/accounts.json`, which your `npx`-configured MCP server then picks up automatically.

**Option B — Add accounts straight from your AI client:**

Once the MCP server is configured, just ask your assistant to add an account — it uses the `imap_add_account` tool. For example:

> "Add my IMAP account: host imap.gmail.com, port 993, user me@gmail.com, password …"

No separate setup step required.

### Web-Based Setup Wizard (Recommended)

After installation, run the setup wizard:

```bash
npm run setup
```

Or if installed globally:

```bash
imap-setup
```

Or directly via npx without installing:

```bash
npx -p imap-mcp-server imap-setup
```

This will:
1. Start a local web server
2. Open your browser to the setup wizard
3. Guide you through adding email accounts with pre-configured settings

### Overriding Credentials via Environment Variables

You can override the username and password of an already-configured account at
runtime with environment variables — useful when you inject secrets from a
password manager or CI system instead of storing them in `accounts.json`.

The variables are keyed by the account **name**, uppercased with every
non-alphanumeric character replaced by `_`. For an account named `Work Gmail`
(key `WORK_GMAIL`):

| Variable | Overrides |
| --- | --- |
| `IMAP_MCP_ACCOUNT_WORK_GMAIL_IMAP_USERNAME` | IMAP username (`user`) |
| `IMAP_MCP_ACCOUNT_WORK_GMAIL_IMAP_PASSWORD` | IMAP password |
| `IMAP_MCP_ACCOUNT_WORK_GMAIL_SMTP_USERNAME` | SMTP username (`smtp.user`) |
| `IMAP_MCP_ACCOUNT_WORK_GMAIL_SMTP_PASSWORD` | SMTP password |
| `IMAP_MCP_ACCOUNT_WORK_GMAIL_OAUTH_REFRESH_TOKEN` | OAuth 2.0 refresh token (`oauth.refreshToken`, OAuth accounts only) |

Notes:
- Overrides apply **only to existing accounts**; if no account's normalized name
  matches, the variable is ignored.
- They are applied **in memory only** — nothing is written back to
  `accounts.json`, and the values are used as-is (not re-encrypted).
- Variables are **consumed at startup**: on server start they are captured into
  an AES-256-encrypted in-memory cache and removed from `process.env`, so the
  plaintext secret does not linger in the environment (where it could leak to
  child processes or diagnostics). Set them before launching the server.

The setup wizard integrates with this: each credential field (IMAP password,
IMAP username, SMTP username, SMTP password) has a **"Do not save to config; set
later using an environment variable"** checkbox. When ticked, the value you enter
is still used to test the connection, but it is not written to `accounts.json` —
the wizard shows the exact variable name to export, and the account picks the
credential up from that variable at runtime.
- SMTP variables take effect only when the account already has an SMTP config.
- Each variable takes effect independently; set only the ones you need.

**If the variable is missing**, the account still holds the empty placeholder the
wizard wrote. Rather than dialing out with a blank credential — which providers
answer with a generic authentication failure that looks exactly like a wrong
password — the server refuses the connection and names what to set:

```
Account "Work Gmail" has IMAP credentials marked as environment-managed, but
this variable was not set when the server started:
IMAP_MCP_ACCOUNT_WORK_GMAIL_IMAP_PASSWORD. Set it and restart the server, or
store the credentials on the account via imap_update_account.
```

Because the variables are read once at startup, setting one in an already-running
shell has no effect until the server is restarted.

### Choosing where accounts are stored

By default accounts and the encryption key live in `~/.imap-mcp/`. Set
`IMAP_MCP_CONFIG_DIR` to an absolute directory to use a different store, e.g.
to keep several isolated account sets side by side or to place the store on an
encrypted volume. The directory is created owner-only (`0700`) if missing.

### Outlook.com / Microsoft 365 (OAuth 2.0)

Microsoft has switched off **basic authentication** for IMAP and SMTP: neither
your account password nor an app password is accepted any more, for personal
Outlook.com / Hotmail / Live / MSN mailboxes and for Microsoft 365. Those
accounts must sign in with **OAuth 2.0** (XOAUTH2). This server supports that
with the **device-code flow**: the assistant gives you a short code and a URL,
you sign in once in a browser, and the server stores a refresh token
(encrypted, next to your other credentials) and mints short-lived access
tokens from it whenever it connects.

Microsoft ties the sign-in to an **app registration** in Microsoft Entra ID, so
a one-time setup is needed to get an *Application (client) ID*:

1. Open the [Microsoft Entra admin center](https://entra.microsoft.com) (a
   personal Microsoft account works — it creates a small default tenant) and go
   to **Identity → Applications → App registrations → New registration**.
2. Give it a name (e.g. `imap-mcp-server`).
3. **Supported account types**: choose
   *Personal Microsoft accounts only* for Outlook.com / Hotmail / Live, or
   *Accounts in any organizational directory and personal Microsoft accounts*
   if you also want to use it for Microsoft 365 work accounts. (For a
   single-organization setup, *Accounts in this organizational directory only*
   is fine — you then sign in with your tenant ID as `tenant`, see below.)
4. Leave **Redirect URI** empty and click **Register**.
5. In the new registration open **Authentication** and set
   **Allow public client flows** to **Yes**, then save. (This is what enables
   the device-code flow.)
6. Open **API permissions → Add a permission → APIs my organization uses**,
   search for **Office 365 Exchange Online**, choose **Delegated permissions**,
   and add **`IMAP.AccessAsUser.All`** and **`SMTP.Send`**. (`offline_access`
   is requested automatically at sign-in.)
7. On the **Overview** page copy the **Application (client) ID**.

Prefer the CLI? `scripts/register-entra-app.sh` does all seven steps with the
[Azure CLI](https://learn.microsoft.com/cli/azure/install-azure-cli) (`az login`
with the mailbox's Microsoft account, then one app registration with the two
delegated permissions) and prints the client ID.

Give the client ID to the server either per call (`clientId`) or once via the
environment variable **`IMAP_MCP_MS_CLIENT_ID`** in your MCP client config:

```json
{
  "mcpServers": {
    "imap": {
      "command": "npx",
      "args": ["-y", "@nikolausm/imap-mcp-server"],
      "env": { "IMAP_MCP_MS_CLIENT_ID": "00000000-0000-0000-0000-000000000000" }
    }
  }
}
```

Then add the mailbox from your assistant:

> Add my Outlook account `me@outlook.com` using OAuth.

1. The assistant calls **`imap_add_oauth_account`** and shows you a URL
   (https://microsoft.com/devicelogin) and a code such as `ABCD1234`.
2. Open the URL, enter the code, and sign in to the mailbox. Approve the
   requested permissions.
3. The assistant calls **`imap_complete_oauth_login`**, which waits for the
   sign-in (up to ~25 s per call; it retries while the result is `pending`),
   stores the account with `authType: "oauth2"`, and runs a connection test.

**Tenant.** `imap_add_oauth_account` signs in against the `consumers` tenant
by default, which is right for personal accounts. For Microsoft 365 pass
`tenant` as `organizations`, `common`, or your tenant's GUID / verified domain
(`contoso.onmicrosoft.com`) — a single-tenant app registration *requires* the
GUID or domain.

**Tokens.** The refresh token and the cached access token are stored AES-256
encrypted in `~/.imap-mcp/accounts.json`, exactly like passwords, and are never
returned by any tool. Access tokens are refreshed automatically about five
minutes before they expire; if a refresh is rejected (`invalid_grant` — the
token expired or was revoked), the error tells you to run
`imap_add_oauth_account` again with the account's `accountId`, which replaces
the tokens in place. The refresh token can also be supplied through
`IMAP_MCP_ACCOUNT_<NAME>_OAUTH_REFRESH_TOKEN` (see above). For OAuth accounts
the server additionally contacts `login.microsoftonline.com` for the token
exchange and refresh — that is the only host besides your IMAP/SMTP servers.

The web setup wizard does not run the OAuth flow; when you pick Outlook or
Microsoft 365 there it points you to the tool flow above.

### Supported Email Providers

The setup wizard includes pre-configured settings for:
- Gmail / Google Workspace
- Microsoft Outlook / Hotmail / Live and Microsoft 365 — **OAuth 2.0 only**, via
  `imap_add_oauth_account` (see above)
- Yahoo Mail
- Apple iCloud Mail
- GMX
- WEB.DE
- IONOS (1&1)
- ProtonMail (with Bridge)
- Fastmail
- Zoho Mail
- AOL Mail
- mailbox.org
- Posteo
- Custom IMAP servers

## Configuration

### Claude Code (CLI)

#### Option A — via npx (no clone/build needed)

```bash
claude mcp add imap -- npx -y imap-mcp-server
```

This always runs the latest published version and requires no local build.

#### Option B — from a local clone

If you use [Claude Code](https://docs.anthropic.com/en/docs/claude-code) in the terminal, add the MCP server with a single command:

**Step 1:** Make sure you have built the project first (see [Manual Installation](#manual-installation)).

**Step 2:** Run this command in your terminal:

```bash
claude mcp add imap -- node /absolute/path/to/imap-mcp-server/dist/index.js
```

> **Important:** Replace `/absolute/path/to/imap-mcp-server` with the actual path where you cloned the repository. For example:
> ```bash
> # macOS/Linux example:
> claude mcp add imap -- node /Users/yourname/imap-mcp-server/dist/index.js
>
> # Windows example:
> claude mcp add imap -- node C:\Users\yourname\imap-mcp-server\dist\index.js
> ```

**Step 3:** Verify it was added:

```bash
claude mcp list
```

You should see `imap` in the list of configured MCP servers. That's it — the IMAP tools are now available in your Claude Code sessions.

> **Tip:** If you want to remove the server later, run:
> ```bash
> claude mcp remove imap
> ```

### Claude Desktop (GUI App)

Add the IMAP MCP server to your Claude Desktop configuration file:

**macOS**: `~/Library/Application Support/Claude/claude_desktop_config.json`
**Windows**: `%APPDATA%\Claude\claude_desktop_config.json`

**Option A — via npx (recommended, no clone/build needed):**

```json
{
  "mcpServers": {
    "imap": {
      "command": "npx",
      "args": ["-y", "imap-mcp-server"],
      "env": {}
    }
  }
}
```

**Option B — from a local clone:**

```json
{
  "mcpServers": {
    "imap": {
      "command": "node",
      "args": ["/path/to/imap-mcp-server/dist/index.js"],
      "env": {}
    }
  }
}
```

### Restricting tool access (read-only mode / allowlist)

By default all tools are exposed. You can restrict which tools the agent sees
using two environment variables (set them under the `env` key of your MCP
config). This is useful when you want to give an assistant **read-only** access
to a mailbox, or expose only a hand-picked subset of tools.

| Variable | Effect |
| --- | --- |
| `IMAP_MCP_READ_ONLY` | When truthy (`1`, `true`, `yes`, `on`), only the safe, read-only tools are registered — searching, reading, listing folders, unread counts, spam analysis. No tool that sends mail, deletes/moves messages, changes flags, or edits accounts is exposed. |
| `IMAP_MCP_ENABLED_TOOLS` | Comma-separated allowlist of tool names — only these are registered. Names are case-insensitive and the `imap_` prefix is optional (`search_emails` ≡ `imap_search_emails`). When set, it takes precedence over `IMAP_MCP_READ_ONLY`. |

**Example — read-only access:**

```json
{
  "mcpServers": {
    "imap": {
      "command": "npx",
      "args": ["-y", "imap-mcp-server"],
      "env": { "IMAP_MCP_READ_ONLY": "true" }
    }
  }
}
```

**Example — explicit allowlist:**

```json
{
  "mcpServers": {
    "imap": {
      "command": "npx",
      "args": ["-y", "imap-mcp-server"],
      "env": { "IMAP_MCP_ENABLED_TOOLS": "imap_search_emails,imap_get_email,imap_get_latest_emails" }
    }
  }
}
```

The read-only subset is: `imap_list_accounts`, `imap_connect`, `imap_disconnect`,
`imap_test_account`, `imap_search_emails`, `imap_get_email`,
`imap_get_latest_emails`, `imap_download_attachment`, `imap_find_thread_messages`,
`imap_find_email_by_message_id`, `imap_export_messages`, `imap_list_folders`, `imap_folder_status`,
`imap_get_unread_count`, `imap_check_spam`, `imap_domain_stats`,
`imap_list_spam_domains`.

#### Tool annotations (client-side confirmation)

Every tool also carries the [MCP tool annotations](https://modelcontextprotocol.io/specification/latest/server/tools#tool-annotations)
and a short `title`, so a client such as Claude Desktop or Claude Code can tell
`imap_search_emails` from `imap_bulk_delete_by_search` and decide on its own
which calls to auto-approve and which to confirm with you — even when all tools
are exposed:

| Hint | Meaning for the client | `true` for |
| --- | --- | --- |
| `readOnlyHint` | Nothing in the mailbox, the account store, or local config changes; safe to run without asking. | Exactly the `IMAP_MCP_READ_ONLY` subset above (attachment download and export write only to the local download directory). |
| `destructiveHint` | The effect cannot be undone by this server — confirm before calling. | Every delete (`imap_delete_*`, `imap_bulk_delete*`, `imap_delete_folder`), `imap_remove_account`, sending mail (`imap_send_email`, `imap_reply_to_email`, `imap_forward_email`), and `imap_sweep` (it can delete; dry-run by default). |
| `idempotentHint` | Repeating the call with the same arguments has no additional effect (a retry is harmless). | Reads, flag/keyword/move/folder edits, config edits. `false` for deletes, sends, drafts, uploads, exports, and adding accounts. |
| `openWorldHint` | The tool reaches beyond your own mail server — third-party recipients or another host. | Sending mail and the Microsoft OAuth sign-in tools (`login.microsoftonline.com`). |

Annotations are hints; they do not replace `IMAP_MCP_READ_ONLY` /
`IMAP_MCP_ENABLED_TOOLS`, which remove tools from the list entirely. The
`readOnlyHint: true` set and `IMAP_MCP_READ_ONLY` are kept identical by a test.

## Usage

Once configured, the IMAP MCP server provides the following tools in Claude:

> **Choosing an account.** For the email and folder tools, `accountId` is
> **optional** and backward-compatible. You may instead pass `accountName`, and
> if you only have a **single** account configured you can omit both — that
> account is used by default. With multiple accounts and no selector, the tool
> returns a clear error listing your options (`imap_list_accounts`).

### Account Management

- **imap_add_account**: Add a new IMAP account
  ```
  Parameters:
  - name: Friendly name for the account
  - host: IMAP server hostname
  - port: Server port (default: 993)
  - user: Username
  - password: Password
  - tls: Use TLS/SSL (default: true)
  - allowStartTLS: When tls is false, set to false to also disable the
      opportunistic STARTTLS upgrade imapflow otherwise attempts whenever the
      server advertises it (validating the cert against `host` regardless of
      `tls`). Needed for providers that advertise STARTTLS on a hostname
      covered only by a shared/wildcard cert. Defaults to true.
      **Warning:** with `tls: false` this yields a fully cleartext session
      (password included) — use only on a trusted/local network.
  - tlsRejectUnauthorized: Set to false to accept self-signed/expired/mismatched
      certificates. Defaults to true. **Warning:** disabling validation lets a
      man-in-the-middle read the password and all mail — only for a trusted
      internal server, never for a public provider.
  - sentFolder: Explicit Sent-folder name for sent-mail copies, e.g. "Gesendet"
      (optional — only needed when the server has no \Sent SPECIAL-USE folder
      and auto-detection fails)
  - defaultBcc: Optional BCC address(es) applied automatically to every
      outbound send, reply, forward, and draft for this account. Merged with
      any per-call `bcc` (duplicates removed case-insensitively)
  ```

- **imap_update_account**: Update an existing account (fix SMTP settings, rename, etc.). **Security:** changing `host`, `port`, `user` or any TLS/SMTP connection setting requires passing `password` (and `smtpPassword` where SMTP has its own) in the same call — the assistant cannot read the stored password back, so a prompt-injected "fix your server settings" can never redirect the credential to another server. OAuth accounts may only point at Microsoft hosts and cannot disable TLS validation. Account names must be unique.
  ```
  Parameters:
  - accountId: ID of the account to update
  - name, host, port, user, password, tls, allowStartTLS, tlsRejectUnauthorized, email: IMAP fields (all optional;
      password / smtpPassword are rejected for OAuth 2.0 accounts — re-authorize
      those with imap_add_oauth_account instead)
  - smtpHost, smtpPort, smtpSecure, smtpUser, smtpPassword: SMTP fields (optional)
  - saveToSent: Save sent emails to the Sent folder (optional)
  - sentFolder: Explicit Sent-folder override (optional). Pass an empty string
      to clear the override and re-enable auto-detection
  - defaultBcc: Optional default BCC address(es) (optional). Pass an empty
      string to clear
  ```

- **imap_add_oauth_account**: Start adding an Outlook.com / Hotmail / Live /
  Microsoft 365 mailbox with OAuth 2.0 (device-code flow) — step 1 of 2. See
  [Outlook.com / Microsoft 365 (OAuth 2.0)](#outlookcom--microsoft-365-oauth-20).
  ```
  Parameters:
  - name: Friendly name (optional, defaults to the email address)
  - email: Mailbox address; used as IMAP/SMTP username and From: address
  - provider: "microsoft" (default; the only provider so far)
  - clientId: Entra Application (client) ID (optional when IMAP_MCP_MS_CLIENT_ID is set)
  - tenant: "consumers" (default, personal accounts), "organizations", "common",
      or a tenant GUID / verified domain (Microsoft 365)
  - accountId: Existing account to re-authorize / convert to OAuth (optional)
  - host, port, smtpHost, smtpPort: Endpoint overrides (optional; defaults
      outlook.office365.com:993 and smtp-mail.outlook.com:587)
  - sentFolder, defaultBcc: As for imap_add_account (optional)
  Returns: { flowId, userCode, verificationUri, expiresAt, instructions }
  ```

- **imap_complete_oauth_login**: Step 2 — waits for the user to finish the
  sign-in, stores the account with its encrypted tokens, and tests it.
  ```
  Parameters:
  - flowId: The flowId from imap_add_oauth_account
  - maxWaitSeconds: How long one call may wait (1–25, default 25; optional)
  Returns: { status: "pending", retryAfterSeconds } while the user has not
      finished (call again), or { status: "complete", accountId, ...,
      connectionTest } — never tokens. "expired" / "denied" / "error" mean
      start over with imap_add_oauth_account.
  ```

- **imap_list_accounts**: List all configured accounts. Each entry includes
  `authType` (`"password"` or `"oauth2"`); OAuth accounts also carry
  `oauth: { provider, tenant, clientId, scopes }`. Tokens are never returned.

- **imap_remove_account**: Remove an account
  ```
  Parameters:
  - accountId: ID of the account to remove
  ```

- **imap_connect**: Connect to an account
  ```
  Parameters:
  - accountId OR accountName: Account identifier
  ```

- **imap_disconnect**: Disconnect from an account
  ```
  Parameters:
  - accountId: Account to disconnect
  ```

### Email Operations

- **imap_search_emails**: Search for emails
  ```
  Parameters:
  - accountId: Account ID
  - folder: Folder name (default: INBOX; ignored when searchAllFolders is true)
  - searchAllFolders: Search across ALL folders at once (default: false).
      Skips Trash/Spam/Drafts and non-selectable folders by default. Use when a
      message may have been filed/moved/archived and you don't know its folder.
  - includeTrash, includeSpam, includeDrafts: Opt those noisy folders back into
      a searchAllFolders run (default: false each)
  - from, to, subject, body: Search criteria
  - since, before: Date filters
  - seen, flagged: Status filters
  - keywords: Match messages with ANY of these custom keywords (server-side OR).
      Read a mailbox's available custom keywords from `imap_folder_status`'s
      `customKeywords` field first.
  - unKeywords: Exclude messages with ANY of these custom keywords (result has
      NONE of them). Same keyword source as `keywords`.
  - limit: Max results (default: 50)
  - includeBody: Include parsed message body in the response (default: false).
      Fetches the RFC822 source once and parses it with mailparser, so you get
      uid + body in a single tool call instead of paying the N+1 cost of one
      `imap_get_email` per match. Body is rendered per `bodyFormat` and capped
      at `bodyMaxLength` per field.
  - bodyFormat: How to render the body when `includeBody` is true — `markdown`
      (default, clean Markdown via Turndown), `text`, `html`, or `auto`.
  - bodyMaxLength: Per-field cap when `includeBody` is true (default: 10000).
  ```
  > With `searchAllFolders`, results include a `folder` field per message plus
  > `foldersSearched`, and any folder that failed to open is reported in
  > `foldersErrored` (so a 0-result answer is never silently incomplete).
  >
  > `includeBody` is honored in the single-folder path only. For a
  > cross-folder sweep the lightweight header shape is preserved by design —
  > pulling RFC822 source for every match across many folders would multiply
  > bandwidth and parse cost. Follow up with `imap_get_email` for the specific
  > uids whose bodies you need.
  >
  > On some servers a "flagged"/starred message carries a custom keyword (e.g.
  > an Open-Xchange color label or Apple's `$MailFlagBit*`) instead of, or in
  > addition to, the `\Flagged` system flag — after any flagged search, check
  > each result's `customKeywords` field before concluding a message is or
  > isn't flagged.

- **imap_get_email**: Get full email content
  ```
  Parameters:
  - accountId: Account ID
  - folder: Folder name
  - uid: Email UID
  - maxContentLength: Max characters for text/html body (default: 10000)
  - includeAttachmentText: Include text attachment previews (default: true)
  - maxAttachmentTextChars: Max characters per text attachment (default: 100000)
  ```

- **imap_get_latest_emails**: Get recent emails
  ```
  Parameters:
  - accountId: Account ID
  - folder: Folder name (default: INBOX)
  - count: Number of emails (default: 10)
  - includeBody: Include parsed message body (default: false). Same semantics
      as the `includeBody` option on `imap_search_emails` — one round-trip
      instead of N×`imap_get_email`.
  - bodyFormat: `markdown` (default), `text`, `html`, or `auto`.
  - bodyMaxLength: Per-field cap (default: 10000).
  ```

- **imap_mark_as_read/unread**: Change email read status
  ```
  Parameters:
  - accountId: Account ID
  - folder: Folder name
  - uid: Email UID, OR an array of UIDs to flag in one call. Batch uses a
      single IMAP STORE so the operation is atomic at the server level — all
      UIDs are flagged, or none. Useful when triaging many messages at once.
  ```

- **imap_flag_email/unflag_email**: Star/unstar an email (sets or clears the IMAP \Flagged system flag — shows as a "star" in Gmail and Apple Mail). Some servers/clients (Open-Xchange, Apple Mail) also set a separate custom keyword (e.g. `$cl_N`, `$MailFlagBit*`) when flagging; unflag only clears `\Flagged`, so if a message still shows as flagged, check `customKeywords` via `imap_get_email` and clear it with `imap_remove_keyword`.
  ```
  Parameters:
  - accountId: Account ID
  - folder: Folder name
  - uid: Email UID
  ```

- **imap_add_keyword/remove_keyword**: Set or clear an arbitrary *custom* (non-system) IMAP keyword/label on an email, passed through verbatim (e.g. provider color labels like Open-Xchange's `$cl_1`..`$cl_10` or Apple Mail's `$MailFlagBit0`..`$MailFlagBit2`, or any other custom keyword). Backslash-prefixed system flags (e.g. `\Flagged`, `\Seen`, `\Deleted`) are rejected — use the dedicated flag/read tools for those. Not every server permits custom-keyword changes (see the mailbox's PERMANENTFLAGS); if the server rejects or silently ignores the change, the call fails instead of reporting success.
  ```
  Parameters:
  - accountId: Account ID
  - folder: Folder name
  - uid: Email UID, or an array of UIDs to change many messages in one command
  - allInFolder (imap_remove_keyword only): true to clear the keyword from every
      message in the folder that carries it, without naming UIDs. Mutually
      exclusive with uid; one of the two is required so a forgotten uid can
      never sweep a folder by accident
  - keyword: IMAP keyword to set/remove (e.g. "$cl_3")
  ```

- **imap_delete_email**: Delete an email
  ```
  Parameters:
  - accountId: Account ID
  - folder: Folder name
  - uid: Email UID
  ```

- **imap_move_email**: Move an email from one folder to another
  ```
  Parameters:
  - accountId: Account ID
  - folder: Source folder name (default: INBOX)
  - uid: Email UID, OR an array of UIDs to move in one call. Batch moves are
      attributed per-uid in the response (`results[]` with per-uid `uidMap`
      and any errors). Single-uid calls return the legacy response shape.
  - targetFolder: Destination folder name
  - createDestinationIfMissing: Create the destination folder if it does not exist (default: false)
  ```

- **imap_find_thread_messages**: Find inbox messages that belong to the same conversation threads as messages already sorted into another folder. Uses RFC 3501 HEADER search on In-Reply-To and References — works on any IMAP server.
  ```
  Parameters:
  - accountId: Account ID
  - sourceFolder: Folder containing the already-sorted thread messages
  - searchFolder: Folder to search for related messages (default: INBOX)
  - searchReferences: Also match the References header for multi-level threads (default: true)
  - includeBody: Include parsed message body for each found thread message
      (default: false). Same semantics as the `includeBody` option on
      `imap_search_emails` — one round-trip instead of N×`imap_get_email`.
  - bodyFormat: `markdown` (default), `text`, `html`, or `auto`.
  - bodyMaxLength: Per-field cap (default: 10000).
  ```

- **imap_download_attachment**: Download an email attachment (returns images inline, extracts text from PDFs, or saves to downloads directory)
  ```
  Parameters:
  - accountId: Account ID
  - folder: Folder name (default: INBOX)
  - uid: Email UID
  - filename: Attachment filename or contentId
  - savePath: Optional explicit path. Must lie under the download directory
      (`IMAP_DOWNLOAD_DIR` or `~/Downloads/imap-attachments`) or a directory in
      `IMAP_ATTACHMENT_ROOTS`, and must not already exist; anything else is
      refused so a malicious email can never make the assistant write files
      elsewhere on disk. Usually omit it.
  - extractText: For PDFs, extract and return text content inline (default: true)
  ```

- **imap_bulk_delete**: Delete multiple emails at once with chunking and auto-reconnection
  ```
  Parameters:
  - accountId: Account ID
  - folder: Folder name (default: INBOX)
  - uids: Array of email UIDs to delete
  - chunkSize: Emails to delete per batch (default: 50)
  ```

- **imap_bulk_delete_by_search**: Search for emails matching criteria and delete them all
  ```
  Parameters:
  - accountId: Account ID
  - folder: Folder name (default: INBOX)
  - from, to, subject: Search criteria (optional)
  - before, since: Date filters (optional)
  - chunkSize: Emails to delete per batch (default: 50)
  - dryRun: Preview what would be deleted without deleting (default: false)
  ```

- **imap_sweep**: Age-based inbox hygiene by sender — the server-side equivalent of Outlook's "Sweep". For a list of senders, find messages in a folder older than N days and move them to a folder, mark them read, or delete them, optionally keeping the newest few per sender in place. Outlook.com / Gmail rules cannot express "older than a week"; this tool can, and it works on any IMAP server. **Dry run by default** — the first call returns the plan, nothing changes until `dryRun: false`.
  ```
  Parameters:
  - accountId: Account ID (or accountName)
  - folder: Source folder (default: INBOX). Trash/Junk are refused unless allowSpecialFolders: true
  - senders: Array of sender addresses or domain fragments, e.g. ["news@x.com", "@notifications.github.com"].
      One IMAP FROM search per entry; results are unioned, so a message matched
      by two entries is processed once. At least one is required.
  - olderThanDays: Messages older than this many days qualify (required, >= 0).
      Date-only: uses IMAP BEFORE with today (UTC) minus N days; 0 = before today.
  - keepLatest: Always leave the newest N matching messages PER SENDER in place (default: 0)
  - onlyUnread / onlySeen: Restrict by read state (mutually exclusive, optional)
  - action: "move" (default) | "markRead" | "moveAndMarkRead" | "delete"
  - targetFolder: Destination for move / moveAndMarkRead (required for those; must differ from folder)
  - createFolder: Create targetFolder when missing (default: false — a missing target refuses the run)
  - confirmDelete: Must be true for action "delete" (Trash-aware, same path as imap_bulk_delete)
  - allowSpecialFolders: Allow a Trash/Junk source folder (default: false)
  - dryRun: Report the plan only (default: TRUE). Pass false to apply.
  - chunkSize: UIDs per IMAP command when moving/marking/deleting (default: 200)

  Returns:
  - dryRun, folder, action, targetFolder (+ targetFolderExists / targetFolderCreated), cutoffDate
  - perSender[]: sender, matched, qualifying, keptUids, uids (capped at 200, `truncated` flag), oldest, newest
  - totalMatched, totalPlanned, totalActioned, kept, failed, errors[]
  ```
  **Scheduling:** to emulate a rule such as "keep GitHub notifications in the
  inbox for a week, then file them", run the same call on a schedule — a
  recurring assistant task, a cron job driving the MCP server, or any client
  automation — with `dryRun: false`:
  `imap_sweep({ senders: ["@notifications.github.com"], olderThanDays: 7, keepLatest: 3, targetFolder: "Archive/GitHub", action: "moveAndMarkRead", dryRun: false })`.
  Repeating the call is safe: messages it already filed are no longer in the
  source folder, so a re-run only picks up what has aged since.
  At least one concrete criterion (`from`, `to`, `subject`, `before`, or `since`)
  is required — a call with no criteria is refused, so it can never match and
  delete an entire folder.

- **imap_send_email**: Send a new email
  ```
  Parameters:
  - accountId: Account ID to send from
  - to: Recipient email address(es) — an array, or a single comma-separated string
  - subject: Email subject
  - text: Plain text content (optional)
  - html: HTML content (optional)
  - cc: CC recipients (optional)
  - bcc: BCC recipients (optional)
  - replyTo: Reply-to address (optional)
  - attachments: Array of attachments (optional)
    - filename: Attachment filename
    - content: Base64 encoded content; provide exactly one of `content` or `path`
    - path: Readable local file path to attach; provide exactly one of `path` or `content`.
      URLs are rejected. The path must be under the attachment download/upload
      directory (`imap_upload_file` paths always qualify) or a directory listed
      in `IMAP_ATTACHMENT_ROOTS` (path-delimiter separated), otherwise the send
      is refused.
    - contentType: MIME type (optional; detected from the filename extension when omitted)
    - contentDisposition: "attachment" (default) or "inline" — use "inline" for images shown in the HTML body via cid:
    - cid: Content-ID for inline attachments; must match the `cid:` value used in an `<img src="cid:...">` tag in `html`
  - dryRun: Validate attachments and compose MIME without sending or saving to Sent (optional, default: false)
  ```
  Attachments are validated before SMTP is contacted. Invalid base64, unreadable
  paths, missing filenames, ambiguous sources, and inline attachments without
  `cid` fail fast. Successful sends and dry-runs return `attachmentCount` and
  safe `attachmentDiagnostics`: filename, MIME type, size, source, disposition,
  and cid. Diagnostics omit bytes, raw MIME, and local file paths.

  For large files, upload with `imap_upload_file` first and pass its returned
  local `path`; for inline images, set `contentDisposition: "inline"` and a
  matching `cid`.

  After sending, a copy is saved to the account's Sent folder (unless
  `saveToSent` is disabled on the account). The folder is resolved via the
  account's `sentFolder` override → the server's `\Sent` SPECIAL-USE flag →
  a list of known localized names ("Sent", "Gesendet", "Éléments envoyés", …).
  The response reports the outcome: `savedToSent` (boolean), `sentFolder`
  (the folder used), and — when the save fails — `sentSaveError` explaining
  why, instead of failing silently. The same applies to `imap_reply_to_email`
  and `imap_forward_email`.

  When the account has `defaultBcc` configured, those address(es) are always
  BCC'd on send, reply, forward, and draft (merged with any per-call `bcc`;
  duplicates removed case-insensitively). The Bcc header is kept in the MIME
  stored for drafts and Sent-folder copies so mail clients show it.

- **imap_save_draft**: Save an email as a draft (no send). Takes the same fields as `imap_send_email`, plus `inReplyTo`, `references`, and an optional `folder` override for the Drafts folder.

- **imap_reply_to_email**: Reply to an existing email
  ```
  Parameters:
  - accountId: Account ID
  - folder: Folder containing the original email
  - uid: UID of the email to reply to
  - text: Plain text reply content (optional)
  - html: HTML reply content (optional)
  - replyAll: Reply to all recipients (default: false)
  - bcc: BCC recipients (optional; merged with account defaultBcc)
  - attachments: Array of attachments (optional, same shape as imap_send_email, including contentDisposition/cid for inline images)
  ```

- **imap_forward_email**: Forward an existing email
  ```
  Parameters:
  - accountId: Account ID
  - folder: Folder containing the original email
  - uid: UID of the email to forward
  - to: Forward to email address(es)
  - text: Additional text to include (optional)
  - bcc: BCC recipients (optional; merged with account defaultBcc)
  - includeAttachments: Include original attachments (default: true)
  ```

### Mailbox Export and Rule Analysis

- **imap_export_messages**: Export per-message metadata for a whole mailbox (folder, uid, date, sender address/name/domain, recipients, subject, read/flagged/answered, size, attachment flag, List-Id/List-Unsubscribe, Message-ID — never bodies) to a JSONL or CSV file under `<download dir>/exports/`, and return aggregate statistics plus **rule candidates**. Use it to decide which server-side rules to create in Outlook.com (Settings → Mail → Rules) or Gmail without pulling every message through the conversation. Read-only for the mailbox.
  ```
  Parameters:
  - accountId: Account ID (or accountName)
  - folders: Explicit folder list (optional; default: every selectable folder)
  - includeJunk: Include Junk/Spam when scanning all folders (default: true)
  - includeTrash: Include Trash/Deleted Items (default: false)
  - since, before: Date window, YYYY-MM-DD (optional; 90–180 days is a good sample)
  - limitPerFolder: Newest N per folder (default: 5000)
  - format: "jsonl" (default) or "csv"
  - filename: File name only — always written under the exports directory
  - summaryTop: Top-N senders/domains/lists in the summary (default: 30)
  - minMessagesForRule: Minimum messages for a rule candidate (default: 5)

  Returns:
  - path, rowCount, unreadCount, dateRange, per-folder counts
  - summary: topSenders / topDomains / topLists with unread % and folder spread
  - ruleCandidates: sender domains, addresses and List-Ids with the evidence
      (messages, unread %, dominant folder) and a suggested rule
  ```

### Folder Operations

- **imap_list_folders**: List all folders
  ```
  Parameters:
  - accountId: Account ID
  ```
  Each folder includes its `attributes` (raw IMAP LIST flags) and, when the
  server advertises it, `specialUse` — the RFC 6154 role (`\Sent`, `\Drafts`,
  `\Trash`, `\Junk`, `\Archive`) that identifies a folder independent of its
  localized display name (e.g. "Gesendet" carries `specialUse: "\Sent"`).

- **imap_folder_status**: Get folder information
  ```
  Parameters:
  - accountId: Account ID
  - folder: Folder name

  Returns:
  - messages: { total, new, unseen } — from IMAP STATUS
  - uidvalidity, uidnext
  - flags, permanentFlags: string arrays
  - customKeywords: the mailbox's non-system keywords, usable as the
      `keywords` / `unKeywords` input of imap_search_emails
  ```

- **imap_create_folder**: Create a new IMAP folder/mailbox. Most servers also create any missing parent folders. Returns success even if the folder already exists.
  ```
  Parameters:
  - accountId: Account ID
  - folder: Full folder path to create (e.g. "Archives/2026/2026-05" or "INBOX.Archive")
  ```

- **imap_rename_folder**: Rename or move a folder (messages, flags and subfolders travel with it; the server moves the mailbox rather than copying). Fails if the target exists and refuses INBOX.
  ```
  Parameters:
  - accountId: Account ID
  - folder: Current full folder path (e.g. "Unsorted")
  - newFolder: New full folder path (e.g. "Archive/Unsorted"); must not exist yet
  ```

- **imap_delete_folder**: Delete a folder. **Destructive** — the messages in it are deleted too. Guarded by default: a folder that still holds messages or has a special-use role (Sent, Drafts, Trash, Junk, Archive) is refused and the response says why; INBOX is always refused.
  ```
  Parameters:
  - accountId: Account ID
  - folder: Full path of the folder to delete
  - force: Delete even when non-empty or special-use (default: false)
  ```

- **imap_get_unread_count**: Count unread emails
  ```
  Parameters:
  - accountId: Account ID
  - folders: Specific folders (optional)
  ```

## Security

- Credentials are encrypted using AES-256-CBC encryption — passwords and, for
  OAuth 2.0 accounts, the refresh and access tokens alike
- Encryption keys are stored separately in `~/.imap-mcp/.key`
- Account configurations are stored in `~/.imap-mcp/accounts.json`
- The store directory, `.key`, and `accounts.json` are written owner-only
  (`0700`/`0600`) so other local users cannot read the key or the credentials
- The web setup wizard's HTTP API never returns stored passwords to the browser
- Downloaded attachments are confined to the downloads directory; sender-supplied
  filenames cannot write outside it
- Never commit or share your encryption key or account configurations

## Development

### Running in Development Mode

```bash
npm run dev
```

### Building

```bash
npm run build
```

### Project Structure

```
src/
├── index.ts           # MCP server entry point
├── services/
│   ├── imap-service.ts    # IMAP connection management
│   ├── smtp-service.ts    # SMTP service for sending emails
│   └── account-manager.ts # Account configuration
├── tools/
│   ├── index.ts          # Tool registration
│   ├── account-tools.ts  # Account management tools
│   ├── email-tools.ts    # Email operation tools (including send/reply/forward)
│   ├── folder-tools.ts   # Folder operation tools
│   └── sweep-tools.ts    # imap_sweep: age-based filing by sender (dry-run by default)
└── types/
    └── index.ts          # TypeScript type definitions
```

## Example Usage in Claude

1. **Add an account:**
   "Add my Gmail account with username john@gmail.com"

2. **Check new emails:**
   "Show me the latest 5 emails from my Gmail account"

3. **Search emails:**
   "Search for emails from boss@company.com in the last week"

4. **Send an email:**
   "Send an email to client@example.com with subject 'Project Update'"

5. **Reply to emails:**
   "Reply to the latest email from my boss"

6. **Forward emails:**
   "Forward the email with subject 'Meeting Notes' to team@company.com"

7. **Move an email:**
   "Move the invoice email from INBOX to my Taxes folder"

8. **Manage folders:**
   "List all folders in my email account and show unread counts"

## Troubleshooting

### Connection Issues

- Ensure your IMAP server settings are correct
- Check if your email provider requires app-specific passwords
- Verify that IMAP is enabled in your email account settings
- For sending emails, ensure your account has SMTP access enabled

### Recipients arriving as `["a@x.com","b@y.com"]`

`to`, `cc`, `bcc`, `references` and `uid` accept either a single value or an
array. In JSON Schema that is an `anyOf`, and some MCP clients drop the `anyOf`
before showing the schema to the model — the field then looks untyped or
string-typed, and the client serializes the model's array into a string. The
server used to pass that string straight to nodemailer, which folded the
literal `[` and `]` into the first and last address, so every recipient was
rejected by the receiving mail server (issue #127).

The server now detects a stringified array and restores it, both when
validating tool input and again before composing the message, and logs a
warning to stderr naming the field. Nothing needs to change on your side. If
you want to bypass the client behavior entirely, pass recipients as one
comma-separated string: `"Alice <alice@example.com>, Bob <bob@example.org>"`.

### SMTP Configuration

The server automatically configures SMTP settings based on your IMAP provider. If you need custom SMTP settings, you can specify them when adding an account:

```json
{
  "smtp": {
    "host": "smtp.example.com",
    "port": 587,
    "secure": false
  }
}
```

### Common IMAP Settings

- **Gmail**: 
  - Host: imap.gmail.com
  - Port: 993
  - Requires app-specific password

- **Outlook/Hotmail**:
  - Host: outlook.office365.com
  - Port: 993

- **Yahoo**:
  - Host: imap.mail.yahoo.com
  - Port: 993
  - Requires app-specific password

## License

MIT

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.
