import type { ToolAnnotations } from '@modelcontextprotocol/sdk/types.js';

/**
 * MCP tool annotation presets.
 *
 * The MCP specification lets a server describe each tool with four boolean
 * hints so a client (Claude Desktop, Claude Code, Cursor, …) can decide which
 * calls to auto-approve and which to confirm with the user:
 *
 * - `readOnlyHint`    — the tool does not modify its environment.
 * - `destructiveHint` — the tool may perform destructive (irreversible) updates.
 * - `idempotentHint`  — repeating the call with the same arguments has no
 *                       additional effect.
 * - `openWorldHint`   — the tool interacts with entities outside the closed
 *                       system of user + mailbox (third parties, other hosts).
 *
 * The hints are *hints*: nothing in the server enforces them. Their purpose is
 * to let `imap_search_emails` and `imap_bulk_delete_by_search` look different
 * to a client that is deciding whether to ask before calling.
 *
 * "Closed world" here means the user's own configured IMAP/SMTP servers and
 * the local `~/.imap-mcp` / download directories. Talking to the user's own
 * mail server is therefore *not* open-world; delivering mail to arbitrary
 * recipients or contacting Microsoft's identity service is.
 *
 * Every `server.registerTool(...)` must pass one of these presets (optionally
 * spread with a per-tool override) — `tests/tool-annotations.test.ts` fails
 * otherwise, and also checks that the `readOnlyHint: true` set matches
 * `READ_ONLY_TOOLS` in `src/tools/index.ts`.
 */

/**
 * Pure reads: search, fetch, list, count, analyze. Nothing in the mailbox,
 * the account store, or the local filesystem changes. Safe to auto-approve.
 */
export const READ_ONLY: ToolAnnotations = {
  readOnlyHint: true,
  destructiveHint: false,
  idempotentHint: true,
  openWorldHint: false,
};

/**
 * Connection lifecycle (`imap_connect` / `imap_disconnect` / `imap_test_account`).
 * These open or close a pooled IMAP session but never touch mailbox contents
 * or stored configuration, so from the user's point of view they are as safe
 * as a read. Opening an already-open (or closing an already-closed) connection
 * is a no-op, hence idempotent. They are part of `READ_ONLY_TOOLS` for the
 * same reason.
 */
export const CONNECTION: ToolAnnotations = { ...READ_ONLY };

/**
 * Reads the mailbox and writes the result to the local download directory
 * (`imap_download_attachment`, `imap_export_messages`). The mailbox itself is
 * untouched and nothing leaves the machine, so the tool is still read-only
 * with respect to the user's mail; the local write is the tool's output, not a
 * side effect on shared state. Writes are confined to the configured download
 * directory (or a caller-chosen `savePath`) and never overwrite mail.
 *
 * Idempotency is per-tool: downloading the same attachment again overwrites
 * the same file (idempotent), while each export creates a new timestamped file
 * (override `idempotentHint: false` at the call site).
 */
export const READ_ONLY_LOCAL_OUTPUT: ToolAnnotations = { ...READ_ONLY };

/**
 * Reversible mailbox or config edits: flag / unflag, mark read / unread,
 * custom keywords, moving messages, creating or renaming folders. Nothing is
 * lost and the change can be undone with the inverse tool. Repeating the call
 * yields the same end state (setting a flag twice sets it once).
 */
export const MUTATING: ToolAnnotations = {
  readOnlyHint: false,
  destructiveHint: false,
  idempotentHint: true,
  openWorldHint: false,
};

/**
 * Edits to local configuration under `~/.imap-mcp` that do not touch any
 * mailbox: updating an account's settings, adding or removing entries on the
 * custom spam / whitelist domain lists. Reversible and idempotent (adding a
 * domain that is already listed changes nothing).
 */
export const LOCAL_CONFIG: ToolAnnotations = { ...MUTATING };

/**
 * Creates a new record on every call and never removes or overwrites one:
 * `imap_add_account` (new account id each time), `imap_save_draft` (a fresh
 * draft is appended per call), `imap_upload_file` (each upload gets a unique
 * prefix in the temporary upload directory). Not destructive — no existing
 * data is affected — but not idempotent either, since repeating the call adds
 * another copy. Nothing reaches a third party: a draft is only stored in the
 * user's own Drafts folder, an upload only lands on local disk.
 */
export const CREATES: ToolAnnotations = {
  readOnlyHint: false,
  destructiveHint: false,
  idempotentHint: false,
  openWorldHint: false,
};

/**
 * Deletes: single / bulk / by-search message deletion, spam and by-domain
 * cleanup, folder deletion, and removing a stored account. These cannot be
 * undone by this server (expunge is final; a Trash move depends on the
 * provider), so clients should confirm them. Marked non-idempotent so a client
 * never treats a retry as harmless — a second `imap_bulk_delete_by_search`
 * with the same criteria may well match newly arrived mail.
 */
export const DESTRUCTIVE: ToolAnnotations = {
  readOnlyHint: false,
  destructiveHint: true,
  idempotentHint: false,
  openWorldHint: false,
};

/**
 * Sends mail via SMTP: `imap_send_email`, `imap_reply_to_email`,
 * `imap_forward_email`. Delivering a message cannot be recalled, so it is
 * flagged destructive even though nothing in the mailbox is removed; and it
 * reaches arbitrary third-party recipients, so it is open-world. Each call
 * sends another copy — not idempotent.
 */
export const SENDS_MAIL: ToolAnnotations = {
  readOnlyHint: false,
  destructiveHint: true,
  idempotentHint: false,
  openWorldHint: true,
};

/**
 * OAuth 2.0 device-code flow (`imap_add_oauth_account`,
 * `imap_complete_oauth_login`). These contact `login.microsoftonline.com`
 * (the only outbound host besides the user's own mail servers), so they are
 * open-world; on completion they store or re-authorize an account, so they
 * are not read-only. Not destructive — nothing is removed — and not
 * idempotent, since each call starts a new flow or advances a pending one.
 */
export const NETWORK_AUTH: ToolAnnotations = {
  readOnlyHint: false,
  destructiveHint: false,
  idempotentHint: false,
  openWorldHint: true,
};
