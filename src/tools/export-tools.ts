import { z } from 'zod';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { READ_ONLY_LOCAL_OUTPUT } from './annotations.js';
import { createWriteStream, mkdirSync } from 'fs';
import { once } from 'events';
import { basename, join } from 'path';
import { homedir } from 'os';
import { ImapService } from '../services/imap-service.js';
import { AccountManager } from '../services/account-manager.js';
import type { Folder, MessageExportRow } from '../types/index.js';
import { selectSearchFolders, isNonSelectable } from '../utils/search-folders.js';

const accountSelector = {
  accountId: z.string().optional().describe('Account ID (from imap_list_accounts). Either accountId or accountName is required.'),
  accountName: z.string().optional().describe('Account name as an alternative to accountId.'),
};

/** Resolved at call time (not module load) so tests and wrappers can redirect it. */
function downloadDir(): string {
  return process.env.IMAP_DOWNLOAD_DIR || join(homedir(), 'Downloads', 'imap-attachments');
}

function parseDateOnly(value: string | undefined, label: string): Date | undefined {
  if (!value) return undefined;
  const d = /^\d{4}-\d{2}-\d{2}$/.test(value) ? new Date(`${value}T00:00:00Z`) : new Date(value);
  if (Number.isNaN(d.getTime())) {
    throw new Error(`Invalid ${label} date "${value}" — use YYYY-MM-DD.`);
  }
  return d;
}

function flattenFolders(folders: Folder[]): Folder[] {
  const out: Folder[] = [];
  const walk = (list: Folder[]) => {
    for (const f of list) {
      out.push(f);
      if (f.children?.length) walk(f.children);
    }
  };
  walk(folders);
  return out;
}

const CSV_COLUMNS: (keyof MessageExportRow)[] = [
  'folder', 'uid', 'date', 'from', 'fromName', 'fromDomain', 'replyTo', 'to', 'ccCount', 'subject',
  'seen', 'flagged', 'answered', 'size', 'hasAttachments', 'listId', 'hasListUnsubscribe',
  'precedence', 'autoSubmitted', 'messageId', 'inReplyTo',
];

function csvCell(value: unknown): string {
  const text = Array.isArray(value) ? value.join(';') : String(value ?? '');
  return /[",\r\n]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
}

function csvLine(row: MessageExportRow): string {
  return CSV_COLUMNS.map(col => csvCell(row[col])).join(',') + '\n';
}

type Bucket = {
  key: string;
  label?: string;
  count: number;
  unread: number;
  listMail: number;
  folders: Map<string, number>;
};

function bump(map: Map<string, Bucket>, key: string, row: MessageExportRow, label?: string): void {
  if (!key) return;
  let b = map.get(key);
  if (!b) {
    b = { key, label, count: 0, unread: 0, listMail: 0, folders: new Map() };
    map.set(key, b);
  }
  b.count++;
  if (!row.seen) b.unread++;
  if (row.listId || row.hasListUnsubscribe) b.listMail++;
  b.folders.set(row.folder, (b.folders.get(row.folder) || 0) + 1);
  if (label && !b.label) b.label = label;
}

function topOf(map: Map<string, Bucket>, n: number) {
  return [...map.values()]
    .sort((a, b) => b.count - a.count || a.key.localeCompare(b.key))
    .slice(0, n)
    .map(b => ({
      value: b.key,
      ...(b.label ? { name: b.label } : {}),
      messages: b.count,
      unreadPercent: Math.round((b.unread / b.count) * 100),
      listMailPercent: Math.round((b.listMail / b.count) * 100),
      folders: Object.fromEntries([...b.folders.entries()].sort((a, b) => b[1] - a[1])),
    }));
}

/**
 * Turn the aggregate buckets into rule candidates: things a mail client's
 * server-side rules (Outlook.com "Rules", Gmail filters) can key on — a sender
 * domain, a sender address, or a mailing-list id — with the evidence attached so
 * the caller can decide. Purely descriptive; nothing is changed in the mailbox.
 */
function ruleCandidates(
  domains: Map<string, Bucket>,
  senders: Map<string, Bucket>,
  lists: Map<string, Bucket>,
  minMessages: number,
  limit: number,
) {
  type Candidate = {
    match: { type: 'fromDomain' | 'fromAddress' | 'listId'; value: string };
    messages: number;
    unreadPercent: number;
    listMailPercent: number;
    dominantFolder: string;
    dominantFolderPercent: number;
    suggestion: string;
  };
  const out: Candidate[] = [];
  const consider = (type: Candidate['match']['type'], b: Bucket) => {
    if (b.count < minMessages) return;
    const [dominantFolder, inDominant] = [...b.folders.entries()].sort((a, c) => c[1] - a[1])[0];
    const dominantPct = Math.round((inDominant / b.count) * 100);
    const unreadPct = Math.round((b.unread / b.count) * 100);
    const listPct = Math.round((b.listMail / b.count) * 100);
    let suggestion: string;
    if (dominantFolder.toUpperCase() !== 'INBOX' && dominantPct >= 80) {
      suggestion = `Already filed to "${dominantFolder}" ${dominantPct}% of the time — a rule "move to ${dominantFolder}" would automate this.`;
    } else if (listPct >= 60 && unreadPct >= 70) {
      suggestion = 'Mailing-list/newsletter mail that is mostly left unread — candidate for "move to a Newsletters folder" or "mark as read", or unsubscribe.';
    } else if (unreadPct >= 90) {
      suggestion = 'Almost never read — candidate for a move/delete rule, or unsubscribe.';
    } else if (listPct >= 60) {
      suggestion = 'Newsletter that is actually read — a "move to Newsletters" rule keeps the inbox clean without losing it.';
    } else {
      return; // actively read personal/transactional mail: no rule needed
    }
    out.push({
      match: { type, value: b.key },
      messages: b.count,
      unreadPercent: unreadPct,
      listMailPercent: listPct,
      dominantFolder,
      dominantFolderPercent: dominantPct,
      suggestion,
    });
  };
  for (const b of lists.values()) consider('listId', b);
  for (const b of domains.values()) consider('fromDomain', b);
  // Per-address candidates only where the domain is a shared provider (many
  // senders) — otherwise the domain rule already covers it.
  for (const b of senders.values()) {
    const domain = b.key.split('@').pop() || '';
    const domainBucket = domains.get(domain);
    if (domainBucket && domainBucket.count > b.count * 2) consider('fromAddress', b);
  }
  return out.sort((a, b) => b.messages - a.messages).slice(0, limit);
}

export function exportTools(server: McpServer, imapService: ImapService, accountManager: AccountManager): void {
  server.registerTool('imap_export_messages', {
    title: 'Export message metadata',
    annotations: { ...READ_ONLY_LOCAL_OUTPUT, idempotentHint: false }, // each export writes a new timestamped file
    description:
      'Export lightweight per-message metadata (folder, uid, date, sender address/name/domain, recipients, subject, read/flagged/answered state, size, attachment flag, List-Id / List-Unsubscribe, Message-ID) for MANY messages to a local JSONL or CSV file, and return aggregate statistics plus rule candidates. ' +
      'Use this to analyse a mailbox offline — e.g. to decide which Outlook.com / Gmail server-side rules to create (move newsletters, file receipts, delete never-read senders) — without pulling every message through the conversation. Bodies are never exported. ' +
      'Scans all selectable folders by default (Junk included, Trash and Drafts excluded) so mail the provider already filters or that was filed by hand is counted. The file is written under the download directory (IMAP_DOWNLOAD_DIR or ~/Downloads/imap-attachments) in an "exports" subfolder; only the filename can be chosen. ' +
      'Read-only for the mailbox. For a quick look at one folder prefer imap_search_emails; for one sender\'s history use imap_search_emails with from=.',
    inputSchema: {
      ...accountSelector,
      folders: z.array(z.string()).optional().describe('Explicit folder list to export (full paths). Omit to scan every selectable folder subject to the include flags.'),
      includeJunk: z.boolean().default(true).describe('Include Junk/Spam folders when scanning all folders (default true — shows what the provider already filters).'),
      includeTrash: z.boolean().default(false).describe('Include Trash/Deleted Items when scanning all folders (default false — noisy).'),
      since: z.string().optional().describe('Only messages on/after this date (YYYY-MM-DD). Recommended: the last 90–180 days give a representative rule set.'),
      before: z.string().optional().describe('Only messages before this date (YYYY-MM-DD).'),
      limitPerFolder: z.number().int().min(1).max(50000).default(5000).describe('Newest N messages per folder (default 5000).'),
      format: z.enum(['jsonl', 'csv']).default('jsonl').describe('File format. jsonl (one JSON object per line) is easiest for scripts and for Claude to read back in chunks; csv opens in Excel/Numbers.'),
      filename: z.string().optional().describe('File name only (no directories; any path is reduced to its basename). Defaults to <account>-<timestamp>.<format>.'),
      summaryTop: z.number().int().min(1).max(200).default(30).describe('How many top senders / domains / lists to include in the returned summary (default 30). The file always has everything.'),
      minMessagesForRule: z.number().int().min(1).default(5).describe('A sender domain, address or list needs at least this many messages to appear in ruleCandidates (default 5).'),
    }
  }, async ({ accountId: rawAccountId, accountName, folders, includeJunk, includeTrash, since, before, limitPerFolder, format, filename, summaryTop, minMessagesForRule }) => {
    const accountId = accountManager.resolveAccountId(rawAccountId, accountName);
    const account = accountManager.getAccount(accountId);
    const sinceDate = parseDateOnly(since, 'since');
    const beforeDate = parseDateOnly(before, 'before');

    // Resolve the folder set.
    const all = flattenFolders(await imapService.listFolders(accountId));
    let targets: string[];
    if (folders && folders.length > 0) {
      const known = new Set(all.filter(f => !isNonSelectable(f)).map(f => f.name));
      const unknown = folders.filter(f => !known.has(f));
      if (unknown.length > 0) {
        throw new Error(`Unknown or non-selectable folder(s): ${unknown.join(', ')}. Use imap_list_folders to see the exact paths.`);
      }
      targets = folders;
    } else {
      targets = selectSearchFolders(all, { includeTrash, includeSpam: includeJunk, includeDrafts: false });
    }

    // Prepare the output file (confined to the download directory).
    const exportsDir = join(downloadDir(), 'exports');
    mkdirSync(exportsDir, { recursive: true, mode: 0o700 });
    const stamp = new Date().toISOString().replace(/[:.]/g, '-');
    const safeAccount = (account?.name || accountId).replace(/[^A-Za-z0-9._-]+/g, '_');
    const chosen = basename((filename || '').trim()) || `${safeAccount}-${stamp}.${format}`;
    const targetPath = join(exportsDir, chosen);
    const stream = createWriteStream(targetPath, { flags: 'w', mode: 0o600 });
    if (format === 'csv') stream.write(CSV_COLUMNS.join(',') + '\n');

    const senders = new Map<string, Bucket>();
    const domains = new Map<string, Bucket>();
    const lists = new Map<string, Bucket>();
    const perFolder: Record<string, number> = {};
    const errors: string[] = [];
    let total = 0;
    let unread = 0;
    let earliest: string | undefined;
    let latest: string | undefined;

    const onRow = async (row: MessageExportRow) => {
      const line = format === 'csv' ? csvLine(row) : JSON.stringify(row) + '\n';
      if (!stream.write(line)) await once(stream, 'drain');
      total++;
      if (!row.seen) unread++;
      if (!earliest || row.date < earliest) earliest = row.date;
      if (!latest || row.date > latest) latest = row.date;
      bump(senders, row.from, row, row.fromName);
      bump(domains, row.fromDomain, row);
      bump(lists, row.listId, row);
    };

    try {
      for (const folder of targets) {
        // A failed folder must not poison the next: drop the pooled connection
        // so ensureConnected() dials fresh, and give the folder one retry.
        let lastErr: unknown;
        for (let attempt = 0; attempt < 2; attempt++) {
          try {
            perFolder[folder] = await imapService.exportFolderRows(
              accountId,
              folder,
              { since: sinceDate, before: beforeDate, limit: limitPerFolder },
              onRow,
            );
            lastErr = undefined;
            break;
          } catch (err) {
            lastErr = err;
            try { await imapService.disconnect(accountId); } catch { /* already gone */ }
          }
        }
        if (lastErr !== undefined) {
          errors.push(`${folder}: ${lastErr instanceof Error ? lastErr.message : String(lastErr)}`);
        }
      }
    } finally {
      stream.end();
      await once(stream, 'finish');
    }

    return {
      content: [{
        type: 'text',
        text: JSON.stringify({
          success: errors.length === 0,
          path: targetPath,
          format,
          rowCount: total,
          unreadCount: unread,
          dateRange: { earliest, latest },
          folders: perFolder,
          ...(errors.length ? { errors } : {}),
          summary: {
            topSenders: topOf(senders, summaryTop),
            topDomains: topOf(domains, summaryTop),
            topLists: topOf(lists, summaryTop),
          },
          ruleCandidates: ruleCandidates(domains, senders, lists, minMessagesForRule, summaryTop),
          notes: [
            'Bodies were not exported. Each row is one message; "to" is semicolon-separated in CSV.',
            'Outlook.com rules live under Settings → Mail → Rules (conditions: From address/domain, Subject includes, To; actions: move, mark read, delete, categorize). Gmail: Settings → Filters.',
            'Rule candidates are heuristics from read state and folder placement — review before creating a rule.',
          ],
        }, null, 2)
      }]
    };
  });
}
