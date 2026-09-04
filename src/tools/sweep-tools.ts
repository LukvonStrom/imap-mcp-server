import { z } from 'zod';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { DESTRUCTIVE } from './annotations.js';
import { ImapService } from '../services/imap-service.js';
import { AccountManager } from '../services/account-manager.js';
import { parseSerializedArray } from '../utils/array-input.js';
import { classifySpecialFolder, flattenFolders } from '../utils/search-folders.js';
import type { EmailMessage, SearchCriteria } from '../types/index.js';

const accountSelector = {
  accountId: z.string().optional().describe('Account ID (from imap_list_accounts). Optional if accountName is given or only one account is configured.'),
  accountName: z.string().optional().describe('Account name instead of accountId. Optional if accountId is given or only one account is configured.'),
};

/** Longest UID list echoed back in a response (per sender and for the union). */
export const SWEEP_UID_LIST_CAP = 200;

export type SweepAction = 'move' | 'markRead' | 'moveAndMarkRead' | 'delete';

/**
 * Cutoff for `olderThanDays`: today's **UTC** date minus N days, at midnight.
 *
 * IMAP `BEFORE` is a date-only criterion (RFC 3501: "internal date earlier
 * than the specified date, disregarding time and timezone"), and imapflow
 * formats the Date it is given via `toISOString()`, i.e. in UTC. Building the
 * cutoff in UTC therefore makes the wire date equal to `cutoffDate` in the
 * response, whatever the local timezone of the machine running the server.
 */
export function sweepCutoff(olderThanDays: number, now: Date = new Date()): Date {
  return new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), now.getUTCDate() - olderThanDays));
}

const toIsoDate = (d: Date): string => d.toISOString().slice(0, 10);
const time = (m: EmailMessage): number => new Date(m.date).getTime();
const cap = (uids: number[]) => ({
  uids: uids.slice(0, SWEEP_UID_LIST_CAP),
  truncated: uids.length > SWEEP_UID_LIST_CAP,
});

interface PerSenderPlan {
  sender: string;
  matched: number;
  qualifying: number;
  keptUids: number[];
  uids: number[];
  truncated: boolean;
  oldest: string | null;
  newest: string | null;
  error?: string;
}

export function sweepTools(server: McpServer, imapService: ImapService, accountManager: AccountManager): void {
  server.registerTool('imap_sweep', {
    title: 'Sweep old mail by sender',
    // Mutating, and with `action: "delete"` destructive (a Trash move or expunge
    // through the same path as imap_bulk_delete), so the DESTRUCTIVE preset is
    // the honest default for a client deciding whether to confirm. Unlike a
    // bulk delete, a repeated sweep is idempotent: the messages it moved or
    // marked are no longer in the source folder / are already read, so the same
    // call again finds nothing left to do — and the dryRun default means a call
    // without explicit `dryRun: false` never mutates at all. Not read-only, so
    // it must stay out of READ_ONLY_TOOLS.
    annotations: { ...DESTRUCTIVE, idempotentHint: true },
    description:
      'Inbox hygiene like Outlook\'s "Sweep": for a list of senders, find messages in `folder` older than N days and file them (move), mark them read, or delete them — optionally keeping the newest few per sender in place. ' +
      'Use it for "keep newsletters/notifications from these senders in the inbox for a week, then move them to a folder", which Outlook.com / Gmail server-side rules cannot express because rules cannot see message age. ' +
      'Age is measured with the IMAP BEFORE criterion, which is date-only: the cutoff is today (UTC) minus olderThanDays at midnight, and a message qualifies when its internal date falls on an earlier day. ' +
      'DEFAULTS TO A DRY RUN: the response is the plan (per sender: how many messages qualify, which UIDs, which ones would be kept) and nothing changes until you call it again with dryRun:false. ' +
      'Only messages returned by the per-sender searches are ever touched. Run it on demand or from a recurring assistant task / cron with dryRun:false to emulate an age-based rule. ' +
      'Refuses to run without senders, with action "delete" unless confirmDelete:true, and against a Trash/Junk source folder unless allowSpecialFolders:true.',
    inputSchema: {
      ...accountSelector,
      folder: z.string().default('INBOX').describe('Source folder to sweep (default: INBOX). Trash/Junk folders are refused unless allowSpecialFolders is true.'),
      senders: z.preprocess(
        value => parseSerializedArray(value, 'senders'),
        z.array(z.string()).min(1),
      ).nonoptional().describe('Sender addresses or domain fragments, e.g. ["newsletter@example.com", "@notifications.github.com", "linkedin.com"]. Each entry is one IMAP FROM search (substring match on the From header, case-insensitive on most servers); the results are unioned and a message matched by two entries is processed once. At least one is required — the tool never sweeps a folder unconditionally.'),
      olderThanDays: z.coerce.number().int().min(0).describe('Only messages older than this many days qualify (date-only: internal date before today-UTC minus N days). 0 means "everything before today". Required.'),
      keepLatest: z.coerce.number().int().min(0).default(0).describe('Always leave the newest N matching messages PER SENDER in the source folder even when they are old enough (default 0). A message kept for one sender is never actioned on behalf of another.'),
      onlyUnread: z.boolean().optional().describe('If true, only unread (unseen) messages qualify. Mutually exclusive with onlySeen.'),
      onlySeen: z.boolean().optional().describe('If true, only already-read (seen) messages qualify — e.g. "file away what I have already looked at". Mutually exclusive with onlyUnread.'),
      action: z.enum(['move', 'markRead', 'moveAndMarkRead', 'delete']).default('move').describe('What to do with qualifying messages: "move" (default) files them into targetFolder; "markRead" only sets \\Seen; "moveAndMarkRead" does both; "delete" removes them via the same Trash-aware path as imap_bulk_delete and additionally requires confirmDelete:true.'),
      targetFolder: z.string().optional().describe('Destination folder for "move" / "moveAndMarkRead" (required for those actions, e.g. "Archive/Newsletters"). Must differ from folder.'),
      createFolder: z.boolean().default(false).describe('If true, create targetFolder when it does not exist (default: false — a missing target refuses the run instead).'),
      confirmDelete: z.boolean().default(false).describe('Must be true for action "delete". Ignored for other actions. Deletion cannot be undone by this server.'),
      allowSpecialFolders: z.boolean().default(false).describe('If true, allow a Trash or Junk/Spam folder as the source folder (default: false — refused).'),
      dryRun: z.boolean().default(true).describe('Default TRUE: only report the plan without changing anything. Pass false to actually move / mark / delete. Always inspect a dry run before the first real run of a new sender list.'),
      chunkSize: z.coerce.number().int().min(1).default(200).describe('UIDs per IMAP command when moving, marking, or deleting (default 200). Lower it if the server rejects large UID sets.'),
    },
  }, async ({
    accountId: rawAccountId,
    accountName,
    folder = 'INBOX',
    senders,
    olderThanDays,
    keepLatest = 0,
    onlyUnread,
    onlySeen,
    action = 'move',
    targetFolder,
    createFolder = false,
    confirmDelete = false,
    allowSpecialFolders = false,
    dryRun = true,
    chunkSize = 200,
  }) => {
    const accountId = accountManager.resolveAccountId(rawAccountId, accountName);
    const sweepAction = action as SweepAction;
    const needsTarget = sweepAction === 'move' || sweepAction === 'moveAndMarkRead';
    const target = needsTarget ? (targetFolder ?? '').trim() : undefined;

    const respond = (body: Record<string, unknown>) => ({
      content: [{ type: 'text' as const, text: JSON.stringify(body, null, 2) }],
    });
    const refuse = (error: string) => respond({
      success: false,
      dryRun,
      folder,
      action: sweepAction,
      ...(target ? { targetFolder: target } : {}),
      totalMatched: 0,
      totalActioned: 0,
      error,
    });

    // --- Guards (all evaluated before any mailbox access) -------------------
    const rawSenders = Array.isArray(senders) ? senders : senders === undefined ? [] : [senders];
    const senderList = [...new Set(rawSenders.map(s => String(s).trim()).filter(Boolean))];
    if (senderList.length === 0) {
      return refuse('Refusing to sweep without senders. Pass at least one sender address or domain fragment — an empty list would match the whole folder.');
    }
    if (!Number.isInteger(olderThanDays) || olderThanDays < 0) {
      return refuse('olderThanDays must be an integer >= 0.');
    }
    if (!Number.isInteger(keepLatest) || keepLatest < 0) {
      return refuse('keepLatest must be an integer >= 0.');
    }
    if (!Number.isInteger(chunkSize) || chunkSize < 1) {
      return refuse('chunkSize must be an integer >= 1.');
    }
    if (onlyUnread && onlySeen) {
      return refuse('onlyUnread and onlySeen are mutually exclusive — a message cannot be both unread and read.');
    }
    if (sweepAction === 'delete' && confirmDelete !== true) {
      return refuse('action "delete" requires confirmDelete:true. Review a dry run first; deletion cannot be undone by this server. Consider action "move" to a folder instead.');
    }
    if (needsTarget && !target) {
      return refuse(`action "${sweepAction}" requires targetFolder (the folder to file qualifying messages into).`);
    }
    if (needsTarget && target === folder) {
      return refuse('targetFolder must differ from the source folder.');
    }

    // Trash/Junk as a *source* is almost always a mistake ("sweep my inbox" with
    // folder mis-set), and a delete from Trash is an expunge. SPECIAL-USE flags
    // from the live folder list take precedence over the name heuristics.
    const folders = flattenFolders(await imapService.listFolders(accountId));
    const special = classifySpecialFolder(folder, folders);
    if (special && !allowSpecialFolders) {
      return refuse(`Refusing to sweep the ${special === 'trash' ? 'Trash' : 'Junk/Spam'} folder "${folder}" as a source. Pass allowSpecialFolders:true if you really mean it.`);
    }
    const targetExists = target ? folders.some(f => f.name === target) : undefined;

    // --- Plan: one FROM search per sender, newest `keepLatest` set aside ------
    const cutoff = sweepCutoff(olderThanDays);
    const cutoffDate = toIsoDate(cutoff);
    const errors: string[] = [];
    const perSender: PerSenderPlan[] = [];
    const keptUids = new Set<number>();
    const candidates = new Set<number>();
    const seenUids = new Set<number>();

    for (const sender of senderList) {
      const criteria: SearchCriteria = { from: sender, before: cutoff };
      if (onlyUnread) criteria.seen = false;
      else if (onlySeen) criteria.seen = true;

      let messages: EmailMessage[];
      try {
        messages = await imapService.searchEmails(accountId, folder, criteria);
      } catch (err) {
        const message = err instanceof Error ? err.message : String(err);
        errors.push(`Search for sender "${sender}" failed: ${message}`);
        perSender.push({ sender, matched: 0, qualifying: 0, keptUids: [], uids: [], truncated: false, oldest: null, newest: null, error: message });
        continue;
      }

      // Newest first; UID breaks ties (a higher UID arrived later).
      const sorted = [...messages].sort((a, b) => (time(b) - time(a)) || (b.uid - a.uid));
      const kept = sorted.slice(0, keepLatest);
      const qualifying = sorted.slice(keepLatest);
      for (const m of sorted) seenUids.add(m.uid);
      for (const m of kept) keptUids.add(m.uid);
      for (const m of qualifying) candidates.add(m.uid);

      const qualifyingUids = qualifying.map(m => m.uid).sort((a, b) => a - b);
      perSender.push({
        sender,
        matched: sorted.length,
        qualifying: qualifying.length,
        keptUids: kept.map(m => m.uid),
        ...cap(qualifyingUids),
        oldest: qualifying.length ? new Date(time(qualifying[qualifying.length - 1])).toISOString() : null,
        newest: qualifying.length ? new Date(time(qualifying[0])).toISOString() : null,
      });
    }

    // Union, minus anything some sender wants kept: "keep the newest N from X"
    // must hold even when the same message also matches a broader entry.
    const actionUids = [...candidates].filter(uid => !keptUids.has(uid)).sort((a, b) => a - b);

    const base = {
      dryRun,
      folder,
      action: sweepAction,
      ...(target ? { targetFolder: target, targetFolderExists: targetExists } : {}),
      cutoffDate,
      olderThanDays,
      keepLatest,
      ...(onlyUnread ? { onlyUnread: true } : {}),
      ...(onlySeen ? { onlySeen: true } : {}),
      perSender,
      totalMatched: seenUids.size,
      totalPlanned: actionUids.length,
      kept: keptUids.size,
      ...cap(actionUids),
    };

    if (dryRun) {
      return respond({
        success: errors.length === 0,
        ...base,
        totalActioned: 0,
        errors,
        message: actionUids.length === 0
          ? 'Dry run — nothing qualifies; nothing was changed.'
          : `Dry run — ${actionUids.length} message(s) would be ${describeAction(sweepAction, target)}. Nothing was changed; re-run with dryRun:false to apply.`,
      });
    }

    if (actionUids.length === 0) {
      return respond({
        success: errors.length === 0,
        ...base,
        totalActioned: 0,
        errors,
        message: 'Nothing qualifies; nothing was changed.',
      });
    }

    // --- Apply ---------------------------------------------------------------
    let targetFolderCreated = false;
    if (needsTarget && target && !targetExists) {
      if (!createFolder) {
        return refuse(`Target folder "${target}" does not exist. Pass createFolder:true to create it, or pick an existing folder from imap_list_folders.`);
      }
      const created = await imapService.createFolder(accountId, target);
      targetFolderCreated = created.created;
    }

    let markedRead = 0;
    let moved = 0;
    let deleted = 0;
    let failed = 0;

    if (sweepAction === 'markRead' || sweepAction === 'moveAndMarkRead') {
      for (let i = 0; i < actionUids.length; i += chunkSize) {
        const chunk = actionUids.slice(i, i + chunkSize);
        const result = await imapService.markAsRead(accountId, folder, chunk);
        markedRead += result.marked.length;
        failed += sweepAction === 'markRead' ? result.failed.length : 0;
        if (result.errors?.length) errors.push(...result.errors);
      }
    }

    if (sweepAction === 'move' || sweepAction === 'moveAndMarkRead') {
      const result = await imapService.moveUids(accountId, folder, actionUids, target!, chunkSize);
      moved = result.moved;
      failed += result.failed;
      errors.push(...result.errors);
    }

    if (sweepAction === 'delete') {
      const result = await imapService.bulkDelete(accountId, folder, actionUids, chunkSize);
      deleted = result.deleted;
      failed += result.failed;
      errors.push(...result.errors);
    }

    const totalActioned = sweepAction === 'markRead' ? markedRead : sweepAction === 'delete' ? deleted : moved;

    return respond({
      success: errors.length === 0,
      ...base,
      ...(targetFolderCreated ? { targetFolderCreated: true } : {}),
      totalActioned,
      ...(sweepAction === 'moveAndMarkRead' ? { markedRead } : {}),
      failed,
      errors,
      message: `${totalActioned}/${actionUids.length} message(s) ${describeAction(sweepAction, target)}${failed ? `, ${failed} failed` : ''}.`,
    });
  });
}

function describeAction(action: SweepAction, target?: string): string {
  switch (action) {
    case 'move': return `moved to "${target}"`;
    case 'markRead': return 'marked as read';
    case 'moveAndMarkRead': return `marked as read and moved to "${target}"`;
    case 'delete': return 'deleted';
  }
}
