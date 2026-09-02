import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { DESTRUCTIVE, MUTATING, READ_ONLY } from './annotations.js';
import { ImapService } from '../services/imap-service.js';
import { AccountManager } from '../services/account-manager.js';
import { isSystemFlag } from '../types/index.js';
import { z } from 'zod';

// Backward-compatible account selector (accountId stays accepted; accountName
// and the single-account default are additive conveniences).
const accountSelector = {
  accountId: z.string().optional().describe('Account ID (from imap_list_accounts). Optional if accountName is given or only one account is configured.'),
  accountName: z.string().optional().describe('Account name instead of accountId. Optional if accountId is given or only one account is configured.'),
};

export function folderTools(
  server: McpServer,
  imapService: ImapService,
  accountManager: AccountManager
): void {
  // List folders tool
  server.registerTool('imap_list_folders', {
    title: 'List folders',
    annotations: READ_ONLY,
    description: 'List all folders/mailboxes for an account (names, hierarchy delimiter, attributes, RFC 6154 special-use role). Use this first to discover exact folder names before searching, moving, or creating subfolders — folder naming varies by provider (e.g. "Archive" vs "[Gmail]/All Mail" vs "INBOX.Archive"). The specialUse field ("\\\\Sent", "\\\\Drafts", "\\\\Trash", "\\\\Junk", "\\\\Archive") identifies a folder\'s role independent of its localized name (e.g. "Gesendet" is the Sent folder when specialUse is "\\\\Sent").',
    inputSchema: {
      ...accountSelector,
    }
  }, async ({ accountId: rawAccountId, accountName }) => {
    const accountId = accountManager.resolveAccountId(rawAccountId, accountName);
    const folders = await imapService.listFolders(accountId);
    
    return {
      content: [{
        type: 'text',
        text: JSON.stringify({
          folders: folders.map(folder => ({
            name: folder.name,
            delimiter: folder.delimiter,
            attributes: folder.attributes,
            // RFC 6154 special-use role (language-independent): lets callers
            // find e.g. the Sent folder even when it is named "Gesendet".
            specialUse: folder.specialUse,
            hasChildren: !!folder.children && folder.children.length > 0,
          })),
        }, null, 2)
      }]
    };
  });

  // Get folder status tool
  server.registerTool('imap_folder_status', {
    title: 'Folder status',
    annotations: READ_ONLY,
    description: 'Get status information about a folder',
    inputSchema: {
      ...accountSelector,
      folder: z.string().describe('Folder name'),
    }
  }, async ({ accountId: rawAccountId, accountName, folder }) => {
    const accountId = accountManager.resolveAccountId(rawAccountId, accountName);

    // Counts come from STATUS, not SELECT: imapflow's mailboxOpen() reports
    // `exists` but neither RECENT nor UNSEEN. STATUS runs first because RFC 3501
    // discourages issuing it against the currently selected mailbox.
    const status = await imapService.getFolderStatus(accountId, folder);
    const box = await imapService.selectFolder(accountId, folder);

    // imapflow returns flag sets as `Set`, which JSON.stringify renders as {}.
    const toFlagArray = (flags: Iterable<string> | undefined): string[] =>
      Array.from(flags ?? []);
    const flags = toFlagArray(box.flags);

    return {
      content: [{
        type: 'text',
        text: JSON.stringify({
          folder: folder,
          messages: {
            total: status.messages,
            new: status.recent,
            unseen: status.unseen,
          },
          uidvalidity: status.uidValidity,
          uidnext: status.uidNext,
          flags,
          permanentFlags: toFlagArray(box.permanentFlags),
          customKeywords: flags.filter(f => !isSystemFlag(f)),
        }, null, 2)
      }]
    };
  });

  // Create folder tool
  server.registerTool('imap_create_folder', {
    title: 'Create folder',
    annotations: MUTATING,
    description:
      'Create a new IMAP folder/mailbox. Most servers also create any missing parent folders ' +
      '(e.g. creating "Archives/2026/2026-05" auto-creates "Archives" and "Archives/2026"). ' +
      'Returns success even if the folder already exists.',
    inputSchema: {
      ...accountSelector,
      folder: z.string().describe('Full folder path to create (e.g. "Archives/2026/2026-05" or "INBOX.Archive")'),
    }
  }, async ({ accountId: rawAccountId, accountName, folder }) => {
    const accountId = accountManager.resolveAccountId(rawAccountId, accountName);
    try {
      const result = await imapService.createFolder(accountId, folder);
      return {
        content: [{
          type: 'text',
          text: JSON.stringify({
            success: true,
            folder: result.path,
            created: result.created,
            alreadyExisted: result.alreadyExisted,
            message: result.alreadyExisted
              ? `Folder "${result.path}" already existed`
              : `Folder "${result.path}" created`,
          }, null, 2)
        }]
      };
    } catch (err) {
      return {
        content: [{
          type: 'text',
          text: JSON.stringify({
            success: false,
            folder,
            error: err instanceof Error ? err.message : 'Unknown error',
          }, null, 2)
        }]
      };
    }
  });

  // Rename folder tool
  server.registerTool('imap_rename_folder', {
    title: 'Rename folder',
    annotations: MUTATING,
    description:
      'Rename or move a folder. Messages, flags and subfolders travel with it, and the server moves the mailbox rather than copying, so this is cheap no matter how much the folder holds. ' +
      'Also moves a folder within the hierarchy when the new path has a different parent (e.g. "Projekt" to "Archiv/Projekt"). ' +
      'Fails if the target name already exists, and refuses INBOX — renaming INBOX on IMAP means moving every message into a new mailbox, which is not what the name suggests.',
    inputSchema: {
      ...accountSelector,
      folder: z.string().describe('Current full folder path (e.g. "Unsortiert2").'),
      newFolder: z.string().describe('New full folder path (e.g. "Unsortiert" or "Archiv/Unsortiert"). Must not exist yet.'),
    }
  }, async ({ accountId: rawAccountId, accountName, folder, newFolder }) => {
    try {
      const accountId = accountManager.resolveAccountId(rawAccountId, accountName);
      const result = await imapService.renameFolder(accountId, folder, newFolder);
      return {
        content: [{
          type: 'text',
          text: JSON.stringify({
            success: true,
            folder: result.path,
            newFolder: result.newPath,
            message: `Folder "${result.path}" renamed to "${result.newPath}"`,
          }, null, 2)
        }]
      };
    } catch (err) {
      return {
        content: [{
          type: 'text',
          text: JSON.stringify({
            success: false,
            folder,
            newFolder,
            error: err instanceof Error ? err.message : 'Unknown error',
          }, null, 2)
        }]
      };
    }
  });

  // Delete folder tool
  server.registerTool('imap_delete_folder', {
    title: 'Delete folder',
    annotations: DESTRUCTIVE,
    description:
      'Delete a folder. Destructive and not undoable: deleting a mailbox deletes the messages in it. ' +
      'Guarded by default — a folder that still holds messages, or that carries a special-use role (Sent, Drafts, Trash, Junk, Archive), is refused and the response says why. Set force to override that; INBOX is refused regardless. ' +
      'To keep the messages, move them elsewhere first (imap_move_email) or rename the folder instead (imap_rename_folder).',
    inputSchema: {
      ...accountSelector,
      folder: z.string().describe('Full path of the folder to delete.'),
      force: z.boolean().default(false).describe('Delete even when the folder still holds messages or is a special-use folder. The messages are deleted with it. Leave false unless the caller explicitly asked to discard the contents.'),
    }
  }, async ({ accountId: rawAccountId, accountName, folder, force }) => {
    try {
      const accountId = accountManager.resolveAccountId(rawAccountId, accountName);
      const result = await imapService.deleteFolder(accountId, folder, { force });
      return {
        content: [{
          type: 'text',
          text: JSON.stringify({
            success: true,
            folder: result.path,
            messagesDeleted: result.messagesDeleted,
            message: result.messagesDeleted > 0
              ? `Folder "${result.path}" deleted along with ${result.messagesDeleted} message(s)`
              : `Empty folder "${result.path}" deleted`,
          }, null, 2)
        }]
      };
    } catch (err) {
      return {
        content: [{
          type: 'text',
          text: JSON.stringify({
            success: false,
            folder,
            error: err instanceof Error ? err.message : 'Unknown error',
          }, null, 2)
        }]
      };
    }
  });

  // Get unread count tool
  server.registerTool('imap_get_unread_count', {
    title: 'Get unread count',
    annotations: READ_ONLY,
    description: 'Count unread (unseen) emails per folder, plus a total. Use for "how many unread do I have?" overviews. Defaults to all folders; pass a folders list to limit scope and speed it up.',
    inputSchema: {
      ...accountSelector,
      folders: z.array(z.string()).optional().describe('List of folders to check (default: all)'),
    }
  }, async ({ accountId: rawAccountId, accountName, folders }) => {
    const accountId = accountManager.resolveAccountId(rawAccountId, accountName);
    const allFolders = await imapService.listFolders(accountId);
    const foldersToCheck = folders || allFolders.map(f => f.name);
    
    const unreadCounts: Record<string, number> = {};
    let totalUnread = 0;
    
    for (const folderName of foldersToCheck) {
      try {
        const unreadMessages = await imapService.searchEmails(accountId, folderName, { seen: false });
        const count = unreadMessages.length;
        unreadCounts[folderName] = count;
        totalUnread += count;
      } catch (error) {
        // Skip folders that can't be accessed
        unreadCounts[folderName] = 0;
      }
    }
    
    return {
      content: [{
        type: 'text',
        text: JSON.stringify({
          totalUnread,
          byFolder: unreadCounts,
        }, null, 2)
      }]
    };
  });
}