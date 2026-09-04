import type { Folder } from '../types/index.js';

// Leaf folder names (lowercased) that identify Trash/Spam/Drafts mailboxes on
// servers that do not advertise RFC 6154 SPECIAL-USE flags. Matched against both
// the full path and the last path segment, case-insensitively.
export const TRASH_FOLDER_NAMES = ['trash', 'bin', 'deleted', 'deleted items', 'deleted messages'];
export const SPAM_FOLDER_NAMES = ['spam', 'junk', 'junk email', 'junk e-mail', 'bulk mail'];
export const DRAFTS_FOLDER_NAMES = ['drafts', 'draft'];
// Always skipped regardless of include-flags (provider quarantine folders).
export const BLOCKED_FOLDER_NAMES = ['blocked'];

export interface FolderFilterOptions {
  includeTrash?: boolean;
  includeSpam?: boolean;
  includeDrafts?: boolean;
}

/** Last path segment of a folder, using its hierarchy delimiter (defaults to '/'). */
function leafName(folder: Folder): string {
  const delimiter = folder.delimiter || '/';
  const parts = folder.name.split(delimiter);
  return parts[parts.length - 1] || folder.name;
}

/** True if the folder cannot hold messages (e.g. the Gmail "[Gmail]" container). */
export function isNonSelectable(folder: Folder): boolean {
  return (folder.attributes || []).some(a => a.toLowerCase() === '\\noselect');
}

/** Match a folder against a SPECIAL-USE flag or a list of known leaf/full names. */
function matchesCategory(folder: Folder, specialUse: string, names: string[]): boolean {
  if (folder.specialUse && folder.specialUse.toLowerCase() === specialUse.toLowerCase()) {
    return true;
  }
  const full = folder.name.toLowerCase();
  const leaf = leafName(folder).toLowerCase();
  return names.includes(full) || names.includes(leaf);
}

/**
 * Pick the folders a cross-folder search should scan: every selectable mailbox
 * except, by default, Trash/Spam/Drafts (noisy) and provider quarantine folders.
 * Each noisy category can be opted back in. Detection prefers RFC 6154
 * SPECIAL-USE flags and falls back to common folder names (case-insensitive,
 * leaf-aware so "[Gmail]/Trash" and "INBOX.Drafts" are matched).
 */
export function selectSearchFolders(folders: Folder[], opts: FolderFilterOptions = {}): string[] {
  const { includeTrash = false, includeSpam = false, includeDrafts = false } = opts;

  return folders
    .filter(folder => !isNonSelectable(folder))
    .filter(folder => {
      const full = folder.name.toLowerCase();
      const leaf = leafName(folder).toLowerCase();

      if (BLOCKED_FOLDER_NAMES.includes(full) || BLOCKED_FOLDER_NAMES.includes(leaf)) return false;
      if (!includeTrash && matchesCategory(folder, '\\Trash', TRASH_FOLDER_NAMES)) return false;
      if (!includeSpam && matchesCategory(folder, '\\Junk', SPAM_FOLDER_NAMES)) return false;
      if (!includeDrafts && matchesCategory(folder, '\\Drafts', DRAFTS_FOLDER_NAMES)) return false;

      return true;
    })
    .map(folder => folder.name);
}

/** Depth-first flatten of a nested folder tree (as returned by `ImapService.listFolders`). */
export function flattenFolders(folders: Folder[]): Folder[] {
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

/**
 * Classify a folder *name* as Trash, Spam, or neither, for tools that must not
 * treat those mailboxes as an ordinary source (`imap_sweep`). When the account's
 * folder list is supplied, its RFC 6154 SPECIAL-USE flag wins; otherwise — and
 * as a fallback for a name that is not in the list — the common leaf/full names
 * above are matched case-insensitively.
 */
export function classifySpecialFolder(name: string, folders: Folder[] = []): 'trash' | 'spam' | null {
  const known = flattenFolders(folders).find(f => f.name === name);
  const folder: Folder = known ?? { name, delimiter: '/', attributes: [] };
  if (matchesCategory(folder, '\\Trash', TRASH_FOLDER_NAMES)) return 'trash';
  if (matchesCategory(folder, '\\Junk', SPAM_FOLDER_NAMES)) return 'spam';
  return null;
}
