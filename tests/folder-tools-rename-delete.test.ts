import { describe, it, expect, vi, beforeEach } from 'vitest';
import { folderTools } from '../src/tools/folder-tools.js';
import { ImapService } from '../src/services/imap-service.js';

const handlers = new Map<string, Function>();

const mockServer = {
  registerTool: vi.fn((name: string, _schema: any, handler: Function) => {
    handlers.set(name, handler);
  }),
};

const mockAccountManager = { resolveAccountId: vi.fn((id: string) => id) };
const parse = (result: any) => JSON.parse(result.content[0].text);

describe('imap_rename_folder / imap_delete_folder tools', () => {
  const mockImapService = { renameFolder: vi.fn(), deleteFolder: vi.fn() };

  beforeEach(() => {
    vi.clearAllMocks();
    handlers.clear();
    folderTools(mockServer as any, mockImapService as any, mockAccountManager as any);
  });

  it('registers both tools', () => {
    expect(handlers.has('imap_rename_folder')).toBe(true);
    expect(handlers.has('imap_delete_folder')).toBe(true);
  });

  it('renames and reports both paths', async () => {
    mockImapService.renameFolder.mockResolvedValueOnce({ path: 'Alt', newPath: 'Neu' });

    const parsed = parse(await handlers.get('imap_rename_folder')!({
      accountId: 'acc1', folder: 'Alt', newFolder: 'Neu',
    }));

    expect(mockImapService.renameFolder).toHaveBeenCalledWith('acc1', 'Alt', 'Neu');
    expect(parsed).toMatchObject({ success: true, folder: 'Alt', newFolder: 'Neu' });
  });

  it('surfaces a rename failure as success:false instead of throwing', async () => {
    mockImapService.renameFolder.mockRejectedValueOnce(new Error('a folder with that name already exists'));

    const parsed = parse(await handlers.get('imap_rename_folder')!({
      accountId: 'acc1', folder: 'Alt', newFolder: 'Belegt',
    }));

    expect(parsed.success).toBe(false);
    expect(parsed.error).toContain('already exists');
  });

  it('passes force through to the service', async () => {
    mockImapService.deleteFolder.mockResolvedValueOnce({ path: 'Weg', messagesDeleted: 12 });

    const parsed = parse(await handlers.get('imap_delete_folder')!({
      accountId: 'acc1', folder: 'Weg', force: true,
    }));

    expect(mockImapService.deleteFolder).toHaveBeenCalledWith('acc1', 'Weg', { force: true });
    expect(parsed.message).toContain('12 message(s)');
  });

  it('reports the guard refusal verbatim', async () => {
    mockImapService.deleteFolder.mockRejectedValueOnce(
      new Error('Refusing to delete "Wichtig": it still holds 5 message(s).')
    );

    const parsed = parse(await handlers.get('imap_delete_folder')!({
      accountId: 'acc1', folder: 'Wichtig', force: false,
    }));

    expect(parsed.success).toBe(false);
    expect(parsed.error).toContain('still holds 5 message(s)');
  });
});

describe('ImapService.renameFolder / deleteFolder guards', () => {
  let service: any;
  let client: any;

  beforeEach(() => {
    service = new ImapService();
    client = {
      list: vi.fn(),
      status: vi.fn(),
      mailboxRename: vi.fn(),
      mailboxDelete: vi.fn(),
    };
    vi.spyOn(service as any, 'ensureConnected').mockResolvedValue(client);
  });

  it('refuses to rename INBOX', async () => {
    await expect(service.renameFolder('acc1', 'INBOX', 'Alt')).rejects.toThrow(/Refusing to rename INBOX/);
    expect(client.mailboxRename).not.toHaveBeenCalled();
  });

  it('refuses a rename onto an existing folder', async () => {
    client.list.mockResolvedValue([{ path: 'Alt' }, { path: 'Belegt' }]);
    await expect(service.renameFolder('acc1', 'Alt', 'Belegt')).rejects.toThrow(/already exists/);
    expect(client.mailboxRename).not.toHaveBeenCalled();
  });

  it('refuses a rename of a folder that does not exist', async () => {
    client.list.mockResolvedValue([{ path: 'Andere' }]);
    await expect(service.renameFolder('acc1', 'Fehlt', 'Neu')).rejects.toThrow(/does not exist/);
  });

  it('renames when source exists and target is free', async () => {
    client.list.mockResolvedValue([{ path: 'Alt' }]);
    client.mailboxRename.mockResolvedValue({ path: 'Alt', newPath: 'Neu' });

    await expect(service.renameFolder('acc1', 'Alt', 'Neu')).resolves.toEqual({ path: 'Alt', newPath: 'Neu' });
    expect(client.mailboxRename).toHaveBeenCalledWith('Alt', 'Neu');
  });

  it('refuses to delete INBOX', async () => {
    await expect(service.deleteFolder('acc1', 'INBOX')).rejects.toThrow(/Refusing to delete INBOX/);
    expect(client.mailboxDelete).not.toHaveBeenCalled();
  });

  it('refuses to delete a folder that still holds messages', async () => {
    client.list.mockResolvedValue([{ path: 'Voll' }]);
    client.status.mockResolvedValue({ messages: 5 });

    await expect(service.deleteFolder('acc1', 'Voll')).rejects.toThrow(/still holds 5 message/);
    expect(client.mailboxDelete).not.toHaveBeenCalled();
  });

  it('refuses to delete a special-use folder even when empty', async () => {
    client.list.mockResolvedValue([{ path: 'Gesendet', specialUse: '\\Sent' }]);
    client.status.mockResolvedValue({ messages: 0 });

    await expect(service.deleteFolder('acc1', 'Gesendet')).rejects.toThrow(/it is the \\Sent folder/);
    expect(client.mailboxDelete).not.toHaveBeenCalled();
  });

  it('deletes an empty ordinary folder', async () => {
    client.list.mockResolvedValue([{ path: 'Leer' }]);
    client.status.mockResolvedValue({ messages: 0 });
    client.mailboxDelete.mockResolvedValue(undefined);

    await expect(service.deleteFolder('acc1', 'Leer')).resolves.toEqual({ path: 'Leer', messagesDeleted: 0 });
  });

  it('deletes a non-empty folder with force and reports the loss', async () => {
    client.list.mockResolvedValue([{ path: 'Voll' }]);
    client.status.mockResolvedValue({ messages: 42 });
    client.mailboxDelete.mockResolvedValue(undefined);

    await expect(service.deleteFolder('acc1', 'Voll', { force: true }))
      .resolves.toEqual({ path: 'Voll', messagesDeleted: 42 });
  });

  it('treats an unSTATUSable container as empty', async () => {
    client.list.mockResolvedValue([{ path: 'Container' }]);
    client.status.mockRejectedValue(new Error('Mailbox is not selectable'));
    client.mailboxDelete.mockResolvedValue(undefined);

    await expect(service.deleteFolder('acc1', 'Container')).resolves.toEqual({
      path: 'Container', messagesDeleted: 0,
    });
  });
});
