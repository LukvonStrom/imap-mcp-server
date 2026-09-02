import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { promises as fsp } from 'fs';
import path from 'path';
import os from 'os';

// Real filesystem test: IMAP_MCP_CONFIG_DIR must relocate both accounts.json
// and the .key file so several isolated stores can coexist, and the default
// (~/.imap-mcp) must remain untouched when the override is set.
describe('AccountManager IMAP_MCP_CONFIG_DIR override', () => {
  let tmpHome: string;
  let tmpStore: string;
  let prevHome: string | undefined;
  let prevConfigDir: string | undefined;

  beforeEach(async () => {
    prevHome = process.env.HOME;
    prevConfigDir = process.env.IMAP_MCP_CONFIG_DIR;
    tmpHome = await fsp.mkdtemp(path.join(os.tmpdir(), 'imap-mcp-home-'));
    tmpStore = path.join(await fsp.mkdtemp(path.join(os.tmpdir(), 'imap-mcp-store-')), 'nested', 'store');
    process.env.HOME = tmpHome;
    process.env.IMAP_MCP_CONFIG_DIR = tmpStore;
  });

  afterEach(async () => {
    if (prevHome === undefined) delete process.env.HOME;
    else process.env.HOME = prevHome;
    if (prevConfigDir === undefined) delete process.env.IMAP_MCP_CONFIG_DIR;
    else process.env.IMAP_MCP_CONFIG_DIR = prevConfigDir;
    await fsp.rm(tmpHome, { recursive: true, force: true });
    await fsp.rm(path.dirname(path.dirname(tmpStore)), { recursive: true, force: true });
  });

  it('writes the key and the account store under the override directory', async () => {
    const { AccountManager } = await import('../src/services/account-manager.js');
    const manager = new AccountManager();

    const added = await manager.addAccount({
      name: 'Isolated',
      host: 'imap.test.com',
      port: 993,
      user: 'user@test.com',
      password: 'topsecret',
      tls: true,
    });

    await expect(fsp.stat(path.join(tmpStore, '.key'))).resolves.toBeTruthy();
    const raw = await fsp.readFile(path.join(tmpStore, 'accounts.json'), 'utf8');
    expect(raw).toContain(added.id);
    expect(raw).not.toContain('topsecret');
    // Nothing must leak into the default location.
    await expect(fsp.stat(path.join(tmpHome, '.imap-mcp'))).rejects.toThrow();

    // A second manager pointed at the same override reads the account back.
    const again = new AccountManager();
    expect(again.getAccount(added.id)?.password).toBe('topsecret');
  });
});
