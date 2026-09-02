import { describe, it, expect } from 'vitest';
import type { ToolAnnotations } from '@modelcontextprotocol/sdk/types.js';
import { registerTools, READ_ONLY_TOOLS } from '../src/tools/index.js';

/**
 * Every tool must carry MCP annotations so clients can tell a search from a
 * bulk delete, and the `readOnlyHint: true` set must be exactly the subset
 * exposed by `IMAP_MCP_READ_ONLY` — two sources of truth kept honest here.
 */

interface CapturedTool {
  name: string;
  title?: string;
  annotations?: ToolAnnotations;
}

/** Register every tool against a fake server and capture `(name, config)`. */
function captureTools(): CapturedTool[] {
  const saved = {
    enabled: process.env.IMAP_MCP_ENABLED_TOOLS,
    readOnly: process.env.IMAP_MCP_READ_ONLY,
  };
  delete process.env.IMAP_MCP_ENABLED_TOOLS;
  delete process.env.IMAP_MCP_READ_ONLY;

  const tools: CapturedTool[] = [];
  const fakeServer = {
    registerTool: (name: string, config: { title?: string; annotations?: ToolAnnotations }) => {
      tools.push({ name, title: config?.title, annotations: config?.annotations });
    },
  };
  const stub: any = {};
  try {
    registerTools(fakeServer as any, stub, stub, stub, stub);
  } finally {
    if (saved.enabled === undefined) delete process.env.IMAP_MCP_ENABLED_TOOLS;
    else process.env.IMAP_MCP_ENABLED_TOOLS = saved.enabled;
    if (saved.readOnly === undefined) delete process.env.IMAP_MCP_READ_ONLY;
    else process.env.IMAP_MCP_READ_ONLY = saved.readOnly;
  }
  return tools;
}

const HINTS = ['readOnlyHint', 'destructiveHint', 'idempotentHint', 'openWorldHint'] as const;

/** Name fragments that always denote an irreversible action. */
const DESTRUCTIVE_FRAGMENTS = ['delete', 'remove_account', 'send', 'reply', 'forward'];

describe('MCP tool annotations', () => {
  const tools = captureTools();

  it('registers a non-trivial set of uniquely named tools', () => {
    expect(tools.length).toBeGreaterThan(40);
    expect(new Set(tools.map(t => t.name)).size).toBe(tools.length);
  });

  it('every tool has a human-readable title', () => {
    const missing = tools.filter(t => !t.title || !t.title.trim()).map(t => t.name);
    expect(missing).toEqual([]);
    // Titles are short display names, not tool ids or descriptions.
    for (const t of tools) {
      expect(t.title!.length).toBeLessThanOrEqual(40);
      expect(t.title!.startsWith('imap_')).toBe(false);
    }
  });

  it('every tool declares all four annotation hints as booleans', () => {
    const missing = tools.filter(t => !t.annotations).map(t => t.name);
    expect(missing).toEqual([]);
    for (const t of tools) {
      for (const hint of HINTS) {
        expect(typeof t.annotations![hint], `${t.name}.${hint}`).toBe('boolean');
      }
    }
  });

  it('the readOnlyHint set equals READ_ONLY_TOOLS exactly', () => {
    const hinted = tools
      .filter(t => t.annotations?.readOnlyHint === true)
      .map(t => t.name)
      .sort();
    expect(hinted).toEqual([...READ_ONLY_TOOLS].sort());
  });

  it('deletes, account removal, and sending are all marked destructive', () => {
    const shouldBeDestructive = tools.filter(t =>
      DESTRUCTIVE_FRAGMENTS.some(fragment => t.name.includes(fragment))
    );
    expect(shouldBeDestructive.length).toBeGreaterThan(0);
    const wrong = shouldBeDestructive
      .filter(t => t.annotations?.destructiveHint !== true)
      .map(t => t.name);
    expect(wrong).toEqual([]);
  });

  it('no read-only tool is marked destructive', () => {
    const wrong = tools
      .filter(t => t.annotations?.readOnlyHint === true && t.annotations?.destructiveHint === true)
      .map(t => t.name);
    expect(wrong).toEqual([]);
  });

  it('only mail delivery and OAuth sign-in are open-world', () => {
    const openWorld = tools
      .filter(t => t.annotations?.openWorldHint === true)
      .map(t => t.name)
      .sort();
    expect(openWorld).toEqual([
      'imap_add_oauth_account',
      'imap_complete_oauth_login',
      'imap_forward_email',
      'imap_reply_to_email',
      'imap_send_email',
    ]);
  });
});
