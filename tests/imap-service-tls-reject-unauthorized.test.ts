import { describe, it, expect, vi, beforeEach } from 'vitest';
import { ImapService } from '../src/services/imap-service.js';
import type { ImapAccount } from '../src/types/index.js';

// Records the options every ImapFlow is constructed with so we can assert on
// the TLS parameters the service hands to imapflow.
const constructorOptions: any[] = [];

vi.mock('imapflow', () => ({
  ImapFlow: class {
    constructor(options: any) {
      constructorOptions.push(options);
    }
    connect() {
      return Promise.resolve();
    }
    logout() {
      return Promise.resolve();
    }
    on() {}
  },
}));

describe('ImapService per-account TLS certificate validation', () => {
  let service: ImapService;

  const account = (tlsRejectUnauthorized?: boolean): ImapAccount => ({
    id: 'acct',
    name: 'Internal host',
    host: 'imap.internal.example',
    port: 993,
    user: 'user@example.com',
    password: 'secret',
    tls: true,
    ...(tlsRejectUnauthorized !== undefined ? { tlsRejectUnauthorized } : {}),
  });

  beforeEach(() => {
    constructorOptions.length = 0;
    service = new ImapService();
  });

  it('validates certificates by default', async () => {
    await service.connect(account());

    expect(constructorOptions).toHaveLength(1);
    expect(constructorOptions[0].tls).toMatchObject({ host: 'imap.internal.example', rejectUnauthorized: true });
  });

  it('keeps validation on when tlsRejectUnauthorized is explicitly true', async () => {
    await service.connect(account(true));

    expect(constructorOptions[0].tls.rejectUnauthorized).toBe(true);
  });

  it('passes rejectUnauthorized: false only when the account opts out', async () => {
    await service.connect(account(false));

    expect(constructorOptions).toHaveLength(1);
    expect(constructorOptions[0].tls.rejectUnauthorized).toBe(false);
    // The opt-out must not disturb the SNI/cert host we validate against.
    expect(constructorOptions[0].tls.host).toBe('imap.internal.example');
  });

  it('applies the same opt-out to testConnection', async () => {
    await service.testConnection(account(false));

    expect(constructorOptions).toHaveLength(1);
    expect(constructorOptions[0].tls.rejectUnauthorized).toBe(false);
  });
});
