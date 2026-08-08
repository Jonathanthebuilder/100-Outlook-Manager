import { describe, expect, it } from 'vitest';
import {
  exportAvailableMailboxes,
  formatMailboxCredential,
  getInventoryStats,
  getMailboxGroups,
  markMailboxUsed,
  parseMailboxText,
  pickRandomAvailable,
} from './mailboxLedger';

const alphaRaw =
  'AliceExample1001@outlook.com----pass-one----client-a----refresh-token-a';
const betaRaw =
  '邮箱：BrendaExample2002@outlook.es----pass-two----client-b----refresh-token-b';
const usedRaw =
  'CarlosExample3003@outlook.com----pass-three----client-c----refresh-token-c 【6 月 19 日出 Perplexity】';

describe('mailbox ledger parsing', () => {
  it('parses available and used TXT lines while preserving malformed rows', () => {
    const parsed = parseMailboxText(
      [alphaRaw, '', usedRaw, betaRaw, 'this is not a mailbox row'].join('\n'),
    );

    expect(parsed.records).toHaveLength(3);
    expect(parsed.errors).toEqual([
      { lineNumber: 5, line: 'this is not a mailbox row', reason: '格式不符合 user----password----client_id----refresh_token' },
    ]);

    expect(parsed.records[0]).toMatchObject({
      email: 'AliceExample1001@outlook.com',
      password: 'pass-one',
      clientId: 'client-a',
      refreshToken: 'refresh-token-a',
      domain: 'outlook.com',
      firstLetter: 'A',
      status: 'available',
      remark: '',
      sourceLineNumber: 1,
      rawCredential: alphaRaw,
    });
    expect(parsed.records[1]).toMatchObject({
      email: 'CarlosExample3003@outlook.com',
      status: 'used',
      remark: '6 月 19 日出 Perplexity',
      sourceLineNumber: 3,
      rawCredential: 'CarlosExample3003@outlook.com----pass-three----client-c----refresh-token-c',
    });
    expect(parsed.records[2].email).toBe('BrendaExample2002@outlook.es');
  });

  it('exports only available mailboxes in original credential format', () => {
    const parsed = parseMailboxText([alphaRaw, usedRaw, betaRaw].join('\n'));

    expect(exportAvailableMailboxes(parsed.records)).toBe([alphaRaw, betaRaw.replace('邮箱：', '')].join('\n'));
  });

  it('formats one mailbox as a complete deliverable credential line', () => {
    const parsed = parseMailboxText([usedRaw].join('\n'));

    expect(formatMailboxCredential(parsed.records[0])).toBe(
      'CarlosExample3003@outlook.com----pass-three----client-c----refresh-token-c',
    );
  });

  it('auto-detects supplier files where refresh_token appears before client_id', () => {
    const clientId = '9e5f94bc-e8a4-4e73-b8be-63364c29d753';
    const reversed = `SupplierExample@outlook.com----pass-four----opaque-refresh-token----${clientId}`;
    const parsed = parseMailboxText(reversed);

    expect(parsed.errors).toEqual([]);
    expect(parsed.records[0]).toMatchObject({
      email: 'SupplierExample@outlook.com',
      clientId,
      refreshToken: 'opaque-refresh-token',
      rawCredential: `SupplierExample@outlook.com----pass-four----${clientId}----opaque-refresh-token`,
    });
  });
});

describe('mailbox ledger workflow helpers', () => {
  it('groups available mailboxes by first letter and domain', () => {
    const parsed = parseMailboxText(
      [
        alphaRaw,
        betaRaw,
        'AmeliaExample4004@outlook.es----pass-four----client-d----refresh-token-d',
        usedRaw,
      ].join('\n'),
    );

    expect(getMailboxGroups(parsed.records)).toEqual([
      { key: 'A|outlook.com', firstLetter: 'A', domain: 'outlook.com', availableCount: 1, usedCount: 0 },
      { key: 'A|outlook.es', firstLetter: 'A', domain: 'outlook.es', availableCount: 1, usedCount: 0 },
      { key: 'B|outlook.es', firstLetter: 'B', domain: 'outlook.es', availableCount: 1, usedCount: 0 },
      { key: 'C|outlook.com', firstLetter: 'C', domain: 'outlook.com', availableCount: 0, usedCount: 1 },
    ]);
  });

  it('picks a deterministic random available mailbox from a selected group', () => {
    const parsed = parseMailboxText(
      [
        alphaRaw,
        'AmeliaExample4004@outlook.com----pass-four----client-d----refresh-token-d',
        usedRaw,
      ].join('\n'),
    );

    const picked = pickRandomAvailable(
      parsed.records,
      { firstLetter: 'A', domain: 'outlook.com' },
      () => 0.75,
    );

    expect(picked?.email).toBe('AmeliaExample4004@outlook.com');
  });

  it('marks a mailbox as used without mutating the original list', () => {
    const parsed = parseMailboxText([alphaRaw, betaRaw].join('\n'));
    const updated = markMailboxUsed(parsed.records, parsed.records[0].id, '交给 Alex 使用', '2026-06-19T10:00:00.000Z');

    expect(parsed.records[0].status).toBe('available');
    expect(updated[0]).toMatchObject({
      email: 'AliceExample1001@outlook.com',
      status: 'used',
      remark: '交给 Alex 使用',
      usedAt: '2026-06-19T10:00:00.000Z',
    });
    expect(exportAvailableMailboxes(updated)).toBe(betaRaw.replace('邮箱：', ''));
  });

  it('summarizes inventory counts', () => {
    const parsed = parseMailboxText([alphaRaw, betaRaw, usedRaw].join('\n'));

    expect(getInventoryStats(parsed.records)).toEqual({
      total: 3,
      available: 2,
      used: 1,
      outlookCom: 2,
      outlookEs: 1,
    });
  });
});
