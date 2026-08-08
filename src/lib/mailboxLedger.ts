export type MailboxStatus = 'available' | 'used';
export type TokenStatus = 'unknown' | 'healthy' | 'error';

export interface MailboxRecord {
  id: string;
  email: string;
  password: string;
  clientId: string;
  refreshToken: string;
  domain: string;
  firstLetter: string;
  status: MailboxStatus;
  remark: string;
  usedAt?: string;
  importedAt?: string;
  tokenCheckedAt?: string;
  tokenRefreshedAt?: string;
  tokenStatus?: TokenStatus;
  tokenError?: string;
  sourceLineNumber: number;
  rawCredential: string;
}

export interface ParseError {
  lineNumber: number;
  line: string;
  reason: string;
}

export interface ParsedMailboxFile {
  records: MailboxRecord[];
  errors: ParseError[];
}

export interface MailboxGroup {
  key: string;
  firstLetter: string;
  domain: string;
  availableCount: number;
  usedCount: number;
}

export interface PickFilters {
  firstLetter?: string;
  domain?: string;
}

export interface InventoryStats {
  total: number;
  available: number;
  used: number;
  outlookCom: number;
  outlookEs: number;
}

const CREDENTIAL_PARTS = 4;
const FORMAT_ERROR = '格式不符合 user----password----client_id----refresh_token';

export function parseMailboxText(text: string): ParsedMailboxFile {
  const records: MailboxRecord[] = [];
  const errors: ParseError[] = [];

  text.split(/\r?\n/).forEach((line, index) => {
    const sourceLineNumber = index + 1;
    const trimmed = line.trim();

    if (!trimmed) {
      return;
    }

    const { credentialText, remark } = splitRemark(trimmed);
    const normalizedCredential = credentialText.replace(/^邮箱：\s*/, '').trim();
    const parts = normalizedCredential.split('----');

    if (parts.length !== CREDENTIAL_PARTS || parts.some((part) => part.trim() === '')) {
      errors.push({ lineNumber: sourceLineNumber, line: trimmed, reason: FORMAT_ERROR });
      return;
    }

    const [email, password, thirdField, fourthField] = parts.map((part) => part.trim());
    const fieldsAreReversed = !isClientId(thirdField) && isClientId(fourthField);
    const clientId = fieldsAreReversed ? fourthField : thirdField;
    const refreshToken = fieldsAreReversed ? thirdField : fourthField;
    const parsedEmail = parseEmail(email);

    if (!parsedEmail) {
      errors.push({ lineNumber: sourceLineNumber, line: trimmed, reason: '邮箱地址格式不正确' });
      return;
    }

    records.push({
      id: buildRecordId(email, sourceLineNumber),
      email,
      password,
      clientId,
      refreshToken,
      domain: parsedEmail.domain,
      firstLetter: parsedEmail.firstLetter,
      status: remark ? 'used' : 'available',
      remark,
      tokenStatus: 'unknown',
      sourceLineNumber,
      rawCredential: [email, password, clientId, refreshToken].join('----'),
    });
  });

  return { records, errors };
}

export function exportAvailableMailboxes(records: MailboxRecord[]): string {
  return records
    .filter((record) => record.status === 'available')
    .map(formatMailboxCredential)
    .join('\n');
}

export function formatMailboxCredential(record: MailboxRecord): string {
  return record.rawCredential;
}

export function getMailboxGroups(records: MailboxRecord[]): MailboxGroup[] {
  const groups = new Map<string, MailboxGroup>();

  records.forEach((record) => {
    const key = `${record.firstLetter}|${record.domain}`;
    const current =
      groups.get(key) ??
      {
        key,
        firstLetter: record.firstLetter,
        domain: record.domain,
        availableCount: 0,
        usedCount: 0,
      };

    if (record.status === 'available') {
      current.availableCount += 1;
    } else {
      current.usedCount += 1;
    }

    groups.set(key, current);
  });

  return Array.from(groups.values()).sort((a, b) => {
    const letterSort = a.firstLetter.localeCompare(b.firstLetter);
    return letterSort === 0 ? a.domain.localeCompare(b.domain) : letterSort;
  });
}

export function pickRandomAvailable(
  records: MailboxRecord[],
  filters: PickFilters,
  random: () => number = Math.random,
): MailboxRecord | undefined {
  const candidates = records.filter((record) => {
    if (record.status !== 'available') {
      return false;
    }

    if (filters.firstLetter && record.firstLetter !== filters.firstLetter) {
      return false;
    }

    if (filters.domain && record.domain !== filters.domain) {
      return false;
    }

    return true;
  });

  if (candidates.length === 0) {
    return undefined;
  }

  const index = Math.min(Math.floor(random() * candidates.length), candidates.length - 1);
  return candidates[index];
}

export function markMailboxUsed(
  records: MailboxRecord[],
  id: string,
  remark: string,
  usedAt: string = new Date().toISOString(),
): MailboxRecord[] {
  const normalizedRemark = remark.trim();

  return records.map((record) => {
    if (record.id !== id) {
      return record;
    }

    return {
      ...record,
      status: 'used',
      remark: normalizedRemark,
      usedAt,
    };
  });
}

export function getInventoryStats(records: MailboxRecord[]): InventoryStats {
  return records.reduce<InventoryStats>(
    (stats, record) => {
      stats.total += 1;

      if (record.status === 'available') {
        stats.available += 1;
      } else {
        stats.used += 1;
      }

      if (record.domain === 'outlook.com') {
        stats.outlookCom += 1;
      }

      if (record.domain === 'outlook.es') {
        stats.outlookEs += 1;
      }

      return stats;
    },
    { total: 0, available: 0, used: 0, outlookCom: 0, outlookEs: 0 },
  );
}

function splitRemark(line: string): { credentialText: string; remark: string } {
  const match = line.match(/\s*【([^】]+)】\s*$/);

  if (!match) {
    return { credentialText: line, remark: '' };
  }

  return {
    credentialText: line.slice(0, match.index).trim(),
    remark: match[1].trim(),
  };
}

function parseEmail(email: string): { domain: string; firstLetter: string } | undefined {
  const match = email.match(/^([A-Za-z0-9._%+-]+)@([A-Za-z0-9.-]+\.[A-Za-z]{2,})$/);

  if (!match) {
    return undefined;
  }

  return {
    domain: match[2].toLowerCase(),
    firstLetter: match[1].charAt(0).toUpperCase(),
  };
}

function buildRecordId(email: string, sourceLineNumber: number): string {
  return `${email.toLowerCase()}-${sourceLineNumber}`;
}

function isClientId(value: string): boolean {
  return /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(value);
}
