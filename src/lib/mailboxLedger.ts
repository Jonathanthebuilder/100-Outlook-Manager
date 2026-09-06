export type MailboxStatus = 'available' | 'prepared' | 'used';
export type UnsoldMailboxStatus = Exclude<MailboxStatus, 'used'>;
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
  preparedFor?: string;
  preparationRemark?: string;
  preparedAt?: string;
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
  preparedCount: number;
  usedCount: number;
}

export interface PickFilters {
  firstLetter?: string;
  domain?: string;
  status?: UnsoldMailboxStatus | 'unsold';
}

export interface InventoryStats {
  total: number;
  available: number;
  prepared: number;
  unsold: number;
  used: number;
  outlookCom: number;
  outlookEs: number;
}

const CREDENTIAL_PARTS = 4;
const FORMAT_ERROR = '格式不符合 user----password----client_id----refresh_token';
const DELIVERY_TOKEN_MAX_AGE_MS = 24 * 60 * 60 * 1000;

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
    .filter(isUnsoldMailbox)
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
        preparedCount: 0,
        usedCount: 0,
      };

    if (record.status === 'available') {
      current.availableCount += 1;
    } else if (record.status === 'prepared') {
      current.preparedCount += 1;
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
    if (!isUnsoldMailbox(record) || record.tokenStatus !== 'healthy') {
      return false;
    }

    if (filters.status && filters.status !== 'unsold' && record.status !== filters.status) {
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

export function isUnsoldMailbox(record: MailboxRecord): boolean {
  return record.status === 'available' || record.status === 'prepared';
}

export function isReadyForDelivery(record: MailboxRecord, now: Date = new Date()): boolean {
  if (!isUnsoldMailbox(record) || record.tokenStatus !== 'healthy') {
    return false;
  }

  const tokenCheckedAt = Date.parse(record.tokenCheckedAt || '');
  const tokenAgeMs = now.getTime() - tokenCheckedAt;
  return Number.isFinite(tokenAgeMs) && tokenAgeMs >= 0 && tokenAgeMs <= DELIVERY_TOKEN_MAX_AGE_MS;
}

export function markMailboxPrepared(
  records: MailboxRecord[],
  id: string,
  preparedFor: string,
  preparationRemark: string = '',
  preparedAt: string = new Date().toISOString(),
): MailboxRecord[] {
  const normalizedService = preparedFor.trim();
  const normalizedRemark = preparationRemark.trim();

  return records.map((record) => {
    if (record.id !== id || record.status !== 'available') {
      return record;
    }

    return {
      ...record,
      status: 'prepared',
      preparedFor: normalizedService,
      preparationRemark: normalizedRemark,
      preparedAt,
    };
  });
}

export function markMailboxAvailable(records: MailboxRecord[], id: string): MailboxRecord[] {
  return records.map((record) => {
    if (record.id !== id || record.status !== 'prepared') {
      return record;
    }

    const {
      preparedFor: _preparedFor,
      preparationRemark: _preparationRemark,
      preparedAt: _preparedAt,
      ...rest
    } = record;

    return {
      ...rest,
      status: 'available',
    };
  });
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
        stats.unsold += 1;
      } else if (record.status === 'prepared') {
        stats.prepared += 1;
        stats.unsold += 1;
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
    { total: 0, available: 0, prepared: 0, unsold: 0, used: 0, outlookCom: 0, outlookEs: 0 },
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
