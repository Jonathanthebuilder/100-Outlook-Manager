import { mkdir, readFile, rename, writeFile } from 'node:fs/promises';
import { dirname } from 'node:path';
import { createDeviceCode, fetchMailboxMessages, pollDeviceCode, refreshOAuthToken } from './outlookMailbox.mjs';

const maxBodyBytes = 10 * 1024 * 1024;
const defaultOAuthClientId = process.env.OAUTH_CLIENT_ID || '9e5f94bc-e8a4-4e73-b8be-63364c29d753';
const deliveryTokenMaxAgeMs = 24 * 60 * 60 * 1000;

export function createLedgerStore(filePath) {
  let writeQueue = Promise.resolve();

  return {
    async load() {
      try {
        const content = await readFile(filePath, 'utf8');
        const records = JSON.parse(content);
        return validateRecords(records);
      } catch (error) {
        if (error?.code === 'ENOENT') {
          return [];
        }

        throw error;
      }
    },

    async save(records) {
      const validatedRecords = validateRecords(records);
      await mkdir(dirname(filePath), { recursive: true });

      const tempPath = `${filePath}.${process.pid}.tmp`;
      await writeFile(tempPath, `${JSON.stringify(validatedRecords, null, 2)}\n`, 'utf8');
      await rename(tempPath, filePath);
    },

    async update(updater) {
      const runUpdate = async () => {
        const currentRecords = await this.load();
        const nextRecords = validateRecords(await updater(currentRecords));
        await this.save(nextRecords);
        return nextRecords;
      };

      writeQueue = writeQueue.then(runUpdate, runUpdate);
      return writeQueue;
    },
  };
}

export function validateRecords(records) {
  if (!Array.isArray(records)) {
    throw new Error('Invalid records payload');
  }

  records.forEach((record, index) => {
    const isValid =
      record &&
      typeof record.id === 'string' &&
      typeof record.email === 'string' &&
      typeof record.password === 'string' &&
      typeof record.clientId === 'string' &&
      typeof record.refreshToken === 'string' &&
      typeof record.domain === 'string' &&
      typeof record.firstLetter === 'string' &&
      ['available', 'prepared', 'used'].includes(record.status) &&
      typeof record.remark === 'string' &&
      typeof record.sourceLineNumber === 'number' &&
      typeof record.rawCredential === 'string' &&
      (record.preparedFor === undefined || typeof record.preparedFor === 'string') &&
      (record.preparationRemark === undefined || typeof record.preparationRemark === 'string') &&
      (record.preparedAt === undefined || typeof record.preparedAt === 'string') &&
      (record.status !== 'prepared' || (
        typeof record.preparedFor === 'string' &&
        record.preparedFor.trim().length > 0 &&
        typeof record.preparedAt === 'string' &&
        record.preparedAt.length > 0
      )) &&
      (record.usedAt === undefined || typeof record.usedAt === 'string') &&
      (record.importedAt === undefined || typeof record.importedAt === 'string') &&
      (record.tokenCheckedAt === undefined || typeof record.tokenCheckedAt === 'string') &&
      (record.tokenRefreshedAt === undefined || typeof record.tokenRefreshedAt === 'string') &&
      (record.tokenStatus === undefined || ['unknown', 'healthy', 'error'].includes(record.tokenStatus)) &&
      (record.tokenError === undefined || typeof record.tokenError === 'string');

    if (!isValid) {
      throw new Error(`Invalid mailbox record at index ${index}`);
    }
  });

  return records;
}

export function isUnsoldRecord(record) {
  return record?.status === 'available' || record?.status === 'prepared';
}

export function selectRefreshTargets(records, options = {}) {
  const limit = Math.min(Math.max(Number.parseInt(options.limit, 10) || 25, 1), 100);
  const importedAt = String(options.importedAt || '').trim();
  const clientId = String(options.clientId || '').trim();
  const tokenStatuses = new Set(
    Array.isArray(options.tokenStatuses)
      ? options.tokenStatuses.filter((status) => ['unknown', 'healthy', 'error'].includes(status))
      : [],
  );

  return records
    .filter(isUnsoldRecord)
    .filter((record) => !importedAt || record.importedAt === importedAt)
    .filter((record) => !clientId || record.clientId === clientId)
    .filter((record) => tokenStatuses.size === 0 || tokenStatuses.has(record.tokenStatus || 'unknown'))
    .sort((left, right) => {
      const leftChecked = left.tokenCheckedAt || '';
      const rightChecked = right.tokenCheckedAt || '';
      return leftChecked.localeCompare(rightChecked) || left.email.localeCompare(right.email);
    })
    .slice(0, limit);
}

export function readJsonBody(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    let totalBytes = 0;

    req
      .on('data', (chunk) => {
        totalBytes += chunk.length;

        if (totalBytes > maxBodyBytes) {
          reject(new Error('Request body is too large'));
          return;
        }

        chunks.push(chunk);
      })
      .on('end', () => {
        try {
          const body = Buffer.concat(chunks).toString('utf8') || '{}';
          resolve(JSON.parse(body));
        } catch {
          reject(new Error('Invalid JSON body'));
        }
      })
      .on('error', reject);
  });
}

export function sendJson(res, statusCode, payload) {
  const body = JSON.stringify(payload);
  res.writeHead(statusCode, {
    'cache-control': 'no-store',
    'content-length': Buffer.byteLength(body),
    'content-type': 'application/json; charset=utf-8',
  });
  res.end(body);
}

function updateTokenFields(record, refreshToken, checkedAt) {
  const nextToken = refreshToken || record.refreshToken;
  return {
    ...record,
    refreshToken: nextToken,
    rawCredential: [record.email, record.password, record.clientId, nextToken].join('----'),
    tokenCheckedAt: checkedAt,
    tokenRefreshedAt: nextToken !== record.refreshToken ? checkedAt : record.tokenRefreshedAt,
    tokenStatus: 'healthy',
    tokenError: '',
  };
}

function updateTokenError(record, error, checkedAt) {
  return {
    ...record,
    tokenCheckedAt: checkedAt,
    tokenStatus: 'error',
    tokenError: error instanceof Error ? error.message : 'Token 操作失败',
  };
}

async function findRecord(store, id) {
  const records = await store.load();
  return records.find((record) => record.id === id);
}

export async function handleLedgerRequest(req, res, store, now = () => new Date().toISOString()) {
  const parsedUrl = new URL(req.url, 'http://localhost');

  try {
    if (req.method === 'GET' && parsedUrl.pathname === '/api/records') {
      sendJson(res, 200, { records: await store.load() });
      return true;
    }

    if (req.method === 'PUT' && parsedUrl.pathname === '/api/records') {
      const body = await readJsonBody(req);
      const importedAt = now();
      const records = validateRecords(body.records).map((record) => ({
        ...record,
        importedAt: record.importedAt || importedAt,
        tokenStatus: record.tokenStatus || 'unknown',
      }));
      await store.save(records);
      sendJson(res, 200, { records: await store.load() });
      return true;
    }

    if (req.method === 'POST' && parsedUrl.pathname === '/api/records/import') {
      const body = await readJsonBody(req);
      const incoming = validateRecords(body.records);
      const importedAt = now();
      let created = 0;
      let updated = 0;
      const records = await store.update((current) => {
        const byEmail = new Map(current.map((record) => [record.email.toLowerCase(), record]));
        for (const candidate of incoming) {
          const key = candidate.email.toLowerCase();
          const existing = byEmail.get(key);
          if (!existing) {
            byEmail.set(key, {
              ...candidate,
              importedAt,
              tokenStatus: candidate.tokenStatus || 'unknown',
            });
            created += 1;
            continue;
          }
          const tokenChanged = existing.refreshToken !== candidate.refreshToken || existing.clientId !== candidate.clientId;
          byEmail.set(key, {
            ...existing,
            password: candidate.password,
            clientId: candidate.clientId,
            refreshToken: candidate.refreshToken,
            domain: candidate.domain,
            firstLetter: candidate.firstLetter,
            rawCredential: candidate.rawCredential,
            ...(tokenChanged ? {
              tokenStatus: 'unknown',
              tokenCheckedAt: undefined,
              tokenRefreshedAt: undefined,
              tokenError: '',
            } : {}),
          });
          updated += 1;
        }
        return Array.from(byEmail.values()).sort((left, right) => left.email.localeCompare(right.email));
      });
      sendJson(res, 200, { records, created, updated });
      return true;
    }

    const mailMatch = parsedUrl.pathname.match(/^\/api\/records\/([^/]+)\/mail$/);
    if (req.method === 'POST' && mailMatch) {
      const id = decodeURIComponent(mailMatch[1]);
      const record = await findRecord(store, id);
      if (!record) {
        sendJson(res, 404, { error: 'Mailbox record not found' });
        return true;
      }
      if (!isUnsoldRecord(record)) {
        sendJson(res, 409, { error: '已售账号默认禁止读取邮件' });
        return true;
      }
      const body = await readJsonBody(req);
      const checkedAt = now();
      try {
        const result = await fetchMailboxMessages({
          email: record.email,
          clientId: record.clientId,
          refreshToken: record.refreshToken,
          folder: ['INBOX', 'Junk', 'ALL'].includes(body.folder) ? body.folder : 'ALL',
          keyword: String(body.keyword || '').trim(),
          maxCount: Math.min(Math.max(Number.parseInt(body.maxCount, 10) || 10, 1), 50),
        });
        const records = await store.update((current) => current.map((candidate) => (
          candidate.id === id ? updateTokenFields(candidate, result.refreshToken, checkedAt) : candidate
        )));
        sendJson(res, 200, {
          messages: result.messages,
          records,
          tokenUpdated: result.refreshToken !== record.refreshToken,
        });
      } catch (error) {
        const records = await store.update((current) => current.map((candidate) => (
          candidate.id === id ? updateTokenError(candidate, error, checkedAt) : candidate
        )));
        sendJson(res, 400, {
          error: error instanceof Error ? error.message : '读取邮件失败',
          records,
        });
      }
      return true;
    }

    const refreshMatch = parsedUrl.pathname.match(/^\/api\/records\/([^/]+)\/refresh$/);
    if (req.method === 'POST' && refreshMatch) {
      const id = decodeURIComponent(refreshMatch[1]);
      const record = await findRecord(store, id);
      if (!record) {
        sendJson(res, 404, { error: 'Mailbox record not found' });
        return true;
      }
      if (!isUnsoldRecord(record)) {
        sendJson(res, 409, { error: '已售账号禁止刷新 Token' });
        return true;
      }
      const checkedAt = now();
      try {
        const token = await refreshOAuthToken(record.clientId, record.refreshToken);
        const records = await store.update((current) => current.map((candidate) => (
          candidate.id === id ? updateTokenFields(candidate, token.refreshToken, checkedAt) : candidate
        )));
        sendJson(res, 200, { records, tokenUpdated: token.refreshToken !== record.refreshToken });
      } catch (error) {
        const records = await store.update((current) => current.map((candidate) => (
          candidate.id === id ? updateTokenError(candidate, error, checkedAt) : candidate
        )));
        sendJson(res, 400, {
          error: error instanceof Error ? error.message : 'Token 刷新失败',
          records,
        });
      }
      return true;
    }

    if (req.method === 'POST' && parsedUrl.pathname === '/api/records/refresh-available') {
      const body = await readJsonBody(req);
      const max = Math.min(Math.max(Number.parseInt(body.limit, 10) || 25, 1), 100);
      const records = await store.load();
      const targets = selectRefreshTargets(records, {
        limit: max,
        importedAt: body.importedAt,
        clientId: body.clientId,
        tokenStatuses: body.tokenStatuses,
      });
      const successes = [];
      const failures = [];
      for (const target of targets) {
        const checkedAt = now();
        try {
          const token = await refreshOAuthToken(target.clientId, target.refreshToken);
          await store.update((current) => current.map((candidate) => (
            candidate.id === target.id ? updateTokenFields(candidate, token.refreshToken, checkedAt) : candidate
          )));
          successes.push(target.email);
        } catch (error) {
          await store.update((current) => current.map((candidate) => (
            candidate.id === target.id ? updateTokenError(candidate, error, checkedAt) : candidate
          )));
          failures.push({ email: target.email, error: error instanceof Error ? error.message : '刷新失败' });
        }
      }
      sendJson(res, 200, {
        total: targets.length,
        successCount: successes.length,
        failureCount: failures.length,
        successes,
        failures,
        records: await store.load(),
      });
      return true;
    }

    if (req.method === 'POST' && parsedUrl.pathname === '/api/oauth/device-code') {
      const body = await readJsonBody(req);
      const email = String(body.email || '').trim();
      const clientId = String(body.clientId || defaultOAuthClientId).trim();
      if (!/^\S+@\S+\.\S+$/.test(email)) {
        sendJson(res, 400, { error: '请输入有效 Outlook 邮箱' });
        return true;
      }
      if (!/^[0-9a-f-]{36}$/i.test(clientId)) {
        sendJson(res, 400, { error: 'client_id 格式无效' });
        return true;
      }
      const code = await createDeviceCode(clientId);
      sendJson(res, 200, {
        email,
        clientId,
        deviceCode: code.device_code,
        userCode: code.user_code,
        verificationUri: code.verification_uri,
        verificationUriComplete: code.verification_uri_complete || '',
        expiresIn: code.expires_in,
        interval: code.interval || 5,
      });
      return true;
    }

    if (req.method === 'POST' && parsedUrl.pathname === '/api/oauth/poll') {
      const body = await readJsonBody(req);
      const email = String(body.email || '').trim();
      const password = String(body.password || '');
      const clientId = String(body.clientId || defaultOAuthClientId).trim();
      const deviceCode = String(body.deviceCode || '');
      if (!email || !deviceCode) {
        sendJson(res, 400, { error: '授权信息不完整' });
        return true;
      }
      const token = await pollDeviceCode(clientId, deviceCode);
      if (token.pending) {
        sendJson(res, 200, { pending: true, slowDown: token.slowDown });
        return true;
      }
      const savedAt = now();
      const records = await store.update((current) => {
        const existing = current.find((record) => record.email.toLowerCase() === email.toLowerCase());
        if (existing) {
          return current.map((record) => record.id === existing.id
            ? updateTokenFields({ ...record, password: password || record.password, clientId }, token.refreshToken, savedAt)
            : record);
        }
        const [localPart, domain] = email.split('@');
        const sourceLineNumber = Math.max(0, ...current.map((record) => record.sourceLineNumber)) + 1;
        const record = {
          id: `${email.toLowerCase()}-${sourceLineNumber}`,
          email,
          password,
          clientId,
          refreshToken: token.refreshToken,
          domain: domain.toLowerCase(),
          firstLetter: localPart.charAt(0).toUpperCase(),
          status: 'available',
          remark: '',
          importedAt: savedAt,
          tokenCheckedAt: savedAt,
          tokenRefreshedAt: savedAt,
          tokenStatus: 'healthy',
          tokenError: '',
          sourceLineNumber,
          rawCredential: [email, password, clientId, token.refreshToken].join('----'),
        };
        return [...current, record];
      });
      sendJson(res, 200, { pending: false, records });
      return true;
    }

    const prepareMatch = parsedUrl.pathname.match(/^\/api\/records\/([^/]+)\/prepare$/);
    if (req.method === 'POST' && prepareMatch) {
      const body = await readJsonBody(req);
      const preparedFor = String(body.preparedFor ?? '').trim();
      const preparationRemark = String(body.preparationRemark ?? '').trim();

      if (!preparedFor) {
        sendJson(res, 400, { error: '准备服务不能为空' });
        return true;
      }

      const id = decodeURIComponent(prepareMatch[1]);
      const updatedRecords = await store.update((records) => {
        const record = records.find((candidate) => candidate.id === id);
        if (!record) {
          const error = new Error('Mailbox record not found');
          error.statusCode = 404;
          error.records = records;
          throw error;
        }
        if (record.status !== 'available') {
          const error = new Error(record.status === 'prepared' ? 'Mailbox record is already prepared' : '已售账号不能标记为准备状态');
          error.statusCode = 409;
          error.records = records;
          throw error;
        }
        if (record.tokenStatus !== 'healthy') {
          const error = new Error('Token 尚未验证或已失效，请先刷新或重新授权');
          error.statusCode = 409;
          error.records = records;
          throw error;
        }

        return records.map((candidate) => candidate.id === id
          ? {
              ...candidate,
              status: 'prepared',
              preparedFor,
              preparationRemark,
              preparedAt: now(),
            }
          : candidate);
      });

      sendJson(res, 200, { records: updatedRecords });
      return true;
    }

    const unprepareMatch = parsedUrl.pathname.match(/^\/api\/records\/([^/]+)\/unprepare$/);
    if (req.method === 'POST' && unprepareMatch) {
      const id = decodeURIComponent(unprepareMatch[1]);
      const updatedRecords = await store.update((records) => {
        const record = records.find((candidate) => candidate.id === id);
        if (!record) {
          const error = new Error('Mailbox record not found');
          error.statusCode = 404;
          error.records = records;
          throw error;
        }
        if (record.status !== 'prepared') {
          const error = new Error(record.status === 'used' ? '已售账号不能撤销准备状态' : 'Mailbox record is not prepared');
          error.statusCode = 409;
          error.records = records;
          throw error;
        }

        return records.map((candidate) => {
          if (candidate.id !== id) return candidate;
          const {
            preparedFor: _preparedFor,
            preparationRemark: _preparationRemark,
            preparedAt: _preparedAt,
            ...rest
          } = candidate;
          return { ...rest, status: 'available' };
        });
      });

      sendJson(res, 200, { records: updatedRecords });
      return true;
    }

    const useMatch = parsedUrl.pathname.match(/^\/api\/records\/([^/]+)\/use$/);
    if (req.method === 'POST' && useMatch) {
      const body = await readJsonBody(req);
      const remark = String(body.remark ?? '').trim();

      if (!remark) {
        sendJson(res, 400, { error: 'Remark is required' });
        return true;
      }

      const id = decodeURIComponent(useMatch[1]);
      const usedAt = now();
      const updatedRecords = await store.update((records) => {
        const record = records.find((candidate) => candidate.id === id);

        if (!record) {
          const error = new Error('Mailbox record not found');
          error.statusCode = 404;
          error.records = records;
          throw error;
        }

        if (!isUnsoldRecord(record)) {
          const error = new Error('Mailbox record is already used');
          error.statusCode = 409;
          error.records = records;
          throw error;
        }

        if (record.tokenStatus !== 'healthy') {
          const error = new Error('Token 尚未验证或已失效，请先刷新或重新授权');
          error.statusCode = 409;
          error.records = records;
          throw error;
        }

        const tokenCheckedAt = Date.parse(record.tokenCheckedAt || '');
        const tokenAgeMs = Date.parse(usedAt) - tokenCheckedAt;
        if (!Number.isFinite(tokenAgeMs) || tokenAgeMs < 0 || tokenAgeMs > deliveryTokenMaxAgeMs) {
          const error = new Error('发货前必须完成 24 小时内的 Token 健康检查');
          error.statusCode = 409;
          error.records = records;
          throw error;
        }

        return records.map((candidate) =>
          candidate.id === id
            ? {
                ...candidate,
                status: 'used',
                remark,
                usedAt,
              }
            : candidate,
        );
      });

      sendJson(res, 200, { records: updatedRecords });
      return true;
    }

    const rollbackSaleMatch = parsedUrl.pathname.match(/^\/api\/records\/([^/]+)\/rollback-sale$/);
    if (req.method === 'POST' && rollbackSaleMatch) {
      const id = decodeURIComponent(rollbackSaleMatch[1]);
      let restoredStatus = 'available';
      const updatedRecords = await store.update((records) => {
        const record = records.find((candidate) => candidate.id === id);

        if (!record) {
          const error = new Error('Mailbox record not found');
          error.statusCode = 404;
          error.records = records;
          throw error;
        }

        if (record.status !== 'used') {
          const error = new Error('只有已售账号才能撤回到未售库存');
          error.statusCode = 409;
          error.records = records;
          throw error;
        }

        restoredStatus = record.preparedFor && record.preparedAt ? 'prepared' : 'available';
        return records.map((candidate) => {
          if (candidate.id !== id) return candidate;
          const { usedAt: _usedAt, ...rest } = candidate;
          return {
            ...rest,
            status: restoredStatus,
            remark: '',
          };
        });
      });

      sendJson(res, 200, { records: updatedRecords, restoredStatus });
      return true;
    }

    if (parsedUrl.pathname.startsWith('/api/')) {
      sendJson(res, 404, { error: 'API route not found' });
      return true;
    }

    return false;
  } catch (error) {
    sendJson(res, error.statusCode ?? 400, {
      error: error instanceof Error ? error.message : 'Bad request',
      ...(error.records ? { records: error.records } : {}),
    });
    return true;
  }
}
