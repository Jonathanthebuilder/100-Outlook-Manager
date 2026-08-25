import { mkdtemp, rm } from 'node:fs/promises';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import { afterEach, describe, expect, it } from 'vitest';
import {
  createLedgerStore,
  handleLedgerRequest,
  isUnsoldRecord,
  readJsonBody,
  sendJson,
  validateRecords,
} from './serverLedger.mjs';

const tempDirs = [];

async function makeTempStore() {
  const dir = await mkdtemp(join(tmpdir(), 'outlook-ledger-'));
  tempDirs.push(dir);
  return createLedgerStore(join(dir, 'ledger.json'));
}

afterEach(async () => {
  await Promise.all(tempDirs.splice(0).map((dir) => rm(dir, { recursive: true, force: true })));
});

describe('server ledger store', () => {
  it('starts empty when the ledger file does not exist', async () => {
    const store = await makeTempStore();

    await expect(store.load()).resolves.toEqual([]);
  });

  it('persists records so another store instance reads the same shared ledger', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'outlook-ledger-'));
    tempDirs.push(dir);
    const filePath = join(dir, 'ledger.json');
    const firstStore = createLedgerStore(filePath);
    const secondStore = createLedgerStore(filePath);
    const records = [
      {
        id: 'alice-1',
        email: 'AliceExample1001@outlook.com',
        password: 'pass-one',
        clientId: 'client-a',
        refreshToken: 'refresh-token-a',
        domain: 'outlook.com',
        firstLetter: 'A',
        status: 'available',
        remark: '',
        sourceLineNumber: 1,
        rawCredential: 'AliceExample1001@outlook.com----pass-one----client-a----refresh-token-a',
      },
    ];

    await firstStore.save(records);

    await expect(secondStore.load()).resolves.toEqual(records);
  });

  it('serializes updates against the latest stored records', async () => {
    const store = await makeTempStore();
    await store.save([]);

    await Promise.all([
      store.update((records) => [...records, makeRecord('alice-1')]),
      store.update((records) => [...records, makeRecord('brenda-1')]),
    ]);

    await expect(store.load()).resolves.toEqual([makeRecord('alice-1'), makeRecord('brenda-1')]);
  });

  it('rejects malformed record payloads', () => {
    const invalidRecords = [{ id: 'missing-fields' }];

    expect(() => validateRecords(invalidRecords)).toThrow(/Invalid mailbox record/);
  });

  it('accepts prepared records as unsold inventory only when preparation metadata is complete', () => {
    const prepared = {
      ...makeRecord('prepared-1'),
      status: 'prepared',
      preparedFor: 'Perplexity',
      preparationRemark: '免费账号已注册',
      preparedAt: '2026-08-14T09:00:00.000Z',
    };

    expect(validateRecords([prepared])).toEqual([prepared]);
    expect(isUnsoldRecord(prepared)).toBe(true);
    expect(isUnsoldRecord({ ...prepared, status: 'used' })).toBe(false);
    expect(() => validateRecords([{ ...prepared, preparedFor: '' }])).toThrow(/Invalid mailbox record/);
  });
});

function makeRecord(id) {
  const localPart = id.split('-')[0];
  const email = `${localPart}@outlook.com`;

  return {
    id,
    email,
    password: 'pass-one',
    clientId: 'client-a',
    refreshToken: 'refresh-token-a',
    domain: 'outlook.com',
    firstLetter: localPart.charAt(0).toUpperCase(),
    status: 'available',
    remark: '',
    sourceLineNumber: 1,
    rawCredential: `${email}----pass-one----client-a----refresh-token-a`,
  };
}

describe('server ledger HTTP helpers', () => {
  it('parses JSON request bodies', async () => {
    const request = {
      on(event, callback) {
        if (event === 'data') {
          callback(Buffer.from('{"records":[]}'));
        }

        if (event === 'end') {
          callback();
        }

        return this;
      },
    };

    await expect(readJsonBody(request)).resolves.toEqual({ records: [] });
  });

  it('writes JSON responses', () => {
    const response = {
      headers: undefined,
      statusCode: undefined,
      body: '',
      writeHead(statusCode, headers) {
        this.statusCode = statusCode;
        this.headers = headers;
      },
      end(body) {
        this.body = body;
      },
    };

    sendJson(response, 200, { records: [] });

    expect(response.statusCode).toBe(200);
    expect(response.headers['content-type']).toBe('application/json; charset=utf-8');
    expect(JSON.parse(response.body)).toEqual({ records: [] });
  });
});

describe('server ledger API handler', () => {
  function createResponse() {
    return {
      headers: undefined,
      statusCode: undefined,
      body: '',
      writeHead(statusCode, headers) {
        this.statusCode = statusCode;
        this.headers = headers;
      },
      end(body = '') {
        this.body = body;
      },
    };
  }

  function createJsonRequest(method, url, body) {
    return {
      method,
      url,
      on(event, callback) {
        if (event === 'data' && body !== undefined) {
          callback(Buffer.from(JSON.stringify(body)));
        }

        if (event === 'end') {
          callback();
        }

        return this;
      },
    };
  }

  it('replaces and reads the shared ledger through the API', async () => {
    const store = await makeTempStore();
    const records = [
      {
        id: 'alice-1',
        email: 'AliceExample1001@outlook.com',
        password: 'pass-one',
        clientId: 'client-a',
        refreshToken: 'refresh-token-a',
        domain: 'outlook.com',
        firstLetter: 'A',
        status: 'available',
        remark: '',
        sourceLineNumber: 1,
        rawCredential: 'AliceExample1001@outlook.com----pass-one----client-a----refresh-token-a',
      },
    ];
    const putResponse = createResponse();

    const importedAt = '2026-08-03T10:00:00.000Z';
    await handleLedgerRequest(createJsonRequest('PUT', '/api/records', { records }), putResponse, store, () => importedAt);

    expect(putResponse.statusCode).toBe(200);

    const getResponse = createResponse();
    await handleLedgerRequest(createJsonRequest('GET', '/api/records'), getResponse, store);

    expect(getResponse.statusCode).toBe(200);
    expect(JSON.parse(getResponse.body)).toEqual({
      records: [{ ...records[0], importedAt, tokenStatus: 'unknown' }],
    });
  });

  it('merges imports by email while preserving sold status and remarks', async () => {
    const store = await makeTempStore();
    await store.save([
      {
        id: 'alice-1',
        email: 'AliceExample1001@outlook.com',
        password: 'old-pass',
        clientId: 'old-client',
        refreshToken: 'old-token',
        domain: 'outlook.com',
        firstLetter: 'A',
        status: 'used',
        remark: 'Order 1001',
        usedAt: '2026-08-01T10:00:00.000Z',
        sourceLineNumber: 1,
        rawCredential: 'AliceExample1001@outlook.com----old-pass----old-client----old-token',
      },
    ]);
    const incoming = [
      {
        id: 'alice-new-1',
        email: 'AliceExample1001@outlook.com',
        password: 'new-pass',
        clientId: 'new-client',
        refreshToken: 'new-token',
        domain: 'outlook.com',
        firstLetter: 'A',
        status: 'available',
        remark: '',
        sourceLineNumber: 1,
        rawCredential: 'AliceExample1001@outlook.com----new-pass----new-client----new-token',
      },
    ];
    const response = createResponse();

    await handleLedgerRequest(createJsonRequest('POST', '/api/records/import', { records: incoming }), response, store);

    const payload = JSON.parse(response.body);
    expect(response.statusCode).toBe(200);
    expect(payload).toMatchObject({ created: 0, updated: 1 });
    expect(payload.records[0]).toMatchObject({
      id: 'alice-1',
      password: 'new-pass',
      clientId: 'new-client',
      refreshToken: 'new-token',
      status: 'used',
      remark: 'Order 1001',
      usedAt: '2026-08-01T10:00:00.000Z',
      tokenStatus: 'unknown',
    });
  });

  it('marks an available mailbox as used on the server-side ledger', async () => {
    const store = await makeTempStore();
    const records = [
      {
        id: 'alice-1',
        email: 'AliceExample1001@outlook.com',
        password: 'pass-one',
        clientId: 'client-a',
        refreshToken: 'refresh-token-a',
        domain: 'outlook.com',
        firstLetter: 'A',
        status: 'available',
        remark: '',
        sourceLineNumber: 1,
        rawCredential: 'AliceExample1001@outlook.com----pass-one----client-a----refresh-token-a',
      },
    ];
    await store.save(records);

    const response = createResponse();
    await handleLedgerRequest(
      createJsonRequest('POST', '/api/records/alice-1/use', { remark: '交给 Alex 使用' }),
      response,
      store,
      () => '2026-06-20T09:30:00.000Z',
    );

    expect(response.statusCode).toBe(200);
    expect(JSON.parse(response.body).records[0]).toMatchObject({
      id: 'alice-1',
      status: 'used',
      remark: '交给 Alex 使用',
      usedAt: '2026-06-20T09:30:00.000Z',
    });
  });

  it('moves an available mailbox into and out of prepared inventory', async () => {
    const store = await makeTempStore();
    await store.save([makeRecord('alice-1')]);

    const prepareResponse = createResponse();
    await handleLedgerRequest(
      createJsonRequest('POST', '/api/records/alice-1/prepare', {
        preparedFor: 'Perplexity',
        preparationRemark: '免费账号已注册',
      }),
      prepareResponse,
      store,
      () => '2026-08-14T09:00:00.000Z',
    );

    expect(prepareResponse.statusCode).toBe(200);
    expect(JSON.parse(prepareResponse.body).records[0]).toMatchObject({
      status: 'prepared',
      preparedFor: 'Perplexity',
      preparationRemark: '免费账号已注册',
      preparedAt: '2026-08-14T09:00:00.000Z',
    });

    const unprepareResponse = createResponse();
    await handleLedgerRequest(
      createJsonRequest('POST', '/api/records/alice-1/unprepare', {}),
      unprepareResponse,
      store,
    );

    const restored = JSON.parse(unprepareResponse.body).records[0];
    expect(unprepareResponse.statusCode).toBe(200);
    expect(restored.status).toBe('available');
    expect(restored).not.toHaveProperty('preparedFor');
    expect(restored).not.toHaveProperty('preparedAt');
  });

  it('allows a prepared mailbox to be marked as used while retaining preparation history', async () => {
    const store = await makeTempStore();
    await store.save([{
      ...makeRecord('alice-1'),
      status: 'prepared',
      preparedFor: 'Perplexity',
      preparationRemark: '免费账号已注册',
      preparedAt: '2026-08-14T09:00:00.000Z',
    }]);

    const response = createResponse();
    await handleLedgerRequest(
      createJsonRequest('POST', '/api/records/alice-1/use', { remark: 'Order 1001' }),
      response,
      store,
      () => '2026-08-14T10:00:00.000Z',
    );

    expect(response.statusCode).toBe(200);
    expect(JSON.parse(response.body).records[0]).toMatchObject({
      status: 'used',
      preparedFor: 'Perplexity',
      preparedAt: '2026-08-14T09:00:00.000Z',
      remark: 'Order 1001',
      usedAt: '2026-08-14T10:00:00.000Z',
    });
  });

  it('blocks mailbox reads for used records before making an upstream request', async () => {
    const store = await makeTempStore();
    await store.save([
      {
        id: 'sold-1',
        email: 'SoldExample@outlook.com',
        password: 'pass-one',
        clientId: 'client-a',
        refreshToken: 'refresh-token-a',
        domain: 'outlook.com',
        firstLetter: 'S',
        status: 'used',
        remark: 'Order 1001',
        usedAt: '2026-08-03T10:00:00.000Z',
        sourceLineNumber: 1,
        rawCredential: 'SoldExample@outlook.com----pass-one----client-a----refresh-token-a',
      },
    ]);
    const response = createResponse();

    await handleLedgerRequest(
      createJsonRequest('POST', '/api/records/sold-1/mail', { keyword: 'Perplexity' }),
      response,
      store,
    );

    expect(response.statusCode).toBe(409);
    expect(JSON.parse(response.body).error).toContain('已售账号');
  });

  it('rejects marking a mailbox that is already used', async () => {
    const store = await makeTempStore();
    await store.save([
      {
        id: 'alice-1',
        email: 'AliceExample1001@outlook.com',
        password: 'pass-one',
        clientId: 'client-a',
        refreshToken: 'refresh-token-a',
        domain: 'outlook.com',
        firstLetter: 'A',
        status: 'used',
        remark: '已交付',
        sourceLineNumber: 1,
        rawCredential: 'AliceExample1001@outlook.com----pass-one----client-a----refresh-token-a',
      },
    ]);

    const response = createResponse();
    await handleLedgerRequest(
      createJsonRequest('POST', '/api/records/alice-1/use', { remark: '交给 Alex 使用' }),
      response,
      store,
    );

    expect(response.statusCode).toBe(409);
    expect(JSON.parse(response.body).records[0].remark).toBe('已交付');
  });

  it('rolls an accidentally sold mailbox back to available inventory', async () => {
    const store = await makeTempStore();
    await store.save([
      {
        id: 'sold-1',
        email: 'SoldExample@outlook.com',
        password: 'pass-one',
        clientId: 'client-a',
        refreshToken: 'refresh-token-a',
        domain: 'outlook.com',
        firstLetter: 'S',
        status: 'used',
        remark: 'Order 1001',
        usedAt: '2026-08-25T05:00:00.000Z',
        sourceLineNumber: 1,
        rawCredential: 'SoldExample@outlook.com----pass-one----client-a----refresh-token-a',
      },
    ]);
    const response = createResponse();

    await handleLedgerRequest(
      createJsonRequest('POST', '/api/records/sold-1/rollback-sale', {}),
      response,
      store,
    );

    const payload = JSON.parse(response.body);
    expect(response.statusCode).toBe(200);
    expect(payload.restoredStatus).toBe('available');
    expect(payload.records[0]).toMatchObject({ status: 'available', remark: '' });
    expect(payload.records[0]).not.toHaveProperty('usedAt');
  });

  it('restores a previously prepared mailbox to prepared inventory', async () => {
    const store = await makeTempStore();
    await store.save([
      {
        id: 'prepared-sold-1',
        email: 'PreparedSold@outlook.com',
        password: 'pass-one',
        clientId: 'client-a',
        refreshToken: 'refresh-token-a',
        domain: 'outlook.com',
        firstLetter: 'P',
        status: 'used',
        remark: 'Order 1002',
        preparedFor: 'Perplexity',
        preparedAt: '2026-08-24T08:00:00.000Z',
        usedAt: '2026-08-25T05:00:00.000Z',
        sourceLineNumber: 1,
        rawCredential: 'PreparedSold@outlook.com----pass-one----client-a----refresh-token-a',
      },
    ]);
    const response = createResponse();

    await handleLedgerRequest(
      createJsonRequest('POST', '/api/records/prepared-sold-1/rollback-sale', {}),
      response,
      store,
    );

    const payload = JSON.parse(response.body);
    expect(response.statusCode).toBe(200);
    expect(payload.restoredStatus).toBe('prepared');
    expect(payload.records[0]).toMatchObject({
      status: 'prepared',
      preparedFor: 'Perplexity',
      preparedAt: '2026-08-24T08:00:00.000Z',
      remark: '',
    });
    expect(payload.records[0]).not.toHaveProperty('usedAt');
  });
});
