import { mkdtemp, readFile, rm } from 'node:fs/promises';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import { afterEach, describe, expect, it } from 'vitest';
import {
  createLedgerStore,
  handleLedgerRequest,
  isUnsoldRecord,
  readJsonBody,
  selectRefreshTargets,
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

describe('Token refresh target selection', () => {
  it('selects only unverified unsold mailboxes from the requested import batch', () => {
    const records = [
      { ...makeRecord('batch-a-unknown'), importedAt: 'batch-a', clientId: 'client-a', tokenStatus: 'unknown' },
      { ...makeRecord('batch-a-healthy'), importedAt: 'batch-a', clientId: 'client-a', tokenStatus: 'healthy' },
      { ...makeRecord('batch-a-error'), importedAt: 'batch-a', clientId: 'client-a', tokenStatus: 'error' },
      { ...makeRecord('batch-b-unknown'), importedAt: 'batch-b', clientId: 'client-a', tokenStatus: 'unknown' },
      { ...makeRecord('other-client'), importedAt: 'batch-a', clientId: 'client-b', tokenStatus: 'unknown' },
      { ...makeRecord('sold-unknown'), importedAt: 'batch-a', clientId: 'client-a', tokenStatus: 'unknown', status: 'used' },
    ];

    const selected = selectRefreshTargets(records, {
      limit: 10,
      importedAt: 'batch-a',
      clientId: 'client-a',
      tokenStatuses: ['unknown'],
    });

    expect(selected.map((record) => record.id)).toEqual(['batch-a-unknown']);
  });

  it('selects only healthy unsold mailboxes whose last successful check is due', () => {
    const records = [
      { ...makeRecord('missing-check'), tokenStatus: 'healthy' },
      { ...makeRecord('oldest-healthy'), tokenStatus: 'healthy', tokenCheckedAt: '2026-07-01T00:00:00.000Z' },
      { ...makeRecord('boundary-healthy'), tokenStatus: 'healthy', tokenCheckedAt: '2026-08-08T12:00:00.000Z' },
      { ...makeRecord('recent-healthy'), tokenStatus: 'healthy', tokenCheckedAt: '2026-08-20T00:00:00.000Z' },
      { ...makeRecord('old-error'), tokenStatus: 'error', tokenCheckedAt: '2026-07-01T00:00:00.000Z' },
      { ...makeRecord('old-sold'), tokenStatus: 'healthy', tokenCheckedAt: '2026-07-01T00:00:00.000Z', status: 'used' },
    ];

    const selected = selectRefreshTargets(records, {
      limit: 30,
      tokenStatuses: ['healthy'],
      checkedBefore: '2026-08-08T12:00:00.000Z',
    });

    expect(selected.map((record) => record.id)).toEqual([
      'missing-check',
      'oldest-healthy',
      'boundary-healthy',
    ]);
  });
});

describe('automatic Token maintenance', () => {
  it('refreshes only due healthy unsold mailboxes and persists rotated credentials', async () => {
    const ledgerModule = await import('./serverLedger.mjs');
    expect(ledgerModule.createAutomaticTokenMaintainer).toBeTypeOf('function');
    if (!ledgerModule.createAutomaticTokenMaintainer) return;

    const dir = await mkdtemp(join(tmpdir(), 'outlook-maintenance-'));
    tempDirs.push(dir);
    const store = createLedgerStore(join(dir, 'ledger.json'));
    const stateFilePath = join(dir, 'token-maintenance.json');
    await store.save([
      { ...makeRecord('due-available'), tokenStatus: 'healthy', tokenCheckedAt: '2026-07-01T00:00:00.000Z' },
      {
        ...makeRecord('due-prepared'),
        status: 'prepared',
        preparedFor: 'Perplexity',
        preparedAt: '2026-08-01T00:00:00.000Z',
        tokenStatus: 'healthy',
        tokenCheckedAt: '2026-07-02T00:00:00.000Z',
      },
      { ...makeRecord('recent-available'), tokenStatus: 'healthy', tokenCheckedAt: '2026-08-20T00:00:00.000Z' },
      { ...makeRecord('due-error'), tokenStatus: 'error', tokenCheckedAt: '2026-07-01T00:00:00.000Z' },
      { ...makeRecord('due-sold'), status: 'used', tokenStatus: 'healthy', tokenCheckedAt: '2026-07-01T00:00:00.000Z' },
    ]);

    const maintainer = ledgerModule.createAutomaticTokenMaintainer({
      store,
      stateFilePath,
      refreshToken: async (_clientId, refreshToken) => ({ refreshToken: `${refreshToken}-rotated` }),
      now: () => '2026-09-07T12:00:00.000Z',
      batchSize: 30,
      maxAgeMs: 30 * 24 * 60 * 60 * 1000,
      minimumRunIntervalMs: 23 * 60 * 60 * 1000,
      pauseMs: 0,
    });

    await expect(maintainer.runIfDue()).resolves.toMatchObject({
      status: 'completed',
      total: 2,
      successCount: 2,
      failureCount: 0,
    });

    const records = await store.load();
    expect(records.find((record) => record.id === 'due-available')).toMatchObject({
      refreshToken: 'refresh-token-a-rotated',
      rawCredential: 'due@outlook.com----pass-one----client-a----refresh-token-a-rotated',
      tokenStatus: 'healthy',
      tokenCheckedAt: '2026-09-07T12:00:00.000Z',
      tokenRefreshedAt: '2026-09-07T12:00:00.000Z',
    });
    expect(records.find((record) => record.id === 'recent-available')?.refreshToken).toBe('refresh-token-a');
    expect(records.find((record) => record.id === 'due-error')?.refreshToken).toBe('refresh-token-a');
    expect(records.find((record) => record.id === 'due-sold')?.refreshToken).toBe('refresh-token-a');

    const state = JSON.parse(await readFile(stateFilePath, 'utf8'));
    expect(state).toMatchObject({
      lastRunAt: '2026-09-07T12:00:00.000Z',
      status: 'completed',
      total: 2,
      successCount: 2,
      failureCount: 0,
    });
  });

  it('persists the daily run time so a service restart cannot run another batch too soon', async () => {
    const ledgerModule = await import('./serverLedger.mjs');
    const dir = await mkdtemp(join(tmpdir(), 'outlook-maintenance-interval-'));
    tempDirs.push(dir);
    const store = createLedgerStore(join(dir, 'ledger.json'));
    const stateFilePath = join(dir, 'token-maintenance.json');
    await store.save([
      { ...makeRecord('due-once'), tokenStatus: 'healthy', tokenCheckedAt: '2026-07-01T00:00:00.000Z' },
    ]);

    const firstMaintainer = ledgerModule.createAutomaticTokenMaintainer({
      store,
      stateFilePath,
      refreshToken: async (_clientId, refreshToken) => ({ refreshToken }),
      now: () => '2026-09-07T12:00:00.000Z',
      minimumRunIntervalMs: 23 * 60 * 60 * 1000,
      pauseMs: 0,
    });
    await firstMaintainer.runIfDue();

    const restartedMaintainer = ledgerModule.createAutomaticTokenMaintainer({
      store,
      stateFilePath,
      refreshToken: async (_clientId, refreshToken) => ({ refreshToken }),
      now: () => '2026-09-07T13:00:00.000Z',
      minimumRunIntervalMs: 23 * 60 * 60 * 1000,
      pauseMs: 0,
    });

    await expect(restartedMaintainer.runIfDue()).resolves.toEqual({
      status: 'skipped',
      reason: 'interval-not-elapsed',
      lastRunAt: '2026-09-07T12:00:00.000Z',
    });
  });

  it('rejects an overlapping maintenance run before it can start another refresh', async () => {
    const ledgerModule = await import('./serverLedger.mjs');
    const dir = await mkdtemp(join(tmpdir(), 'outlook-maintenance-lock-'));
    tempDirs.push(dir);
    const store = createLedgerStore(join(dir, 'ledger.json'));
    await store.save([
      { ...makeRecord('due-locked'), tokenStatus: 'healthy', tokenCheckedAt: '2026-07-01T00:00:00.000Z' },
    ]);

    let releaseRefresh;
    let signalStarted;
    const refreshStarted = new Promise((resolve) => { signalStarted = resolve; });
    const refreshGate = new Promise((resolve) => { releaseRefresh = resolve; });
    const maintainer = ledgerModule.createAutomaticTokenMaintainer({
      store,
      stateFilePath: join(dir, 'token-maintenance.json'),
      refreshToken: async (_clientId, refreshToken) => {
        signalStarted();
        await refreshGate;
        return { refreshToken };
      },
      now: () => '2026-09-07T12:00:00.000Z',
      pauseMs: 0,
    });

    const firstRun = maintainer.runIfDue();
    await refreshStarted;
    await expect(maintainer.runIfDue()).resolves.toEqual({
      status: 'skipped',
      reason: 'already-running',
    });
    releaseRefresh();
    await firstRun;
  });

  it('quarantines a failed Token and continues refreshing the remaining due inventory', async () => {
    const ledgerModule = await import('./serverLedger.mjs');
    const dir = await mkdtemp(join(tmpdir(), 'outlook-maintenance-failure-'));
    tempDirs.push(dir);
    const store = createLedgerStore(join(dir, 'ledger.json'));
    await store.save([
      {
        ...makeRecord('bad-token'),
        refreshToken: 'invalid-token',
        rawCredential: 'bad@outlook.com----pass-one----client-a----invalid-token',
        tokenStatus: 'healthy',
        tokenCheckedAt: '2026-07-01T00:00:00.000Z',
      },
      {
        ...makeRecord('good-token'),
        refreshToken: 'valid-token',
        rawCredential: 'good@outlook.com----pass-one----client-a----valid-token',
        tokenStatus: 'healthy',
        tokenCheckedAt: '2026-07-02T00:00:00.000Z',
      },
    ]);
    const maintainer = ledgerModule.createAutomaticTokenMaintainer({
      store,
      stateFilePath: join(dir, 'token-maintenance.json'),
      refreshToken: async (_clientId, refreshToken) => {
        if (refreshToken === 'invalid-token') throw new Error('grant expired');
        return { refreshToken: 'valid-token-rotated' };
      },
      now: () => '2026-09-07T12:00:00.000Z',
      pauseMs: 0,
    });

    await expect(maintainer.runIfDue()).resolves.toMatchObject({
      status: 'completed',
      total: 2,
      successCount: 1,
      failureCount: 1,
    });

    const records = await store.load();
    expect(records.find((record) => record.id === 'bad-token')).toMatchObject({
      refreshToken: 'invalid-token',
      tokenStatus: 'error',
      tokenCheckedAt: '2026-09-07T12:00:00.000Z',
      tokenError: 'grant expired',
    });
    expect(records.find((record) => record.id === 'good-token')).toMatchObject({
      refreshToken: 'valid-token-rotated',
      tokenStatus: 'healthy',
      tokenCheckedAt: '2026-09-07T12:00:00.000Z',
    });
  });

  it('schedules an initial maintenance check and recurring hourly checks', async () => {
    const ledgerModule = await import('./serverLedger.mjs');
    expect(ledgerModule.startAutomaticTokenMaintenanceScheduler).toBeTypeOf('function');
    if (!ledgerModule.startAutomaticTokenMaintenanceScheduler) return;

    const scheduled = {};
    const cleared = [];
    let runCount = 0;
    const results = [];
    const scheduler = ledgerModule.startAutomaticTokenMaintenanceScheduler({
      maintainer: {
        async runIfDue() {
          runCount += 1;
          return { status: 'completed', total: 1, successCount: 1, failureCount: 0 };
        },
      },
      initialDelayMs: 60_000,
      pollIntervalMs: 60 * 60 * 1000,
      setTimeoutFn(callback, delay) {
        scheduled.initial = { callback, delay };
        return 'initial-timer';
      },
      setIntervalFn(callback, delay) {
        scheduled.interval = { callback, delay };
        return 'interval-timer';
      },
      clearTimeoutFn(timer) { cleared.push(timer); },
      clearIntervalFn(timer) { cleared.push(timer); },
      onResult(result) { results.push(result); },
    });

    expect(scheduled.initial.delay).toBe(60_000);
    expect(scheduled.interval.delay).toBe(60 * 60 * 1000);
    await scheduled.initial.callback();
    await scheduled.interval.callback();
    expect(runCount).toBe(2);
    expect(results).toEqual([
      { status: 'completed', total: 1, successCount: 1, failureCount: 0 },
      { status: 'completed', total: 1, successCount: 1, failureCount: 0 },
    ]);

    scheduler.stop();
    expect(cleared).toEqual(['initial-timer', 'interval-timer']);
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
        tokenStatus: 'healthy',
        tokenCheckedAt: '2026-06-20T09:00:00.000Z',
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
    await store.save([{ ...makeRecord('alice-1'), tokenStatus: 'healthy' }]);

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

  it('rejects preparing a mailbox whose Token has not been verified', async () => {
    const store = await makeTempStore();
    await store.save([{ ...makeRecord('alice-1'), tokenStatus: 'unknown' }]);

    const response = createResponse();
    await handleLedgerRequest(
      createJsonRequest('POST', '/api/records/alice-1/prepare', {
        preparedFor: 'Perplexity',
        preparationRemark: '',
      }),
      response,
      store,
    );

    expect(response.statusCode).toBe(409);
    expect(JSON.parse(response.body).error).toContain('Token');
    expect((await store.load())[0].status).toBe('available');
  });

  it('allows a prepared mailbox to be marked as used while retaining preparation history', async () => {
    const store = await makeTempStore();
    await store.save([{
      ...makeRecord('alice-1'),
      status: 'prepared',
      preparedFor: 'Perplexity',
      preparationRemark: '免费账号已注册',
      preparedAt: '2026-08-14T09:00:00.000Z',
      tokenStatus: 'healthy',
      tokenCheckedAt: '2026-08-14T09:30:00.000Z',
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

  it('rejects selling a mailbox when its last healthy Token check is older than 24 hours', async () => {
    const store = await makeTempStore();
    await store.save([{
      ...makeRecord('alice-1'),
      tokenStatus: 'healthy',
      tokenCheckedAt: '2026-08-13T09:59:59.000Z',
    }]);

    const response = createResponse();
    await handleLedgerRequest(
      createJsonRequest('POST', '/api/records/alice-1/use', { remark: 'Order 1001' }),
      response,
      store,
      () => '2026-08-14T10:00:00.000Z',
    );

    expect(response.statusCode).toBe(409);
    expect(JSON.parse(response.body).error).toContain('24 小时');
    expect((await store.load())[0].status).toBe('available');
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
