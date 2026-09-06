import { createReadStream, existsSync } from 'node:fs';
import { stat } from 'node:fs/promises';
import { createServer } from 'node:http';
import { extname, join, normalize, resolve, sep } from 'node:path';
import { fileURLToPath } from 'node:url';
import { getAuthResult } from './serverAuth.mjs';
import {
  createAutomaticTokenMaintainer,
  createLedgerStore,
  handleLedgerRequest,
  startAutomaticTokenMaintenanceScheduler,
} from './serverLedger.mjs';

const root = resolve(fileURLToPath(new URL('..', import.meta.url)));
const dist = join(root, 'dist');
const port = Number(process.env.PORT || 4173);
const host = process.env.HOST || '0.0.0.0';
const dataDir = process.env.OUTLOOK_MANAGER_DATA_DIR || join(root, 'data');
const ledgerStore = createLedgerStore(join(dataDir, 'ledger.json'));
const automaticRefreshEnabled = /^(1|true|yes)$/i.test(process.env.AUTO_TOKEN_REFRESH_ENABLED || '');
const automaticRefreshBatchSize = Math.min(
  Math.max(Number.parseInt(process.env.AUTO_TOKEN_REFRESH_BATCH_SIZE || '30', 10) || 30, 1),
  100,
);
const automaticRefreshMaxAgeDays = Math.max(
  Number.parseInt(process.env.AUTO_TOKEN_REFRESH_MAX_AGE_DAYS || '30', 10) || 30,
  1,
);
const automaticRefreshMinIntervalHours = Math.max(
  Number.parseInt(process.env.AUTO_TOKEN_REFRESH_MIN_INTERVAL_HOURS || '23', 10) || 23,
  1,
);
const automaticRefreshPauseMs = Math.max(
  Number.parseInt(process.env.AUTO_TOKEN_REFRESH_PAUSE_MS || '1000', 10) || 1_000,
  0,
);

const contentTypes = new Map([
  ['.css', 'text/css; charset=utf-8'],
  ['.html', 'text/html; charset=utf-8'],
  ['.ico', 'image/x-icon'],
  ['.js', 'text/javascript; charset=utf-8'],
  ['.json', 'application/json; charset=utf-8'],
  ['.map', 'application/json; charset=utf-8'],
  ['.svg', 'image/svg+xml'],
  ['.txt', 'text/plain; charset=utf-8'],
  ['.webp', 'image/webp'],
]);

function resolveRequestPath(url) {
  const parsedUrl = new URL(url, `http://localhost:${port}`);
  const pathname = decodeURIComponent(parsedUrl.pathname);
  const relativePath = pathname === '/' ? 'outlook-manager.html' : pathname.slice(1);
  const normalizedPath = normalize(relativePath).replace(/^(\.\.(\/|\\|$))+/, '');
  const absolutePath = join(dist, normalizedPath);

  if (!absolutePath.startsWith(dist + sep) && absolutePath !== dist) {
    return null;
  }

  return absolutePath;
}

async function serveFile(res, filePath) {
  try {
    const fileStat = await stat(filePath);

    if (!fileStat.isFile()) {
      res.writeHead(404, { 'content-type': 'text/plain; charset=utf-8' });
      res.end('Not found');
      return;
    }

    const contentType = contentTypes.get(extname(filePath)) || 'application/octet-stream';
    res.writeHead(200, {
      'cache-control': 'no-store',
      'content-length': fileStat.size,
      'content-type': contentType,
    });
    createReadStream(filePath).pipe(res);
  } catch {
    res.writeHead(404, { 'content-type': 'text/plain; charset=utf-8' });
    res.end('Not found');
  }
}

const server = createServer(async (req, res) => {
  if (!req.url || !['GET', 'HEAD', 'POST', 'PUT'].includes(req.method ?? '')) {
    res.writeHead(405, { 'content-type': 'text/plain; charset=utf-8' });
    res.end('Method not allowed');
    return;
  }

  const authResult = getAuthResult(req.headers);

  if (!authResult.allowed) {
    res.writeHead(authResult.statusCode, authResult.headers);
    res.end(authResult.body);
    return;
  }

  if (await handleLedgerRequest(req, res, ledgerStore)) {
    return;
  }

  const requestedPath = resolveRequestPath(req.url);

  if (!requestedPath) {
    res.writeHead(400, { 'content-type': 'text/plain; charset=utf-8' });
    res.end('Bad request');
    return;
  }

  if (existsSync(requestedPath)) {
    await serveFile(res, requestedPath);
    return;
  }

  await serveFile(res, join(dist, 'outlook-manager.html'));
});

let maintenanceScheduler;
if (automaticRefreshEnabled) {
  const maintainer = createAutomaticTokenMaintainer({
    store: ledgerStore,
    stateFilePath: join(dataDir, 'token-maintenance.json'),
    batchSize: automaticRefreshBatchSize,
    maxAgeMs: automaticRefreshMaxAgeDays * 24 * 60 * 60 * 1000,
    minimumRunIntervalMs: automaticRefreshMinIntervalHours * 60 * 60 * 1000,
    pauseMs: automaticRefreshPauseMs,
  });
  maintenanceScheduler = startAutomaticTokenMaintenanceScheduler({
    maintainer,
    onResult(result) {
      if (result.status !== 'completed') return;
      console.log(
        `Automatic Token maintenance completed: total=${result.total} success=${result.successCount} failure=${result.failureCount}`,
      );
    },
    onError(error) {
      console.error('Automatic Token maintenance failed:', error instanceof Error ? error.message : error);
    },
  });
  console.log(
    `Automatic Token maintenance enabled: batch=${automaticRefreshBatchSize}, maxAgeDays=${automaticRefreshMaxAgeDays}, minIntervalHours=${automaticRefreshMinIntervalHours}`,
  );
}

server.on('close', () => maintenanceScheduler?.stop());

server.listen(port, host, () => {
  console.log(`Outlook Manager is running at http://${host}:${port}`);
});
