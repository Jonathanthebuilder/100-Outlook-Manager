import tls from 'node:tls';
import { Buffer } from 'node:buffer';

const oauthScope = 'https://outlook.office.com/IMAP.AccessAsUser.All offline_access';
const tokenEndpoint = 'https://login.microsoftonline.com/consumers/oauth2/v2.0/token';

async function readOAuthResponse(response, action) {
  const data = await response.json().catch(() => ({}));
  if (!response.ok) {
    const description = data.error_description || data.error || `HTTP ${response.status}`;
    const error = new Error(`${action}失败：${description}`);
    error.code = data.error || 'oauth_error';
    throw error;
  }
  return data;
}

export async function refreshOAuthToken(clientId, refreshToken) {
  const response = await fetch(tokenEndpoint, {
    method: 'POST',
    headers: { 'content-type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      client_id: clientId,
      grant_type: 'refresh_token',
      refresh_token: refreshToken,
      scope: oauthScope,
    }),
  });
  const data = await readOAuthResponse(response, 'Microsoft Token 刷新');
  if (!data.access_token) throw new Error('Microsoft 没有返回 access_token');
  return {
    accessToken: data.access_token,
    refreshToken: data.refresh_token || refreshToken,
    expiresIn: Number(data.expires_in || 0),
  };
}

export async function createDeviceCode(clientId) {
  const response = await fetch('https://login.microsoftonline.com/consumers/oauth2/v2.0/devicecode', {
    method: 'POST',
    headers: { 'content-type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({ client_id: clientId, scope: oauthScope }),
  });
  return readOAuthResponse(response, 'Microsoft 设备码生成');
}

export async function pollDeviceCode(clientId, deviceCode) {
  const response = await fetch(tokenEndpoint, {
    method: 'POST',
    headers: { 'content-type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      grant_type: 'urn:ietf:params:oauth:grant-type:device_code',
      client_id: clientId,
      device_code: deviceCode,
    }),
  });
  const data = await response.json().catch(() => ({}));
  if (data.error === 'authorization_pending' || data.error === 'slow_down') {
    return { pending: true, slowDown: data.error === 'slow_down' };
  }
  if (!response.ok || !data.refresh_token) {
    const description = data.error_description || data.error || `HTTP ${response.status}`;
    const error = new Error(`Microsoft 授权失败：${description}`);
    error.code = data.error || 'oauth_error';
    throw error;
  }
  return { pending: false, refreshToken: data.refresh_token };
}

class ImapClient {
  constructor(host = 'outlook.office365.com', port = 993) {
    this.host = host;
    this.port = port;
    this.socket = null;
    this.buffer = Buffer.alloc(0);
    this.tagCounter = 0;
    this.socketError = null;
    this.socketClosed = false;
  }

  async connect() {
    await new Promise((resolve, reject) => {
      const socket = tls.connect({ host: this.host, port: this.port, servername: this.host });
      this.socket = socket;
      const timer = setTimeout(() => reject(new Error('连接 Outlook IMAP 超时')), 15_000);
      const onData = (chunk) => {
        this.buffer = Buffer.concat([this.buffer, Buffer.from(chunk)]);
        if (this.buffer.includes(Buffer.from('* OK'))) {
          clearTimeout(timer);
          resolve();
        }
      };
      const onError = (error) => {
        this.socketError = error;
        clearTimeout(timer);
        reject(error);
      };
      socket.on('data', onData);
      socket.on('error', onError);
      socket.on('close', () => {
        this.socketClosed = true;
      });
    });
  }

  nextTag() {
    this.tagCounter += 1;
    return `A${String(this.tagCounter).padStart(4, '0')}`;
  }

  async command(command) {
    const tag = this.nextTag();
    const startOffset = this.buffer.length;
    this.socket.write(`${tag} ${command}\r\n`);
    return new Promise((resolve, reject) => {
      const started = Date.now();
      const wait = () => {
        if (this.socketError) {
          reject(new Error(`Outlook IMAP 连接中断：${this.socketError.message}`));
          return;
        }
        if (this.socketClosed) {
          reject(new Error('Outlook IMAP 连接已关闭'));
          return;
        }
        const chunk = this.buffer.subarray(startOffset).toString('binary');
        const done = chunk.match(new RegExp(`(?:^|\\r?\\n)${tag} (OK|NO|BAD) ([^\\r\\n]*)`));
        if (done) {
          if (done[1] === 'OK') resolve(this.buffer.subarray(startOffset));
          else reject(new Error(`IMAP ${command.split(' ')[0]} 失败：${done[2].trim()}`));
          return;
        }
        if (Date.now() - started > 30_000) {
          reject(new Error(`IMAP ${command.split(' ')[0]} 超时`));
          return;
        }
        setTimeout(wait, 25);
      };
      wait();
    });
  }

  async authenticate(email, accessToken) {
    const xoauth = Buffer.from(`user=${email}\x01auth=Bearer ${accessToken}\x01\x01`).toString('base64');
    await this.command(`AUTHENTICATE XOAUTH2 ${xoauth}`);
  }

  async select(folder) {
    await this.command(`SELECT "${folder}"`);
  }

  async searchAll() {
    const response = (await this.command('UID SEARCH ALL')).toString('binary');
    const line = response.split(/\r?\n/).find((item) => item.startsWith('* SEARCH')) || '* SEARCH';
    return line.replace('* SEARCH', '').trim().split(/\s+/).filter(Boolean).map(Number);
  }

  async fetchRaw(uid) {
    const response = await this.command(`UID FETCH ${uid} (RFC822)`);
    const header = response.toString('binary');
    const literal = header.match(/RFC822 \{(\d+)\}\r?\n/);
    if (!literal || literal.index === undefined) return Buffer.alloc(0);
    const start = literal.index + literal[0].length;
    return response.subarray(start, start + Number.parseInt(literal[1], 10));
  }

  close() {
    if (!this.socket) return;
    try {
      this.socket.write('ZZZZ LOGOUT\r\n');
      this.socket.end();
    } catch {
      this.socket.destroy();
    }
  }
}

function decodeMimeWords(value = '') {
  return String(value).replace(/=\?([^?]+)\?([bqBQ])\?([^?]*)\?=/g, (_match, charset, encoding, text) => {
    try {
      const bytes = encoding.toUpperCase() === 'B'
        ? Buffer.from(text, 'base64')
        : Buffer.from(text.replace(/_/g, ' ').replace(/=([0-9A-F]{2})/gi, (_m, hex) => String.fromCharCode(parseInt(hex, 16))), 'binary');
      return new TextDecoder(charset.toLowerCase()).decode(bytes);
    } catch {
      return text;
    }
  });
}

function getHeader(headers, name) {
  const unfolded = headers.replace(/\r?\n[ \t]+/g, ' ');
  const match = unfolded.match(new RegExp(`^${name}:\\s*([^\\r\\n]*)`, 'im'));
  return decodeMimeWords(match?.[1]?.trim() || '');
}

function decodeQuotedPrintable(text) {
  return String(text || '').replace(/=\r?\n/g, '').replace(/=([0-9A-F]{2})/gi, (_m, hex) => String.fromCharCode(parseInt(hex, 16)));
}

function stripHtml(html) {
  return String(html || '')
    .replace(/<style[\s\S]*?<\/style>/gi, ' ')
    .replace(/<script[\s\S]*?<\/script>/gi, ' ')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>/gi, '\n')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/[ \t]+\n/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

function decodeBody(body, encoding) {
  const normalized = String(encoding || '').toLowerCase();
  if (normalized.includes('base64')) return Buffer.from(body.replace(/\s+/g, ''), 'base64').toString('utf8');
  if (normalized.includes('quoted-printable')) return Buffer.from(decodeQuotedPrintable(body), 'binary').toString('utf8');
  return body;
}

function splitHeadersAndBody(raw) {
  const crlf = raw.indexOf('\r\n\r\n');
  const lf = raw.indexOf('\n\n');
  const split = crlf === -1 ? lf : crlf;
  if (split === -1) return { headers: raw, body: '' };
  const separatorLength = raw.startsWith('\r\n\r\n', split) ? 4 : 2;
  return { headers: raw.slice(0, split), body: raw.slice(split + separatorLength) };
}

function extractBody(raw) {
  const { headers, body } = splitHeadersAndBody(raw);
  const contentType = getHeader(headers, 'Content-Type');
  const encoding = getHeader(headers, 'Content-Transfer-Encoding');
  const boundary = contentType.match(/boundary="?([^";]+)"?/i)?.[1];
  if (!boundary) {
    const decoded = decodeBody(body, encoding);
    return /html/i.test(contentType) ? stripHtml(decoded) : decoded.trim();
  }
  let textPart = '';
  let htmlPart = '';
  for (const part of body.split(`--${boundary}`)) {
    const parsed = splitHeadersAndBody(part);
    const type = getHeader(parsed.headers, 'Content-Type');
    const partEncoding = getHeader(parsed.headers, 'Content-Transfer-Encoding');
    const decoded = decodeBody(parsed.body.replace(/\r?\n--$/, ''), partEncoding);
    if (/text\/plain/i.test(type) && !textPart) textPart = decoded.trim();
    if (/text\/html/i.test(type) && !htmlPart) htmlPart = stripHtml(decoded);
  }
  return (textPart || htmlPart).trim();
}

function extractOtp(text) {
  const prioritized = String(text || '').match(/(?:code|验证码|verification)[^\d]{0,30}(\d{4,8})/i)?.[1];
  return prioritized || (String(text || '').match(/\b\d{4,8}\b/)?.[0] ?? '');
}

function parseEmail(rawBuffer, folder, uid) {
  const raw = rawBuffer.toString('utf8');
  const { headers } = splitHeadersAndBody(raw);
  const subject = getHeader(headers, 'Subject');
  const from = getHeader(headers, 'From');
  const date = getHeader(headers, 'Date');
  const body = extractBody(raw).slice(0, 12_000);
  return {
    folder,
    uid,
    from,
    subject,
    date: date && !Number.isNaN(Date.parse(date)) ? new Date(date).toISOString() : '',
    body,
    otp: extractOtp(`${subject}\n${body}`),
  };
}

async function readFolder(client, folder, keyword, maxCount) {
  await client.select(folder);
  const uids = (await client.searchAll()).slice(-Math.max(maxCount * 4, maxCount)).reverse();
  const messages = [];
  for (const uid of uids) {
    const raw = await client.fetchRaw(uid);
    if (!raw.length) continue;
    const message = parseEmail(raw, folder, uid);
    const searchable = `${message.from}\n${message.subject}\n${message.body}`.toLowerCase();
    if (keyword && !searchable.includes(keyword.toLowerCase())) continue;
    messages.push(message);
    if (messages.length >= maxCount) break;
  }
  return messages;
}

async function readRestFolder(accessToken, folder, keyword, maxCount) {
  const folderName = folder === 'Junk' ? 'junkemail' : 'inbox';
  const query = new URLSearchParams({
    '$top': String(Math.min(Math.max(maxCount * 4, maxCount), 50)),
    '$orderby': 'ReceivedDateTime desc',
    '$select': 'Id,Subject,From,ReceivedDateTime,BodyPreview,Body',
  });
  const response = await fetch(`https://outlook.office.com/api/v2.0/me/mailfolders/${folderName}/messages?${query}`, {
    headers: { authorization: `Bearer ${accessToken}` },
  });
  const data = await response.json().catch(() => ({}));
  if (!response.ok) {
    const description = data.error?.message || data.error?.code || `HTTP ${response.status}`;
    throw new Error(`Outlook 邮件 API 失败：${description}`);
  }
  const messages = [];
  for (const item of Array.isArray(data.value) ? data.value : []) {
    const rawBody = item.Body?.Content || item.BodyPreview || '';
    const body = item.Body?.ContentType === 'HTML' ? stripHtml(rawBody) : String(rawBody).trim();
    const from = item.From?.EmailAddress?.Address || item.From?.EmailAddress?.Name || '';
    const subject = item.Subject || '';
    const searchable = `${from}\n${subject}\n${body}`.toLowerCase();
    if (keyword && !searchable.includes(keyword.toLowerCase())) continue;
    messages.push({
      folder: folder === 'Junk' ? 'Junk' : 'INBOX',
      uid: item.Id || `${folderName}-${messages.length}`,
      from,
      subject,
      date: item.ReceivedDateTime || '',
      body: body.slice(0, 12_000),
      otp: extractOtp(`${subject}\n${body}`),
    });
    if (messages.length >= maxCount) break;
  }
  return messages;
}

async function fetchViaImap({ email, accessToken, folder, keyword, maxCount }) {
  const client = new ImapClient();
  await client.connect();
  try {
    await client.authenticate(email, accessToken);
    const folders = folder === 'ALL' ? ['INBOX', 'Junk'] : [folder];
    const messages = [];
    for (const currentFolder of folders) {
      try {
        messages.push(...await readFolder(client, currentFolder, keyword, maxCount));
      } catch (error) {
        if (folder !== 'ALL' || currentFolder === 'INBOX') throw error;
      }
    }
    return messages;
  } finally {
    client.close();
  }
}

export async function fetchMailboxMessages({ email, clientId, refreshToken, folder = 'ALL', keyword = '', maxCount = 10 }) {
  const token = await refreshOAuthToken(clientId, refreshToken);
  try {
    const folders = folder === 'ALL' ? ['INBOX', 'Junk'] : [folder];
    const messages = [];
    for (const currentFolder of folders) {
      try {
        messages.push(...await readRestFolder(token.accessToken, currentFolder, keyword, maxCount));
      } catch (error) {
        if (folder !== 'ALL' || currentFolder === 'INBOX') throw error;
      }
    }
    messages.sort((a, b) => Date.parse(b.date || '1970-01-01') - Date.parse(a.date || '1970-01-01'));
    return { messages: messages.slice(0, maxCount), refreshToken: token.refreshToken };
  } catch (restError) {
    try {
      const messages = await fetchViaImap({
        email,
        accessToken: token.accessToken,
        folder,
        keyword,
        maxCount,
      });
      messages.sort((a, b) => Date.parse(b.date || '1970-01-01') - Date.parse(a.date || '1970-01-01'));
      return { messages: messages.slice(0, maxCount), refreshToken: token.refreshToken };
    } catch (imapError) {
      throw new Error(`${restError.message}；${imapError.message}`);
    }
  }
}
