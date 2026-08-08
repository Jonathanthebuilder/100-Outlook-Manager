import { timingSafeEqual } from 'node:crypto';

const defaultUsername = 'outlook';
const realm = 'Outlook Manager';

function safeCompare(left, right) {
  const leftBuffer = Buffer.from(left);
  const rightBuffer = Buffer.from(right);

  if (leftBuffer.length !== rightBuffer.length) {
    return false;
  }

  return timingSafeEqual(leftBuffer, rightBuffer);
}

function parseBasicAuth(header) {
  if (!header?.startsWith('Basic ')) {
    return null;
  }

  try {
    const decoded = Buffer.from(header.slice(6), 'base64').toString('utf8');
    const separator = decoded.indexOf(':');

    if (separator === -1) {
      return null;
    }

    return {
      username: decoded.slice(0, separator),
      password: decoded.slice(separator + 1),
    };
  } catch {
    return null;
  }
}

export function getAuthResult(headers, env = process.env) {
  const configuredPassword = env.OUTLOOK_MANAGER_PASSWORD?.trim();

  if (!configuredPassword) {
    return { allowed: true };
  }

  const configuredUsername = env.OUTLOOK_MANAGER_USERNAME?.trim() || defaultUsername;
  const credentials = parseBasicAuth(headers.authorization);
  const challenge = {
    allowed: false,
    statusCode: 401,
    headers: {
      'content-type': 'text/plain; charset=utf-8',
      'www-authenticate': `Basic realm="${realm}", charset="UTF-8"`,
    },
    body: 'Authentication required',
  };

  if (!credentials) {
    return challenge;
  }

  const usernameMatches = safeCompare(credentials.username, configuredUsername);
  const passwordMatches = safeCompare(credentials.password, configuredPassword);

  return usernameMatches && passwordMatches ? { allowed: true } : challenge;
}
