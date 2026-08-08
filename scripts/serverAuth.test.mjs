import { describe, expect, it } from 'vitest';
import { getAuthResult } from './serverAuth.mjs';

function basicAuth(username, password) {
  return `Basic ${Buffer.from(`${username}:${password}`).toString('base64')}`;
}

describe('server auth', () => {
  it('allows requests when no password is configured', () => {
    const result = getAuthResult({}, {});

    expect(result.allowed).toBe(true);
    expect(result.statusCode).toBeUndefined();
  });

  it('challenges requests without credentials when a password is configured', () => {
    const result = getAuthResult({}, { OUTLOOK_MANAGER_PASSWORD: 'secret' });

    expect(result.allowed).toBe(false);
    expect(result.statusCode).toBe(401);
    expect(result.headers['www-authenticate']).toContain('Outlook Manager');
  });

  it('rejects wrong credentials', () => {
    const result = getAuthResult(
      { authorization: basicAuth('outlook', 'wrong') },
      { OUTLOOK_MANAGER_PASSWORD: 'secret' },
    );

    expect(result.allowed).toBe(false);
    expect(result.statusCode).toBe(401);
  });

  it('allows correct credentials with the default username', () => {
    const result = getAuthResult(
      { authorization: basicAuth('outlook', 'secret') },
      { OUTLOOK_MANAGER_PASSWORD: 'secret' },
    );

    expect(result.allowed).toBe(true);
  });

  it('allows a configured username', () => {
    const result = getAuthResult(
      { authorization: basicAuth('jonathan', 'secret') },
      {
        OUTLOOK_MANAGER_PASSWORD: 'secret',
        OUTLOOK_MANAGER_USERNAME: 'jonathan',
      },
    );

    expect(result.allowed).toBe(true);
  });
});
