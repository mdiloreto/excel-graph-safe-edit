import { describe, expect, it } from 'vitest';
import { buildAuthorizeUrl } from '../src/oauth.js';

describe('oauth helpers', () => {
  it('builds an auth-code PKCE authorize URL', () => {
    const url = buildAuthorizeUrl({
      authority: 'https://login.microsoftonline.com/consumers',
      clientId: 'client-id',
      redirectUri: 'http://localhost:1234/callback',
      scopes: ['openid', 'https://graph.microsoft.com/Files.ReadWrite'],
      state: 'state',
      challenge: 'challenge',
    });

    expect(url.origin).toBe('https://login.microsoftonline.com');
    expect(url.pathname).toBe('/consumers/oauth2/v2.0/authorize');
    expect(url.searchParams.get('response_type')).toBe('code');
    expect(url.searchParams.get('code_challenge_method')).toBe('S256');
    expect(url.searchParams.get('scope')).toBe('openid https://graph.microsoft.com/Files.ReadWrite');
  });
});
