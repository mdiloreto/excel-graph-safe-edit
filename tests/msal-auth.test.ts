import type { AccountInfo, PublicClientApplication } from '@azure/msal-node';
import { describe, expect, it, vi } from 'vitest';
import { buildAuthConfig, type AuthConfig } from '../src/config.js';
import { createEncryptedMsalTokenProvider } from '../src/msal-auth.js';

type MsalClient = Pick<PublicClientApplication, 'getTokenCache' | 'acquireTokenSilent' | 'acquireTokenByDeviceCode'>;

function config(overrides: Partial<AuthConfig> = {}): AuthConfig {
  return {
    ...buildAuthConfig({
      clientId: 'client-id',
      cacheKeyFile: '/private/material-archive.key',
      cachePath: '/private/material-archive.enc',
    }),
    ...overrides,
  };
}

function client(accounts: AccountInfo[] = []): MsalClient {
  return {
    getTokenCache: () => ({ getAllAccounts: async () => accounts }) as ReturnType<PublicClientApplication['getTokenCache']>,
    acquireTokenSilent: vi.fn(),
    acquireTokenByDeviceCode: vi.fn(),
  };
}

describe('encrypted MSAL token provider', () => {
  it('deduplicates concurrent device-code acquisition', async () => {
    const fakeClient = client();
    let resolveLogin: ((token: string) => void) | undefined;
    const interactive = vi.fn(() => new Promise<string>((resolve) => { resolveLogin = resolve; }));
    const getToken = createEncryptedMsalTokenProvider({
      getClient: async () => fakeClient,
      interactive,
    });
    const first = getToken(config());
    const second = getToken(config());
    await vi.waitFor(() => expect(interactive).toHaveBeenCalledOnce());
    resolveLogin?.('access-token');
    await expect(Promise.all([first, second])).resolves.toEqual(['access-token', 'access-token']);
  });

  it('fails closed when an encrypted cache contains multiple accounts', async () => {
    const accounts = [
      { homeAccountId: 'one' } as AccountInfo,
      { homeAccountId: 'two' } as AccountInfo,
    ];
    const getToken = createEncryptedMsalTokenProvider({ getClient: async () => client(accounts) });
    await expect(getToken(config())).rejects.toThrow(/multiple accounts/u);
  });

  it('hides cache setup failures in noninteractive mode', async () => {
    const getToken = createEncryptedMsalTokenProvider({
      getClient: async () => { throw new Error('/private/material-archive.key is missing'); },
    });
    await expect(getToken(config({ allowInteractive: false }))).rejects.toThrow(
      /^Noninteractive authentication failed$/u,
    );
  });
});
