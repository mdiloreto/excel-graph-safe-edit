import {
  InteractionRequiredAuthError,
  PublicClientApplication,
  type AccountInfo,
} from '@azure/msal-node';
import { authConfigKey, type AuthConfig } from './config.js';
import { EncryptedMsalCachePlugin } from './encrypted-msal-cache.js';

const clients = new Map<string, Promise<PublicClientApplication>>();
type MsalClient = Pick<PublicClientApplication, 'getTokenCache' | 'acquireTokenSilent' | 'acquireTokenByDeviceCode'>;

async function createClient(config: AuthConfig): Promise<PublicClientApplication> {
  if (!config.cacheKeyFile || !config.cachePath) throw new Error('Encrypted MSAL cache configuration is incomplete');
  const key = `${authConfigKey(config)}:${config.cachePath}:${config.cacheKeyFile}`;
  const existing = clients.get(key);
  if (existing) return existing;
  const pending = EncryptedMsalCachePlugin.create(config.cachePath, config.cacheKeyFile).then((cachePlugin) => (
    new PublicClientApplication({
      auth: { clientId: config.clientId, authority: config.authority },
      cache: { cachePlugin },
    })
  ));
  clients.set(key, pending);
  return pending;
}

function noninteractiveError(): Error {
  return new Error('Noninteractive authentication failed');
}

async function deviceCodeLogin(client: MsalClient, config: AuthConfig): Promise<string> {
  const result = await client.acquireTokenByDeviceCode({
    scopes: config.scopes,
    deviceCodeCallback: (response) => console.error(response.message),
  });
  if (!result?.accessToken) throw new Error('Microsoft device authorization was cancelled');
  return result.accessToken;
}

export function createEncryptedMsalTokenProvider(overrides: {
  getClient?: (config: AuthConfig) => Promise<MsalClient>;
  interactive?: (client: MsalClient, config: AuthConfig) => Promise<string>;
} = {}): (config: AuthConfig) => Promise<string> {
  const getClient = overrides.getClient ?? createClient;
  const interactive = overrides.interactive ?? deviceCodeLogin;
  const inFlight = new Map<string, Promise<string>>();

  const acquire = async (config: AuthConfig): Promise<string> => {
    try {
      const client = await getClient(config);
      const accounts = await client.getTokenCache().getAllAccounts();
      if (accounts.length > 1) throw new Error('Encrypted MSAL cache contains multiple accounts');
      const account = accounts[0] as AccountInfo | undefined;
      if (!account) {
        if (config.allowInteractive === false) throw noninteractiveError();
        return interactive(client, config);
      }
      try {
        const result = await client.acquireTokenSilent({ account, scopes: config.scopes });
        if (!result.accessToken) throw new Error('Microsoft token response did not include an access token');
        return result.accessToken;
      } catch (error) {
        if (config.allowInteractive === false) throw noninteractiveError();
        if (!(error instanceof InteractionRequiredAuthError)) throw error;
        return interactive(client, config);
      }
    } catch (error) {
      if (config.allowInteractive === false) throw noninteractiveError();
      throw error;
    }
  };

  return (config) => {
    const key = `${authConfigKey(config)}:${config.cachePath}:${config.cacheKeyFile}:${config.allowInteractive === false ? 'noninteractive' : 'interactive'}`;
    const existing = inFlight.get(key);
    if (existing) return existing;
    const pending = acquire(config).finally(() => {
      if (inFlight.get(key) === pending) inFlight.delete(key);
    });
    inFlight.set(key, pending);
    return pending;
  };
}

const defaultEncryptedMsalTokenProvider = createEncryptedMsalTokenProvider();

export function getEncryptedMsalAccessToken(config: AuthConfig): Promise<string> {
  return defaultEncryptedMsalTokenProvider(config);
}
