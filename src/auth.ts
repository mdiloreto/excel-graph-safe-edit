import { spawn } from 'node:child_process';
import { createServer } from 'node:http';
import { type Socket } from 'node:net';
import { authConfigKey, authConfigsMatch, type AuthConfig, CACHE_PATH } from './config.js';
import { buildAuthorizeUrl, createPkcePair, createState } from './oauth.js';
import { readTokenCache, type TokenCache, writeTokenCache } from './token-cache.js';

interface TokenResponse {
  access_token: string;
  refresh_token?: string;
  expires_in?: number;
}

interface CallbackResult {
  code: string;
  redirectUri: string;
}

export interface BrowserLaunchCommand {
  command: string;
  args: string[];
}

export function browserLaunchCommand(platform: NodeJS.Platform, url: string): BrowserLaunchCommand {
  if (platform === 'darwin') return { command: 'open', args: [url] };
  if (platform === 'win32') return { command: 'rundll32.exe', args: ['url.dll,FileProtocolHandler', url] };
  return { command: 'xdg-open', args: [url] };
}

async function openBrowser(url: string): Promise<void> {
  const launch = browserLaunchCommand(process.platform, url);
  await new Promise<void>((resolve, reject) => {
    const child = spawn(launch.command, launch.args, { detached: true, shell: false, stdio: 'ignore' });
    let settled = false;
    child.once('error', (error) => {
      if (!settled) {
        settled = true;
        reject(error);
      }
    });
    child.once('spawn', () => child.unref());
    child.once('exit', (code, signal) => {
      if (settled) return;
      settled = true;
      if (code === 0) resolve();
      else reject(new Error(`Browser opener exited unsuccessfully (${signal ?? code ?? 'unknown'})`));
    });
  });
}

function oauthErrorCode(value: unknown): string {
  return typeof value === 'string' && /^[A-Za-z0-9_.-]{1,100}$/.test(value) ? value : 'unknown_error';
}

async function tokenRequest(config: AuthConfig, body: URLSearchParams): Promise<TokenResponse> {
  const response = await fetch(`${config.authority}/oauth2/v2.0/token`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body,
  });
  let payload: unknown;
  try {
    payload = await response.json() as unknown;
  } catch {
    throw new Error(`Token request failed with an invalid response (${response.status})`);
  }
  const record = payload && typeof payload === 'object' && !Array.isArray(payload)
    ? payload as Record<string, unknown>
    : {};
  if (!response.ok) throw new Error(`Token request failed: ${oauthErrorCode(record.error)} (${response.status})`);
  if (typeof record.access_token !== 'string') throw new Error('Token response did not include an access token');
  return {
    access_token: record.access_token,
    refresh_token: typeof record.refresh_token === 'string' ? record.refresh_token : undefined,
    expires_in: typeof record.expires_in === 'number' && Number.isFinite(record.expires_in) ? record.expires_in : undefined,
  };
}

async function exchangeCode(config: AuthConfig, code: string, verifier: string, redirectUri: string): Promise<TokenResponse> {
  return tokenRequest(config, new URLSearchParams({
    client_id: config.clientId,
    scope: config.scopes.join(' '),
    code,
    redirect_uri: redirectUri,
    grant_type: 'authorization_code',
    code_verifier: verifier,
  }));
}

async function refreshAccessToken(config: AuthConfig, refreshToken: string): Promise<TokenResponse> {
  return tokenRequest(config, new URLSearchParams({
    client_id: config.clientId,
    scope: config.scopes.join(' '),
    refresh_token: refreshToken,
    grant_type: 'refresh_token',
  }));
}

export function waitForOAuthCallback(input: {
  expectedState: string;
  port: number;
  timeoutMs: number;
  onListening: (redirectUri: string) => void | Promise<void>;
}): Promise<CallbackResult> {
  return new Promise<CallbackResult>((resolve, reject) => {
    let settled = false;
    let timer: NodeJS.Timeout | undefined;
    const sockets = new Set<Socket>();
    const cleanup = (): void => {
      if (server.listening) server.close();
      for (const socket of sockets) socket.destroy();
      sockets.clear();
    };
    const settle = (error: Error | null, result?: CallbackResult): void => {
      if (settled) return;
      settled = true;
      if (timer) clearTimeout(timer);
      cleanup();
      if (error) reject(error);
      else if (result) resolve(result);
    };
    const server = createServer((request, response) => {
      try {
        const url = new URL(request.url ?? '/', 'http://127.0.0.1');
        if (url.pathname !== '/callback') {
          response.writeHead(404, { Connection: 'close', 'Content-Type': 'text/plain; charset=utf-8' }).end('Not found');
          return;
        }

        const settleAfterResponse = (error: Error | null, result?: CallbackResult): void => {
          let responseCompleted = false;
          const complete = (): void => {
            if (responseCompleted) return;
            responseCompleted = true;
            setImmediate(() => settle(error, result));
          };
          response.once('finish', complete);
          response.once('close', complete);
        };

        if (url.searchParams.get('state') !== input.expectedState) {
          response.writeHead(400, { Connection: 'close', 'Content-Type': 'text/plain; charset=utf-8' }).end('Authentication state did not match. You can close this tab.');
          return;
        }
        const error = url.searchParams.get('error');
        if (error) {
          settleAfterResponse(new Error(`OAuth authorization failed: ${oauthErrorCode(error)}`));
          response.writeHead(400, { Connection: 'close', 'Content-Type': 'text/plain; charset=utf-8' }).end('Authentication failed. You can close this tab.');
          return;
        }
        const code = url.searchParams.get('code');
        if (!code) {
          settleAfterResponse(new Error('Authentication callback did not include a code'));
          response.writeHead(400, { Connection: 'close', 'Content-Type': 'text/plain; charset=utf-8' }).end('Authentication code was missing. You can close this tab.');
          return;
        }
        const address = server.address();
        if (!address || typeof address === 'string') {
          settleAfterResponse(new Error('OAuth callback server address was unavailable'));
          response.writeHead(500, { Connection: 'close', 'Content-Type': 'text/plain; charset=utf-8' }).end('Authentication callback failed.');
          return;
        }
        settleAfterResponse(null, { code, redirectUri: `http://localhost:${address.port}/callback` });
        response.writeHead(200, {
          Connection: 'close',
          'Content-Security-Policy': "default-src 'none'",
          'Content-Type': 'text/html; charset=utf-8',
        }).end('<!doctype html><title>Excel Graph Login</title><p>Login complete. You can close this tab.</p>');
      } catch (error) {
        settle(error instanceof Error ? error : new Error('OAuth callback failed'));
      }
    });

    server.on('connection', (socket) => {
      sockets.add(socket);
      socket.once('close', () => sockets.delete(socket));
    });
    server.once('error', (error) => settle(error));
    server.listen(input.port, '127.0.0.1', () => {
      const address = server.address();
      if (!address || typeof address === 'string') {
        settle(new Error('OAuth callback server did not expose a TCP port'));
        return;
      }
      const redirectUri = `http://localhost:${address.port}/callback`;
      timer = setTimeout(() => settle(new Error('Timed out waiting for the OAuth callback')), input.timeoutMs);
      timer.unref();
      void Promise.resolve(input.onListening(redirectUri)).catch((error: unknown) => {
        settle(error instanceof Error ? error : new Error('Failed to start browser login'));
      });
    });
  });
}

async function acquireInteractiveToken(config: AuthConfig): Promise<TokenResponse> {
  const { verifier, challenge } = createPkcePair();
  const expectedState = createState();
  const callback = await waitForOAuthCallback({
    expectedState,
    port: config.port,
    timeoutMs: 5 * 60 * 1000,
    onListening: async (redirectUri) => {
      const authorizeUrl = buildAuthorizeUrl({
        authority: config.authority,
        clientId: config.clientId,
        redirectUri,
        scopes: config.scopes,
        state: expectedState,
        challenge,
      }).toString();
      try {
        await openBrowser(authorizeUrl);
        console.error('Opened a browser for Microsoft login.');
      } catch {
        console.error(`Could not open a browser. Open this one-time Microsoft authorization URL manually:\n${authorizeUrl}`);
      }
    },
  });
  return exchangeCode(config, callback.code, verifier, callback.redirectUri);
}

interface TokenProviderDependencies {
  now: () => number;
  platform: () => NodeJS.Platform;
  readCache: (config: AuthConfig) => Promise<TokenCache | null>;
  writeCache: (cache: TokenCache) => Promise<void>;
  refresh: (config: AuthConfig, refreshToken: string) => Promise<TokenResponse>;
  interactive: (config: AuthConfig) => Promise<TokenResponse>;
}

export function assertPersistentTokenStorageSupported(persist: boolean, platform: NodeJS.Platform = process.platform): void {
  if (persist && platform === 'win32') {
    throw new Error('Persistent token storage is not supported securely on Windows; use --no-persist.');
  }
}

export function createAccessTokenProvider(overrides: Partial<TokenProviderDependencies> = {}): (config: AuthConfig) => Promise<string> {
  const dependencies: TokenProviderDependencies = {
    now: () => Math.floor(Date.now() / 1000),
    platform: () => process.platform,
    readCache: (config) => readTokenCache(config),
    writeCache: writeTokenCache,
    refresh: refreshAccessToken,
    interactive: acquireInteractiveToken,
    ...overrides,
  };
  const inMemoryTokens = new Map<string, { accessToken: string; expiresAt: number }>();
  const inFlight = new Map<string, Promise<string>>();

  const acquire = async (config: AuthConfig, key: string): Promise<string> => {
    const now = dependencies.now();
    const memory = inMemoryTokens.get(key);
    if (memory && memory.expiresAt > now + 120) return memory.accessToken;

    if (config.persist) {
      const cache = await dependencies.readCache(config);
      if (cache && authConfigsMatch(cache, config)) {
        if (cache.expiresAt > now + 120) {
          inMemoryTokens.set(key, { accessToken: cache.accessToken, expiresAt: cache.expiresAt });
          return cache.accessToken;
        }
        if (cache.refreshToken) {
          try {
            const token = await dependencies.refresh(config, cache.refreshToken);
            const nextCache: TokenCache = {
              clientId: config.clientId,
              authority: config.authority,
              scopes: config.scopes,
              accessToken: token.access_token,
              refreshToken: token.refresh_token ?? cache.refreshToken,
              expiresAt: dependencies.now() + (token.expires_in ?? 3600),
            };
            await dependencies.writeCache(nextCache);
            inMemoryTokens.set(key, { accessToken: nextCache.accessToken, expiresAt: nextCache.expiresAt });
            return nextCache.accessToken;
          } catch {
            console.error('Cached refresh failed; falling back to browser login.');
          }
        }
      }
    }

    const token = await dependencies.interactive(config);
    const expiresAt = dependencies.now() + (token.expires_in ?? 3600);
    if (config.persist) {
      if (!token.refresh_token) throw new Error('No refresh token returned. Ensure offline_access is included in scopes.');
      await dependencies.writeCache({
        clientId: config.clientId,
        authority: config.authority,
        scopes: config.scopes,
        accessToken: token.access_token,
        refreshToken: token.refresh_token,
        expiresAt,
      });
    }
    inMemoryTokens.set(key, { accessToken: token.access_token, expiresAt });
    return token.access_token;
  };

  return (config: AuthConfig): Promise<string> => {
    try {
      assertPersistentTokenStorageSupported(config.persist, dependencies.platform());
    } catch (error) {
      return Promise.reject(error);
    }
    const key = `${authConfigKey(config)}:${config.persist ? 'persist' : 'no-persist'}`;
    const existing = inFlight.get(key);
    if (existing) return existing;
    const pending = acquire(config, key).finally(() => {
      if (inFlight.get(key) === pending) inFlight.delete(key);
    });
    inFlight.set(key, pending);
    return pending;
  };
}

const defaultAccessTokenProvider = createAccessTokenProvider();

export function getAccessToken(config: AuthConfig): Promise<string> {
  return defaultAccessTokenProvider(config);
}

export { CACHE_PATH };
