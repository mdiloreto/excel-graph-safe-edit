import { createServer } from 'node:http';
import { spawn } from 'node:child_process';
import { AuthConfig, CACHE_PATH } from './config.js';
import { buildAuthorizeUrl, createPkcePair, createState } from './oauth.js';
import { readTokenCache, TokenCache, writeTokenCache } from './token-cache.js';

interface TokenResponse {
  access_token: string;
  refresh_token?: string;
  expires_in?: number;
}

let inMemoryToken: { accessToken: string; expiresAt: number } | null = null;

function openBrowser(url: string): void {
  const command = process.platform === 'darwin' ? 'open' : process.platform === 'win32' ? 'cmd' : 'xdg-open';
  const args = process.platform === 'win32' ? ['/c', 'start', '', url] : [url];
  const child = spawn(command, args, { detached: true, stdio: 'ignore' });
  child.on('error', () => undefined);
  child.unref();
}

async function tokenRequest(config: AuthConfig, body: URLSearchParams): Promise<TokenResponse> {
  const response = await fetch(`${config.authority}/oauth2/v2.0/token`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body,
  });
  const payload = (await response.json()) as Record<string, unknown>;
  if (!response.ok) {
    const code = typeof payload.error === 'string' ? payload.error : response.status;
    const description = typeof payload.error_description === 'string' ? payload.error_description : '';
    throw new Error(`Token request failed: ${code} ${description}`.trim());
  }
  if (typeof payload.access_token !== 'string') throw new Error('Token response did not include an access token');
  return {
    access_token: payload.access_token,
    refresh_token: typeof payload.refresh_token === 'string' ? payload.refresh_token : undefined,
    expires_in: typeof payload.expires_in === 'number' ? payload.expires_in : undefined,
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

async function acquireInteractiveToken(config: AuthConfig): Promise<TokenResponse> {
  const { verifier, challenge } = createPkcePair();
  const expectedState = createState();
  let actualPort = config.port;

  const callback = await new Promise<{ code: string; state: string | null }>((resolve, reject) => {
    const server = createServer((request, response) => {
      try {
        const url = new URL(request.url ?? '/', `http://${request.headers.host}`);
        if (url.pathname !== '/callback') {
          response.writeHead(404).end('Not found');
          return;
        }
        const error = url.searchParams.get('error');
        if (error) {
          response.writeHead(400, { 'Content-Type': 'text/plain' }).end('Authentication failed. You can close this tab.');
          reject(new Error(`${error}: ${url.searchParams.get('error_description') ?? 'no description'}`));
          server.close();
          return;
        }
        const code = url.searchParams.get('code');
        if (!code) {
          reject(new Error('Authentication callback did not include a code'));
          server.close();
          return;
        }
        response.writeHead(200, { 'Content-Type': 'text/html' }).end('<!doctype html><title>Excel Graph Login</title><p>Login complete. You can close this tab.</p>');
        resolve({ code, state: url.searchParams.get('state') });
        server.close();
      } catch (error) {
        reject(error);
        server.close();
      }
    });

    server.on('error', reject);
    server.listen(config.port, '127.0.0.1', () => {
      const address = server.address();
      if (typeof address === 'object' && address) actualPort = address.port;
      const redirectUri = `http://localhost:${actualPort}/callback`;
      const url = buildAuthorizeUrl({
        authority: config.authority,
        clientId: config.clientId,
        redirectUri,
        scopes: config.scopes,
        state: expectedState,
        challenge,
      });
      console.error(`Opening browser for Microsoft login: ${url.toString()}`);
      openBrowser(url.toString());
    });
  });

  if (callback.state !== expectedState) throw new Error('OAuth state mismatch');
  return exchangeCode(config, callback.code, verifier, `http://localhost:${actualPort}/callback`);
}

export async function getAccessToken(config: AuthConfig): Promise<string> {
  const now = Math.floor(Date.now() / 1000);
  if (inMemoryToken && inMemoryToken.expiresAt > now + 120) return inMemoryToken.accessToken;

  if (config.persist) {
    const cache = await readTokenCache();
    if (cache?.accessToken && cache.expiresAt > now + 120) {
      inMemoryToken = { accessToken: cache.accessToken, expiresAt: cache.expiresAt };
      return cache.accessToken;
    }
    if (cache?.refreshToken) {
      try {
        const token = await refreshAccessToken(config, cache.refreshToken);
        const nextCache: TokenCache = {
          clientId: config.clientId,
          authority: config.authority,
          scopes: config.scopes,
          accessToken: token.access_token,
          refreshToken: token.refresh_token ?? cache.refreshToken,
          expiresAt: now + (token.expires_in ?? 3600),
        };
        await writeTokenCache(nextCache);
        inMemoryToken = { accessToken: nextCache.accessToken, expiresAt: nextCache.expiresAt };
        return nextCache.accessToken;
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        console.error(`Cached refresh failed; falling back to browser login: ${message}`);
      }
    }
  }

  const token = await acquireInteractiveToken(config);
  const expiresAt = now + (token.expires_in ?? 3600);
  inMemoryToken = { accessToken: token.access_token, expiresAt };
  if (config.persist) {
    if (!token.refresh_token) throw new Error('No refresh token returned. Ensure offline_access is included in scopes.');
    await writeTokenCache({
      clientId: config.clientId,
      authority: config.authority,
      scopes: config.scopes,
      accessToken: token.access_token,
      refreshToken: token.refresh_token,
      expiresAt,
    });
  }
  return token.access_token;
}

export { CACHE_PATH };
