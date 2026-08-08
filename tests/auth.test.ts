import { chmod, mkdtemp, readFile, readdir, rm, stat, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { createConnection, type Socket } from 'node:net';
import { afterEach, describe, expect, it, vi } from 'vitest';
import {
  assertPersistentTokenStorageSupported,
  browserLaunchCommand,
  createAccessTokenProvider,
  waitForOAuthCallback,
} from '../src/auth.js';
import { type AuthConfig } from '../src/config.js';
import { clearTokenCache, readTokenCache, writeTokenCache, type TokenCache } from '../src/token-cache.js';

const temporaryDirectories: string[] = [];

afterEach(async () => {
  vi.restoreAllMocks();
  await Promise.all(temporaryDirectories.splice(0).map((path) => rm(path, { recursive: true, force: true })));
});

function config(overrides: Partial<AuthConfig> = {}): AuthConfig {
  return {
    clientId: 'client-a',
    authority: 'https://login.microsoftonline.com/tenant',
    scopes: ['openid', 'scope-a'],
    persist: false,
    port: 0,
    ...overrides,
  };
}

function interactiveToken(accessToken: string) {
  return { access_token: accessToken, refresh_token: `refresh-${accessToken}`, expires_in: 3600 };
}

describe('access token isolation', () => {
  it('fails closed for persistent Windows auth but permits one-shot mode', async () => {
    expect(() => assertPersistentTokenStorageSupported(true, 'win32')).toThrow(/--no-persist/);
    expect(() => assertPersistentTokenStorageSupported(false, 'win32')).not.toThrow();
    const interactive = vi.fn(async () => interactiveToken('one-shot'));
    const getAccessToken = createAccessTokenProvider({ interactive, platform: () => 'win32' });

    await expect(getAccessToken(config({ persist: true }))).rejects.toThrow(/--no-persist/);
    expect(interactive).not.toHaveBeenCalled();
    await expect(getAccessToken(config({ persist: false }))).resolves.toBe('one-shot');
  });

  it('reuses memory only for matching normalized auth identities', async () => {
    const interactive = vi.fn(async (authConfig: AuthConfig) => interactiveToken(authConfig.clientId));
    const getAccessToken = createAccessTokenProvider({ interactive });

    await expect(getAccessToken(config())).resolves.toBe('client-a');
    await expect(getAccessToken(config({ authority: 'https://login.microsoftonline.com/tenant/', scopes: ['scope-a', 'openid'] }))).resolves.toBe('client-a');
    await expect(getAccessToken(config({ clientId: 'client-b' }))).resolves.toBe('client-b');
    await expect(getAccessToken(config({ scopes: ['scope-b'] }))).resolves.toBe('client-a');

    expect(interactive).toHaveBeenCalledTimes(3);
  });

  it('does not use a cache entry for another auth identity', async () => {
    const cached: TokenCache = {
      clientId: 'other-client',
      authority: 'https://login.microsoftonline.com/tenant',
      scopes: ['openid', 'scope-a'],
      accessToken: 'cached-secret',
      refreshToken: 'cached-refresh',
      expiresAt: 10_000,
    };
    const interactive = vi.fn(async () => interactiveToken('interactive-token'));
    const refresh = vi.fn(async () => interactiveToken('refreshed-token'));
    const getAccessToken = createAccessTokenProvider({
      now: () => 100,
      readCache: async () => cached,
      writeCache: async () => undefined,
      refresh,
      interactive,
    });

    await expect(getAccessToken(config({ persist: true }))).resolves.toBe('interactive-token');
    expect(refresh).not.toHaveBeenCalled();
    expect(interactive).toHaveBeenCalledOnce();
  });

  it('deduplicates concurrent acquisition and clears failed pending work', async () => {
    let now = 0;
    let release: ((value: ReturnType<typeof interactiveToken>) => void) | undefined;
    const interactive = vi.fn(() => new Promise<ReturnType<typeof interactiveToken>>((resolve) => {
      release = resolve;
    }));
    const getAccessToken = createAccessTokenProvider({ interactive, now: () => now });
    const first = getAccessToken(config());
    const second = getAccessToken(config());
    expect(interactive).toHaveBeenCalledOnce();
    release?.(interactiveToken('shared-token'));
    await expect(Promise.all([first, second])).resolves.toEqual(['shared-token', 'shared-token']);
    now = 4000;
    const afterSuccess = getAccessToken(config());
    expect(interactive).toHaveBeenCalledTimes(2);
    release?.(interactiveToken('new-token'));
    await expect(afterSuccess).resolves.toBe('new-token');

    const failingInteractive = vi.fn()
      .mockRejectedValueOnce(new Error('browser failed'))
      .mockResolvedValueOnce(interactiveToken('retry-token'));
    const retryingProvider = createAccessTokenProvider({ interactive: failingInteractive });
    await expect(retryingProvider(config())).rejects.toThrow('browser failed');
    await expect(retryingProvider(config())).resolves.toBe('retry-token');
    expect(failingInteractive).toHaveBeenCalledTimes(2);
  });

  it('does not coalesce persistent cache work into a no-persist acquisition', async () => {
    const readCache = vi.fn(async () => null);
    const writeCache = vi.fn(async () => undefined);
    const interactive = vi.fn(async (authConfig: AuthConfig) => interactiveToken(authConfig.persist ? 'persistent' : 'one-shot'));
    const getAccessToken = createAccessTokenProvider({ readCache, writeCache, interactive });

    await expect(Promise.all([
      getAccessToken(config({ persist: true })),
      getAccessToken(config({ persist: false })),
    ])).resolves.toEqual(['persistent', 'one-shot']);
    expect(readCache).toHaveBeenCalledOnce();
    expect(writeCache).toHaveBeenCalledOnce();
    expect(interactive).toHaveBeenCalledTimes(2);
  });

  it('does not reuse a sequential no-persist token for persistent acquisition', async () => {
    const readCache = vi.fn(async () => null);
    const writeCache = vi.fn(async () => undefined);
    const interactive = vi.fn(async (authConfig: AuthConfig) => interactiveToken(authConfig.persist ? 'persistent' : 'one-shot'));
    const getAccessToken = createAccessTokenProvider({ readCache, writeCache, interactive });
    const sharedScopes = ['openid', 'scope-a'];

    await expect(getAccessToken(config({ persist: false, scopes: sharedScopes }))).resolves.toBe('one-shot');
    await expect(getAccessToken(config({ persist: true, scopes: sharedScopes }))).resolves.toBe('persistent');
    expect(readCache).toHaveBeenCalledOnce();
    expect(writeCache).toHaveBeenCalledOnce();
    expect(interactive).toHaveBeenCalledTimes(2);
  });

  it('does not read or write the persistent cache in no-persist mode', async () => {
    const readCache = vi.fn(async () => null);
    const writeCache = vi.fn(async () => undefined);
    const getAccessToken = createAccessTokenProvider({
      readCache,
      writeCache,
      interactive: async () => interactiveToken('one-shot'),
    });

    await expect(getAccessToken(config({ persist: false }))).resolves.toBe('one-shot');
    expect(readCache).not.toHaveBeenCalled();
    expect(writeCache).not.toHaveBeenCalled();
  });
});

describe('token cache validation', () => {
  it('accepts only cache entries matching the normalized config', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'excel-graph-auth-'));
    temporaryDirectories.push(directory);
    const path = join(directory, 'cache.json');
    const authConfig = config({ persist: true });
    await writeTokenCache({
      clientId: authConfig.clientId,
      authority: `${authConfig.authority}/`,
      scopes: [...authConfig.scopes].reverse(),
      accessToken: 'access-token',
      refreshToken: 'refresh-token',
      expiresAt: 5000,
    }, path);

    await expect(readTokenCache(authConfig, path)).resolves.toMatchObject({ accessToken: 'access-token' });
    await expect(readTokenCache(config({ clientId: 'different' }), path)).resolves.toBeNull();
  });

  it('quarantines malformed cache data without exposing it', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'excel-graph-auth-'));
    temporaryDirectories.push(directory);
    const path = join(directory, 'cache.json');
    await writeFile(path, '{"accessToken":"secret","expiresAt":"invalid"}', { mode: 0o600 });

    await expect(readTokenCache(config(), path)).resolves.toBeNull();
    await expect(readFile(path, 'utf8')).rejects.toMatchObject({ code: 'ENOENT' });
    expect((await readdir(directory)).some((name) => name.startsWith('cache.json.invalid-'))).toBe(true);
  });

  it('hardens the cache parent directory before reading', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'excel-graph-auth-'));
    temporaryDirectories.push(directory);
    const path = join(directory, 'cache.json');
    const authConfig = config({ persist: true });
    await writeTokenCache({
      clientId: authConfig.clientId,
      authority: authConfig.authority,
      scopes: authConfig.scopes,
      accessToken: 'access-token',
      refreshToken: 'refresh-token',
      expiresAt: 5000,
    }, path);
    await chmod(directory, 0o755);

    await expect(readTokenCache(authConfig, path)).resolves.toMatchObject({ accessToken: 'access-token' });
    expect((await stat(directory)).mode & 0o777).toBe(0o700);
  });

  it('clears active, quarantined, and abandoned cache artifacts', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'excel-graph-auth-'));
    temporaryDirectories.push(directory);
    const path = join(directory, 'cache.json');
    const matchingNames = ['cache.json', 'cache.json.invalid-123', 'cache.json.tmp-456'];
    await Promise.all(matchingNames.map((name) => writeFile(join(directory, name), 'refresh-token-secret', { mode: 0o600 })));
    await writeFile(join(directory, 'unrelated.txt'), 'keep', { mode: 0o600 });

    await clearTokenCache(path);

    const remaining = await readdir(directory);
    expect(remaining.filter((name) => name.startsWith('cache.json'))).toEqual([]);
    expect(remaining).toContain('unrelated.txt');
  });
});

describe('browser callback safety', () => {
  it('uses a shell-free Windows browser command', () => {
    const launch = browserLaunchCommand('win32', 'https://example.test/?value=a&b=c');
    expect(launch.command).toBe('rundll32.exe');
    expect(launch.args).toEqual(['url.dll,FileProtocolHandler', 'https://example.test/?value=a&b=c']);
  });

  it('rejects wrong state but keeps listening for the legitimate callback', async () => {
    let wrongStateStatus = 0;
    let wrongStateBody = '';
    let request: Promise<void> | undefined;
    const callback = waitForOAuthCallback({
      expectedState: 'expected',
      port: 0,
      timeoutMs: 1000,
      onListening: (redirectUri) => {
        request = (async () => {
          const wrongStateResponse = await fetch(`${redirectUri}?code=attacker-code&state=wrong`);
          wrongStateStatus = wrongStateResponse.status;
          wrongStateBody = await wrongStateResponse.text();
          const legitimateResponse = await fetch(`${redirectUri}?code=legitimate-code&state=expected`);
          expect(legitimateResponse.status).toBe(200);
          expect(await legitimateResponse.text()).toContain('Login complete');
        })();
      },
    });

    await expect(callback).resolves.toMatchObject({ code: 'legitimate-code' });
    await request;
    expect(wrongStateStatus).toBe(400);
    expect(wrongStateBody).not.toContain('Login complete');
  });

  it('accepts a callback only when state matches', async () => {
    let request: Promise<void> | undefined;
    const callback = waitForOAuthCallback({
      expectedState: 'expected',
      port: 0,
      timeoutMs: 1000,
      onListening: (redirectUri) => {
        request = fetch(`${redirectUri}?code=authorization-code&state=expected`).then(async (response) => {
          expect(response.status).toBe(200);
          expect(await response.text()).toContain('Login complete');
        });
      },
    });

    await expect(callback).resolves.toMatchObject({ code: 'authorization-code' });
    await request;
  });

  it('times out and closes the callback server', async () => {
    let redirectUri = '';
    const callback = waitForOAuthCallback({
      expectedState: 'expected',
      port: 0,
      timeoutMs: 20,
      onListening: (value) => {
        redirectUri = value;
      },
    });

    await expect(callback).rejects.toThrow('Timed out');
    await expect(fetch(redirectUri)).rejects.toThrow();
  });

  it('destroys a slow incomplete callback socket on timeout', async () => {
    let socket: Socket | undefined;
    let resolveClosed: (() => void) | undefined;
    const closed = new Promise<void>((resolve) => {
      resolveClosed = resolve;
    });
    const callback = waitForOAuthCallback({
      expectedState: 'expected',
      port: 0,
      timeoutMs: 50,
      onListening: (redirectUri) => {
        const port = Number(new URL(redirectUri).port);
        socket = createConnection({ host: '127.0.0.1', port });
        socket.once('error', () => undefined);
        socket.once('close', () => resolveClosed?.());
        socket.once('connect', () => socket?.write('GET /callback HTTP/1.1\r\nHost:'));
      },
    });

    await expect(callback).rejects.toThrow('Timed out');
    await closed;
    expect(socket?.destroyed).toBe(true);
  });
});
