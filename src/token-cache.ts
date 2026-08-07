import { chmod, mkdir, readFile, rm, writeFile } from 'node:fs/promises';
import { CACHE_PATH, STATE_DIR } from './config.js';

export interface TokenCache {
  clientId: string;
  authority: string;
  scopes: string[];
  accessToken: string;
  refreshToken: string;
  expiresAt: number;
}

export async function ensureStateDir(): Promise<void> {
  await mkdir(STATE_DIR, { recursive: true, mode: 0o700 });
  await chmod(STATE_DIR, 0o700).catch(() => undefined);
}

export async function readTokenCache(path = CACHE_PATH): Promise<TokenCache | null> {
  try {
    const raw = await readFile(path, 'utf8');
    return JSON.parse(raw) as TokenCache;
  } catch (error) {
    if (error instanceof Error && 'code' in error && error.code === 'ENOENT') return null;
    throw error;
  }
}

export async function writeTokenCache(cache: TokenCache, path = CACHE_PATH): Promise<void> {
  await ensureStateDir();
  await writeFile(path, JSON.stringify(cache, null, 2), { mode: 0o600 });
  await chmod(path, 0o600).catch(() => undefined);
}

export async function clearTokenCache(path = CACHE_PATH): Promise<void> {
  await rm(path, { force: true });
}
