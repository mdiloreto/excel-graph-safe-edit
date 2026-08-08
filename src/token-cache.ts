import { randomBytes } from 'node:crypto';
import { chmod, lstat, mkdir, open, readFile, readdir, rename, rm } from 'node:fs/promises';
import { basename, dirname, join } from 'node:path';
import { CACHE_PATH, STATE_DIR } from './config.js';
import { authConfigsMatch, type AuthConfig } from './config.js';

export interface TokenCache {
  clientId: string;
  authority: string;
  scopes: string[];
  accessToken: string;
  refreshToken: string;
  expiresAt: number;
}

function isNodeError(error: unknown, code: string): boolean {
  return error instanceof Error && 'code' in error && error.code === code;
}

async function ensureSecureDirectory(path: string): Promise<void> {
  await mkdir(path, { recursive: true, mode: 0o700 });
  const info = await lstat(path);
  if (info.isSymbolicLink() || !info.isDirectory()) throw new Error(`Refusing unsafe state directory: ${path}`);
  await chmod(path, 0o700);
}

export async function ensureStateDir(): Promise<void> {
  await ensureSecureDirectory(STATE_DIR);
}

function parseTokenCache(raw: string): TokenCache | null {
  let value: unknown;
  try {
    value = JSON.parse(raw) as unknown;
  } catch {
    return null;
  }
  if (!value || typeof value !== 'object' || Array.isArray(value)) return null;
  const record = value as Record<string, unknown>;
  if (
    typeof record.clientId !== 'string'
    || record.clientId.length === 0
    || typeof record.authority !== 'string'
    || record.authority.length === 0
    || !Array.isArray(record.scopes)
    || !record.scopes.every((scope) => typeof scope === 'string')
    || typeof record.accessToken !== 'string'
    || record.accessToken.length === 0
    || typeof record.refreshToken !== 'string'
    || record.refreshToken.length === 0
    || typeof record.expiresAt !== 'number'
    || !Number.isFinite(record.expiresAt)
  ) return null;
  return {
    clientId: record.clientId,
    authority: record.authority,
    scopes: record.scopes,
    accessToken: record.accessToken,
    refreshToken: record.refreshToken,
    expiresAt: record.expiresAt,
  };
}

async function quarantineTokenCache(path: string): Promise<void> {
  const suffix = `${Date.now()}-${randomBytes(4).toString('hex')}`;
  try {
    await rename(path, `${path}.invalid-${suffix}`);
  } catch (error) {
    if (!isNodeError(error, 'ENOENT')) throw error;
  }
}

export async function readTokenCache(
  config: Pick<AuthConfig, 'clientId' | 'authority' | 'scopes'>,
  path = CACHE_PATH,
): Promise<TokenCache | null> {
  await ensureSecureDirectory(dirname(path));
  try {
    const info = await lstat(path);
    if (info.isSymbolicLink() || !info.isFile()) {
      await quarantineTokenCache(path);
      return null;
    }
    await chmod(path, 0o600);
    const raw = await readFile(path, 'utf8');
    const cache = parseTokenCache(raw);
    if (!cache) {
      await quarantineTokenCache(path);
      return null;
    }
    try {
      return authConfigsMatch(cache, config) ? cache : null;
    } catch {
      await quarantineTokenCache(path);
      return null;
    }
  } catch (error) {
    if (isNodeError(error, 'ENOENT')) return null;
    throw error;
  }
}

export async function writeTokenCache(cache: TokenCache, path = CACHE_PATH): Promise<void> {
  const directory = dirname(path);
  await ensureSecureDirectory(directory);
  try {
    const existing = await lstat(path);
    if (existing.isSymbolicLink() || !existing.isFile()) throw new Error(`Refusing unsafe token cache path: ${path}`);
  } catch (error) {
    if (!isNodeError(error, 'ENOENT')) throw error;
  }

  const tempPath = `${path}.tmp-${process.pid}-${randomBytes(6).toString('hex')}`;
  const handle = await open(tempPath, 'wx', 0o600);
  try {
    await handle.writeFile(JSON.stringify(cache, null, 2), 'utf8');
    await handle.chmod(0o600);
    await handle.sync();
    await handle.close();
    await rename(tempPath, path);
    await chmod(path, 0o600);
  } catch (error) {
    await handle.close().catch(() => undefined);
    await rm(tempPath, { force: true }).catch(() => undefined);
    throw error;
  }
}

export async function clearTokenCache(path = CACHE_PATH): Promise<void> {
  const directory = dirname(path);
  await ensureSecureDirectory(directory);
  const cacheName = basename(path);
  const names = await readdir(directory);
  const matchingNames = names.filter((name) => (
    name === cacheName
    || name.startsWith(`${cacheName}.invalid-`)
    || name.startsWith(`${cacheName}.tmp-`)
  ));
  await Promise.all(matchingNames.map((name) => rm(join(directory, name), { force: true, recursive: true })));
}
