import { createCipheriv, createDecipheriv, randomBytes } from 'node:crypto';
import { chmod, link, lstat, mkdir, open, readFile, readdir, rename, rm, unlink } from 'node:fs/promises';
import { basename, dirname, join } from 'node:path';
import type { ICachePlugin, TokenCacheContext } from '@azure/msal-node';

const CACHE_AAD = Buffer.from('excel-graph-safe-edit:msal-cache:v1', 'utf8');
const MAX_CACHE_BYTES = 2 * 1024 * 1024;
const KEY_PATTERN = /^[A-Za-z0-9+/]{43}=$/u;
const LOCK_WAIT_MS = 30_000;

interface EncryptedEnvelope {
  version: 1;
  algorithm: 'aes-256-gcm';
  iv: string;
  tag: string;
  ciphertext: string;
}

function isNodeError(error: unknown, code: string): boolean {
  return error instanceof Error && 'code' in error && error.code === code;
}

function sleep(milliseconds: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, milliseconds));
}

async function sameIdentity(path: string, identity: { dev: number; ino: number }): Promise<boolean> {
  try {
    const current = await lstat(path);
    return current.dev === identity.dev && current.ino === identity.ino;
  } catch {
    return false;
  }
}

async function acquireCacheLock(cachePath: string): Promise<() => Promise<void>> {
  const lockPath = `${cachePath}.lock`;
  const recoveryPath = `${lockPath}.recovery`;
  const nonce = randomBytes(12).toString('hex');
  const payload = `${JSON.stringify({ pid: process.pid, nonce })}\n`;
  const deadline = Date.now() + LOCK_WAIT_MS;

  async function createOwnedFile(filePath: string, contents: string) {
    const temporary = `${filePath}.candidate-${process.pid}-${randomBytes(6).toString('hex')}`;
    const candidate = await open(temporary, 'wx+', 0o600);
    try {
      await candidate.writeFile(contents, 'utf8');
      await candidate.sync();
      await candidate.close();
      await link(temporary, filePath);
      await unlink(temporary);
      const handle = await open(filePath, 'r+');
      return { handle, identity: await handle.stat(), payload: contents };
    } catch (error) {
      await candidate.close().catch(() => undefined);
      await unlink(temporary).catch(() => undefined);
      throw error;
    }
  }

  async function readOwner(filePath: string) {
    const handle = await open(filePath, 'r');
    try {
      const identity = await handle.stat();
      if (!identity.isFile() || identity.isSymbolicLink() || identity.size < 1 || identity.size > 256) {
        throw new Error('Encrypted cache lock is invalid');
      }
      const serialized = await handle.readFile('utf8');
      const owner = JSON.parse(serialized) as { pid?: unknown; nonce?: unknown };
      if (!Number.isSafeInteger(owner.pid) || (owner.pid as number) <= 0 ||
          typeof owner.nonce !== 'string' || !/^[a-f0-9]{24}$/u.test(owner.nonce)) {
        throw new Error('Encrypted cache lock is invalid');
      }
      return { identity, owner, payload: serialized };
    } finally {
      await handle.close();
    }
  }

  async function ownerIsDead(pid: number): Promise<boolean> {
    try {
      process.kill(pid, 0);
      return false;
    } catch (error) {
      if (isNodeError(error, 'ESRCH')) return true;
      throw error;
    }
  }

  async function releaseOwnedFile(filePath: string, owned: { identity: { dev: number; ino: number }; payload: string }): Promise<void> {
    if (!await sameIdentity(filePath, owned.identity)) throw new Error('Encrypted cache lock ownership changed');
    const current = await readFile(filePath, 'utf8');
    if (current !== owned.payload || !await sameIdentity(filePath, owned.identity)) {
      throw new Error('Encrypted cache lock ownership changed');
    }
    await unlink(filePath);
  }

  let owned;
  while (!owned) {
    try {
      owned = await createOwnedFile(lockPath, payload);
      break;
    } catch (error) {
      if (!isNodeError(error, 'EEXIST')) throw error;
    }
    let stale = false;
    try {
      const current = await readOwner(lockPath);
      stale = await ownerIsDead(current.owner.pid as number);
    } catch (error) {
      if (!isNodeError(error, 'ENOENT')) {
        if (Date.now() >= deadline) throw error;
        await sleep(100);
        continue;
      }
    }
    if (stale) {
      let guard;
      try {
        guard = await createOwnedFile(recoveryPath, payload);
        const current = await readOwner(lockPath);
        if (await ownerIsDead(current.owner.pid as number) && await sameIdentity(lockPath, current.identity)) {
          await unlink(lockPath);
        }
      } catch (error) {
        if (isNodeError(error, 'EEXIST')) {
          try {
            const existingGuard = await readOwner(recoveryPath);
            if (await ownerIsDead(existingGuard.owner.pid as number) &&
                await sameIdentity(recoveryPath, existingGuard.identity)) {
              await unlink(recoveryPath);
            }
          } catch (guardError) {
            if (!isNodeError(guardError, 'ENOENT') && Date.now() >= deadline) throw guardError;
          }
        } else if (!isNodeError(error, 'ENOENT')) {
          throw error;
        }
      } finally {
        if (guard) {
          await releaseOwnedFile(recoveryPath, guard).catch(() => undefined);
          await guard.handle.close().catch(() => undefined);
        }
      }
    }
    if (Date.now() >= deadline) throw new Error('Timed out waiting for encrypted cache lock');
    await sleep(100);
  }

  return async () => {
    const bytes = Buffer.alloc(Buffer.byteLength(payload));
    try {
      const { bytesRead } = await owned.handle.read(bytes, 0, bytes.length, 0);
      if (bytesRead !== bytes.length || bytes.toString('utf8') !== payload ||
          !await sameIdentity(lockPath, owned.identity)) {
        throw new Error('Encrypted cache lock ownership changed');
      }
      await releaseOwnedFile(lockPath, owned);
    } finally {
      bytes.fill(0);
      await owned.handle.close().catch(() => undefined);
    }
  };
}

async function ensurePrivateDirectory(path: string): Promise<void> {
  await mkdir(path, { recursive: true, mode: 0o700 });
  let stat = await lstat(path);
  if (stat.isSymbolicLink() || !stat.isDirectory() || (stat.mode & 0o777) !== 0o700) {
    if (stat.isSymbolicLink() || !stat.isDirectory()) {
      throw new Error('Encrypted cache directory must be a private regular directory');
    }
    await chmod(path, 0o700);
    stat = await lstat(path);
    if (stat.isSymbolicLink() || !stat.isDirectory() || (stat.mode & 0o777) !== 0o700) {
      throw new Error('Encrypted cache directory must be a private regular directory');
    }
  }
}

async function readCacheKey(keyFile: string): Promise<Buffer> {
  const stat = await lstat(keyFile);
  if (stat.isSymbolicLink() || !stat.isFile() || stat.nlink !== 1 || (stat.mode & 0o777) !== 0o600 || stat.size > 128) {
    throw new Error('Encrypted cache key file is not secure');
  }
  const serialized = (await readFile(keyFile, 'utf8')).trim();
  if (!KEY_PATTERN.test(serialized)) throw new Error('Encrypted cache key must be base64-encoded 32 bytes');
  const key = Buffer.from(serialized, 'base64');
  if (key.length !== 32) throw new Error('Encrypted cache key must decode to 32 bytes');
  return key;
}

function parseEnvelope(raw: string): EncryptedEnvelope {
  const value = JSON.parse(raw) as unknown;
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw new Error('Encrypted cache envelope is invalid');
  const record = value as Record<string, unknown>;
  if (Object.keys(record).sort().join(',') !== 'algorithm,ciphertext,iv,tag,version' ||
      record.version !== 1 || record.algorithm !== 'aes-256-gcm' ||
      typeof record.iv !== 'string' || typeof record.tag !== 'string' || typeof record.ciphertext !== 'string') {
    throw new Error('Encrypted cache envelope is invalid');
  }
  const iv = Buffer.from(record.iv, 'base64');
  const tag = Buffer.from(record.tag, 'base64');
  if (iv.length !== 12 || tag.length !== 16) throw new Error('Encrypted cache envelope is invalid');
  return record as unknown as EncryptedEnvelope;
}

export async function createEncryptedCacheKey(keyFile: string): Promise<void> {
  if (!keyFile.startsWith('/')) throw new Error('Encrypted cache key path must be absolute');
  await ensurePrivateDirectory(dirname(keyFile));
  const key = randomBytes(32);
  const temporary = `${keyFile}.candidate-${process.pid}-${randomBytes(6).toString('hex')}`;
  const handle = await open(temporary, 'wx', 0o600);
  try {
    await handle.writeFile(`${key.toString('base64')}\n`, 'utf8');
    await handle.sync();
    await handle.close();
    await link(temporary, keyFile);
    await unlink(temporary);
  } catch (error) {
    await rm(temporary, { force: true });
    throw error;
  } finally {
    key.fill(0);
    await handle.close().catch(() => undefined);
  }
}

export class EncryptedMsalCachePlugin implements ICachePlugin {
  readonly #cachePath: string;
  readonly #key: Buffer;
  #releaseLock: (() => Promise<void>) | undefined;

  private constructor(cachePath: string, key: Buffer) {
    this.#cachePath = cachePath;
    this.#key = key;
  }

  static async create(cachePath: string, keyFile: string): Promise<EncryptedMsalCachePlugin> {
    if (!cachePath.startsWith('/') || !keyFile.startsWith('/')) {
      throw new Error('Encrypted cache paths must be absolute');
    }
    await ensurePrivateDirectory(dirname(cachePath));
    return new EncryptedMsalCachePlugin(cachePath, await readCacheKey(keyFile));
  }

  async beforeCacheAccess(context: TokenCacheContext): Promise<void> {
    if (this.#releaseLock) throw new Error('Encrypted cache access is already active');
    this.#releaseLock = await acquireCacheLock(this.#cachePath);
    let raw: string;
    try {
      const stat = await lstat(this.#cachePath);
      if (stat.isSymbolicLink() || !stat.isFile() || stat.nlink !== 1 || (stat.mode & 0o777) !== 0o600 || stat.size > MAX_CACHE_BYTES) {
        throw new Error('Encrypted cache file is not secure');
      }
      raw = await readFile(this.#cachePath, 'utf8');
    } catch (error) {
      if (isNodeError(error, 'ENOENT')) {
        context.tokenCache.deserialize('{}');
        return;
      }
      await this.#releaseLock();
      this.#releaseLock = undefined;
      throw error;
    }
    try {
      const envelope = parseEnvelope(raw);
      const iv = Buffer.from(envelope.iv, 'base64');
      const tag = Buffer.from(envelope.tag, 'base64');
      const ciphertext = Buffer.from(envelope.ciphertext, 'base64');
      const decipher = createDecipheriv('aes-256-gcm', this.#key, iv);
      decipher.setAAD(CACHE_AAD);
      decipher.setAuthTag(tag);
      const plaintext = Buffer.concat([decipher.update(ciphertext), decipher.final()]);
      try {
        context.tokenCache.deserialize(plaintext.toString('utf8'));
      } finally {
        plaintext.fill(0);
        ciphertext.fill(0);
        tag.fill(0);
        iv.fill(0);
      }
    } catch (error) {
      await this.#releaseLock();
      this.#releaseLock = undefined;
      throw error;
    }
  }

  async afterCacheAccess(context: TokenCacheContext): Promise<void> {
    if (!this.#releaseLock) throw new Error('Encrypted cache access lock is missing');
    try {
      if (!context.cacheHasChanged) return;
      const plaintext = Buffer.from(context.tokenCache.serialize(), 'utf8');
      const iv = randomBytes(12);
      const cipher = createCipheriv('aes-256-gcm', this.#key, iv);
      cipher.setAAD(CACHE_AAD);
      const ciphertext = Buffer.concat([cipher.update(plaintext), cipher.final()]);
      const tag = cipher.getAuthTag();
      const envelope: EncryptedEnvelope = {
        version: 1,
        algorithm: 'aes-256-gcm',
        iv: iv.toString('base64'),
        tag: tag.toString('base64'),
        ciphertext: ciphertext.toString('base64'),
      };
      const temporary = `${this.#cachePath}.tmp-${process.pid}-${randomBytes(6).toString('hex')}`;
      let handle;
      try {
        handle = await open(temporary, 'wx', 0o600);
        await handle.writeFile(JSON.stringify(envelope), 'utf8');
        await handle.sync();
        await handle.close();
        handle = undefined;
        await rename(temporary, this.#cachePath);
        await chmod(this.#cachePath, 0o600);
      } catch (error) {
        await handle?.close().catch(() => undefined);
        await rm(temporary, { force: true });
        throw error;
      } finally {
        plaintext.fill(0);
        ciphertext.fill(0);
        tag.fill(0);
        iv.fill(0);
      }
    } finally {
      await this.#releaseLock();
      this.#releaseLock = undefined;
    }
  }
}

export async function clearEncryptedMsalCache(cachePath: string): Promise<void> {
  await ensurePrivateDirectory(dirname(cachePath));
  const release = await acquireCacheLock(cachePath);
  try {
    const cacheName = basename(cachePath);
    const names = await readdir(dirname(cachePath));
    await Promise.all(names
      .filter((name) => name === cacheName || name.startsWith(`${cacheName}.tmp-`))
      .map((name) => rm(join(dirname(cachePath), name), { force: true })));
  } finally {
    await release();
  }
}
