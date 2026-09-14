import { lstat, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import type { TokenCacheContext } from '@azure/msal-node';
import { afterEach, describe, expect, it, vi } from 'vitest';
import {
  clearEncryptedMsalCache,
  createEncryptedCacheKey,
  EncryptedMsalCachePlugin,
} from '../src/encrypted-msal-cache.js';

const temporaryDirectories: string[] = [];

afterEach(async () => {
  await Promise.all(temporaryDirectories.splice(0).map((path) => rm(path, { recursive: true, force: true })));
});

function cacheContext(input: { serialized?: string; changed: boolean; deserialize?: (value: string) => void }): TokenCacheContext {
  return {
    cacheHasChanged: input.changed,
    tokenCache: {
      serialize: () => input.serialized ?? '',
      deserialize: input.deserialize ?? (() => undefined),
    },
  } as unknown as TokenCacheContext;
}

describe('encrypted MSAL cache', () => {
  it('creates a private independent key without overwriting it', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'excel-graph-msal-'));
    temporaryDirectories.push(directory);
    const keyPath = join(directory, 'material-archive.key');
    await createEncryptedCacheKey(keyPath);
    const serialized = (await readFile(keyPath, 'utf8')).trim();
    expect(serialized).toMatch(/^[A-Za-z0-9+/]{43}=$/u);
    expect((await lstat(keyPath)).mode & 0o777).toBe(0o600);
    await expect(createEncryptedCacheKey(keyPath)).rejects.toMatchObject({ code: 'EEXIST' });
  });

  it('round-trips MSAL state without persisting plaintext tokens', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'excel-graph-msal-'));
    temporaryDirectories.push(directory);
    const keyPath = join(directory, 'material-archive.key');
    const cachePath = join(directory, 'material-archive.enc');
    await createEncryptedCacheKey(keyPath);
    const writer = await EncryptedMsalCachePlugin.create(cachePath, keyPath);
    const secret = JSON.stringify({ accessToken: 'access-secret', refreshToken: 'refresh-secret' });
    await writer.beforeCacheAccess(cacheContext({ changed: false }));
    await writer.afterCacheAccess(cacheContext({ serialized: secret, changed: true }));

    const encrypted = await readFile(cachePath, 'utf8');
    expect(encrypted).not.toContain('access-secret');
    expect(encrypted).not.toContain('refresh-secret');
    expect((await lstat(cachePath)).mode & 0o777).toBe(0o600);

    const deserialize = vi.fn();
    const reader = await EncryptedMsalCachePlugin.create(cachePath, keyPath);
    await reader.beforeCacheAccess(cacheContext({ changed: false, deserialize }));
    expect(deserialize).toHaveBeenCalledWith(secret);
  });

  it('rejects tampering and clears only owned cache artifacts', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'excel-graph-msal-'));
    temporaryDirectories.push(directory);
    const keyPath = join(directory, 'material-archive.key');
    const cachePath = join(directory, 'material-archive.enc');
    await createEncryptedCacheKey(keyPath);
    const plugin = await EncryptedMsalCachePlugin.create(cachePath, keyPath);
    await plugin.beforeCacheAccess(cacheContext({ changed: false }));
    await plugin.afterCacheAccess(cacheContext({ serialized: '{"secret":"value"}', changed: true }));
    const envelope = JSON.parse(await readFile(cachePath, 'utf8')) as { ciphertext: string };
    envelope.ciphertext = `${envelope.ciphertext.slice(0, -2)}AA`;
    await writeFile(cachePath, JSON.stringify(envelope), { mode: 0o600 });
    await expect(plugin.beforeCacheAccess(cacheContext({ changed: false }))).rejects.toThrow();

    await writeFile(`${cachePath}.tmp-abandoned`, 'secret', { mode: 0o600 });
    await writeFile(join(directory, 'unrelated.txt'), 'keep', { mode: 0o600 });
    await clearEncryptedMsalCache(cachePath);
    await expect(readFile(cachePath)).rejects.toMatchObject({ code: 'ENOENT' });
    await expect(readFile(`${cachePath}.tmp-abandoned`)).rejects.toMatchObject({ code: 'ENOENT' });
    await expect(readFile(join(directory, 'unrelated.txt'), 'utf8')).resolves.toBe('keep');
  });
});
