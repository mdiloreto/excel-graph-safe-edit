import { chmod, mkdir, mkdtemp, rm, symlink, utimes, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { buildAuthConfig } from '../src/config.js';
import { createAndUploadLocalFile, readSecureLocalFile } from '../src/local-file.js';

const temporaryDirectories: string[] = [];
const authConfig = buildAuthConfig({ clientId: 'client-id', noPersist: true });

afterEach(async () => {
  vi.restoreAllMocks();
  await Promise.all(temporaryDirectories.splice(0).map((path) => rm(path, { recursive: true, force: true })));
});

describe('secure local file input', () => {
  it('reads a bounded private regular file', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'secure-upload-file-'));
    temporaryDirectories.push(directory);
    const path = join(directory, 'transcript.vtt');
    const bytes = Buffer.from('WEBVTT\n\n00:00.000 --> 00:01.000\nHello');
    await writeFile(path, bytes, { mode: 0o600 });
    await chmod(path, 0o600);

    await expect(readSecureLocalFile(path)).resolves.toEqual(bytes);
  });

  it('rejects permissive files, symlinks, non-files, and files over the requested bound', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'secure-upload-file-'));
    temporaryDirectories.push(directory);
    const path = join(directory, 'transcript.vtt');
    await writeFile(path, 'WEBVTT', { mode: 0o600 });

    await chmod(path, 0o644);
    await expect(readSecureLocalFile(path)).rejects.toThrow(/group or other/);
    await chmod(path, 0o600);

    const link = join(directory, 'transcript-link.vtt');
    await symlink(path, link);
    await expect(readSecureLocalFile(link)).rejects.toThrow();

    const childDirectory = join(directory, 'not-a-file');
    await mkdir(childDirectory, { mode: 0o700 });
    await expect(readSecureLocalFile(childDirectory)).rejects.toThrow(/regular/);
    await expect(readSecureLocalFile(path, 3)).rejects.toThrow(/1 through 3 bytes/);
  });

  it('rejects a regular file reached through a symlink ancestor', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'secure-upload-file-'));
    temporaryDirectories.push(directory);
    const actualDirectory = join(directory, 'actual');
    await mkdir(actualDirectory, { mode: 0o700 });
    const path = join(actualDirectory, 'transcript.vtt');
    await writeFile(path, 'WEBVTT', { mode: 0o600 });
    const linkedDirectory = join(directory, 'linked');
    await symlink(actualDirectory, linkedDirectory);

    await expect(readSecureLocalFile(join(linkedDirectory, 'transcript.vtt'))).rejects.toThrow(/symlink ancestors/);
  });

  it('rejects in-place mutation during a bounded read even when inode and size remain unchanged', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'secure-upload-file-'));
    temporaryDirectories.push(directory);
    const path = join(directory, 'transcript.vtt');
    const bytes = Buffer.alloc(128 * 1024, 1);
    await writeFile(path, bytes, { mode: 0o600 });
    let mutated = false;

    await expect(readSecureLocalFile(path, bytes.length, {
      afterReadChunk: async () => {
        if (mutated) return;
        mutated = true;
        await writeFile(path, Buffer.alloc(bytes.length, 2), { mode: 0o600 });
        await utimes(path, new Date('2000-01-01T00:00:00.000Z'), new Date('2000-01-01T00:00:00.000Z'));
      },
    })).rejects.toThrow(/changed during read/);
    expect(mutated).toBe(true);
  });
});

describe('generic local file upload API', () => {
  it('reads bytes securely and invokes safe create without an extension requirement', async () => {
    const bytes = Buffer.from('WEBVTT');
    const result = {
      id: 'item',
      name: 'transcript.vtt',
      size: bytes.length,
      sha256: 'a'.repeat(64),
      verified: true as const,
    };
    const readFile = vi.fn(async () => bytes);
    const upload = vi.fn(async () => result);

    await expect(createAndUploadLocalFile(authConfig, {
      inputFile: '/private/transcript.vtt',
      path: 'Class/transcript.vtt',
    }, { readFile, upload })).resolves.toEqual(result);

    expect(readFile).toHaveBeenCalledWith('/private/transcript.vtt');
    expect(upload).toHaveBeenCalledWith(authConfig, {
      path: 'Class/transcript.vtt',
      bytes,
      driveId: undefined,
    });
  });
});
