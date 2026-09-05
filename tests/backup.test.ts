import { chmod, mkdir, mkdtemp, readFile, readdir, rm, stat, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { backupFileName, downloadBackup, hasXlsxSignature, publishBackupFile, validateBackupFile } from '../src/backup.js';
import { buildAuthConfig } from '../src/config.js';

const temporaryDirectories: string[] = [];

function crc32(value: Buffer): number {
  let crc = 0xffff_ffff;
  for (const byte of value) {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit += 1) crc = (crc >>> 1) ^ (crc & 1 ? 0xedb8_8320 : 0);
  }
  return (crc ^ 0xffff_ffff) >>> 0;
}

function createMinimalZip(entryNames = ['[Content_Types].xml', 'xl/workbook.xml']): Buffer {
  const localParts: Buffer[] = [];
  const centralParts: Buffer[] = [];
  let localOffset = 0;
  for (const entryName of entryNames) {
    const name = Buffer.from(entryName);
    const data = Buffer.from(`<${entryName.includes('workbook') ? 'workbook' : 'Types'}/>`);
    const checksum = crc32(data);
    const localHeader = Buffer.alloc(30);
    localHeader.writeUInt32LE(0x0403_4b50, 0);
    localHeader.writeUInt16LE(20, 4);
    localHeader.writeUInt32LE(checksum, 14);
    localHeader.writeUInt32LE(data.length, 18);
    localHeader.writeUInt32LE(data.length, 22);
    localHeader.writeUInt16LE(name.length, 26);
    const localPart = Buffer.concat([localHeader, name, data]);
    localParts.push(localPart);

    const centralHeader = Buffer.alloc(46);
    centralHeader.writeUInt32LE(0x0201_4b50, 0);
    centralHeader.writeUInt16LE(20, 4);
    centralHeader.writeUInt16LE(20, 6);
    centralHeader.writeUInt32LE(checksum, 16);
    centralHeader.writeUInt32LE(data.length, 20);
    centralHeader.writeUInt32LE(data.length, 24);
    centralHeader.writeUInt16LE(name.length, 28);
    centralHeader.writeUInt32LE(localOffset, 42);
    centralParts.push(Buffer.concat([centralHeader, name]));
    localOffset += localPart.length;
  }
  const centralDirectory = Buffer.concat(centralParts);
  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x0605_4b50, 0);
  eocd.writeUInt16LE(entryNames.length, 8);
  eocd.writeUInt16LE(entryNames.length, 10);
  eocd.writeUInt32LE(centralDirectory.length, 12);
  eocd.writeUInt32LE(localOffset, 16);
  return Buffer.concat([...localParts, centralDirectory, eocd]);
}

afterEach(async () => {
  vi.restoreAllMocks();
  await Promise.all(temporaryDirectories.splice(0).map((path) => rm(path, { recursive: true, force: true })));
});

describe('backup helpers', () => {
  it('creates safe, bounded, collision-resistant backup names', () => {
    const timestamp = new Date('2026-08-07T17:30:00.123Z');
    const name = backupFileName('Budget 2026.xlsx', timestamp, 'nonce-1');
    expect(name).toBe('Budget_2026.xlsx.2026-08-07T17-30-00-123Z.nonce-1.backup.xlsx');
    expect(backupFileName('x'.repeat(500), timestamp, 'nonce').length).toBeLessThanOrEqual(180);
    expect(backupFileName('Budget.xlsx', timestamp, 'first')).not.toBe(backupFileName('Budget.xlsx', timestamp, 'second'));
  });

  it('recognizes the basic XLSX ZIP signature', () => {
    expect(hasXlsxSignature(Buffer.from([0x50, 0x4b, 0x03, 0x04]))).toBe(true);
    expect(hasXlsxSignature(Buffer.from('not a zip'))).toBe(false);
  });

  it('validates a structurally complete minimal XLSX ZIP and streams its hash', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'excel-graph-backup-'));
    temporaryDirectories.push(directory);
    const content = createMinimalZip();
    const path = join(directory, 'valid.xlsx');
    await writeFile(path, content, { mode: 0o600 });

    await expect(validateBackupFile(path)).resolves.toMatchObject({ bytes: content.length, sha256: expect.stringMatching(/^[a-f0-9]{64}$/) });
  });

  it('rejects garbage, truncation, invalid offsets, and missing XLSX entries', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'excel-graph-backup-'));
    temporaryDirectories.push(directory);
    const valid = createMinimalZip();
    const corruptedOffset = Buffer.from(valid);
    const centralOffset = corruptedOffset.indexOf(Buffer.from([0x50, 0x4b, 0x01, 0x02]));
    corruptedOffset.writeUInt32LE(valid.length + 100, centralOffset + 42);
    const cases: Array<{ name: string; content: Buffer; error: RegExp }> = [
      { name: 'garbage.xlsx', content: Buffer.from([0x50, 0x4b, 0x03, 0x04, 0x00]), error: /too small/ },
      { name: 'truncated.xlsx', content: valid.subarray(0, valid.length - 5), error: /end-of-central-directory/ },
      { name: 'offset.xlsx', content: corruptedOffset, error: /local-header offset/ },
      { name: 'missing.xlsx', content: createMinimalZip(['[Content_Types].xml']), error: /xl\/workbook\.xml/ },
    ];
    for (const testCase of cases) {
      const path = join(directory, testCase.name);
      await writeFile(path, testCase.content, { mode: 0o600 });
      await expect(validateBackupFile(path)).rejects.toThrow(testCase.error);
    }
  });

  it('rejects malformed UTF-8 and terminal-control canaries in OOXML entry names without reflecting them', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'excel-graph-backup-'));
    temporaryDirectories.push(directory);

    const malformed = createMinimalZip(['bad-name.bin', '[Content_Types].xml', 'xl/workbook.xml']);
    const centralOffset = malformed.indexOf(Buffer.from([0x50, 0x4b, 0x01, 0x02]));
    malformed.set([0xc3, 0x28], centralOffset + 46);
    const malformedPath = join(directory, 'malformed-name.xlsx');
    await writeFile(malformedPath, malformed, { mode: 0o600 });
    await expect(validateBackupFile(malformedPath)).rejects.toThrow(/entry name is not valid UTF-8/);

    const canary = '\u001b[31mTERMINAL-CANARY\u0007';
    const controlledPath = join(directory, 'controlled-name.xlsx');
    await writeFile(
      controlledPath,
      createMinimalZip([canary, '[Content_Types].xml', 'xl/workbook.xml']),
      { mode: 0o600 },
    );
    let error: unknown;
    try {
      await validateBackupFile(controlledPath);
    } catch (caught) {
      error = caught;
    }
    expect(error).toBeInstanceOf(Error);
    expect((error as Error).message).toMatch(/entry name contains a control character/);
    expect((error as Error).message).not.toContain('TERMINAL-CANARY');
    expect(Array.from((error as Error).message).some((character) => {
      const codePoint = character.codePointAt(0) ?? 0;
      return codePoint < 0x20 || codePoint === 0x7f;
    })).toBe(false);
  });

  it('publishes with an exclusive hard link and retries a collision', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'excel-graph-backup-'));
    temporaryDirectories.push(directory);
    const tempPath = join(directory, '.validated.partial');
    const collisionPath = join(directory, 'existing.backup.xlsx');
    await writeFile(tempPath, Buffer.from([0x50, 0x4b, 0x03, 0x04]), { mode: 0o600 });
    await writeFile(collisionPath, 'existing backup', { mode: 0o600 });
    const names = ['existing.backup.xlsx', 'new.backup.xlsx'];

    const publishedPath = await publishBackupFile(tempPath, directory, 'Book.xlsx', () => names.shift() ?? 'unused.backup.xlsx');

    expect(publishedPath).toBe(join(directory, 'new.backup.xlsx'));
    expect(await readFile(publishedPath)).toEqual(Buffer.from([0x50, 0x4b, 0x03, 0x04]));
    expect(await readFile(collisionPath, 'utf8')).toBe('existing backup');
    await expect(readFile(tempPath)).rejects.toMatchObject({ code: 'ENOENT' });
  });

  it('never overwrites when concurrent publishers choose the same first name', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'excel-graph-backup-'));
    temporaryDirectories.push(directory);
    const firstTemp = join(directory, '.first.partial');
    const secondTemp = join(directory, '.second.partial');
    await writeFile(firstTemp, 'first-content', { mode: 0o600 });
    await writeFile(secondTemp, 'second-content', { mode: 0o600 });
    const nameFactory = (fallback: string) => {
      let attempt = 0;
      return () => attempt++ === 0 ? 'same.backup.xlsx' : fallback;
    };

    const published = await Promise.all([
      publishBackupFile(firstTemp, directory, 'Book.xlsx', nameFactory('first.backup.xlsx')),
      publishBackupFile(secondTemp, directory, 'Book.xlsx', nameFactory('second.backup.xlsx')),
    ]);

    expect(new Set(published).size).toBe(2);
    const contents = await Promise.all(published.map((path) => readFile(path, 'utf8')));
    expect(new Set(contents)).toEqual(new Set(['first-content', 'second-content']));
  });

  it('rejects an existing permissive custom directory without changing it or authenticating', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'excel-graph-backup-'));
    temporaryDirectories.push(directory);
    const customDirectory = join(directory, 'custom');
    await mkdir(customDirectory, { mode: 0o755 });
    await chmod(customDirectory, 0o755);
    const fetchSpy = vi.spyOn(globalThis, 'fetch');

    await expect(downloadBackup(
      buildAuthConfig({ clientId: 'client', noPersist: true }),
      'item',
      customDirectory,
    )).rejects.toThrow(/group or other permissions/);

    expect((await stat(customDirectory)).mode & 0o777).toBe(0o755);
    expect(fetchSpy).not.toHaveBeenCalled();
  });

  it('cancels a slow-drip authenticated backup body and removes the partial file', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'excel-graph-backup-'));
    temporaryDirectories.push(directory);
    let cancelled = false;
    const body = new ReadableStream<Uint8Array>({
      start: (controller) => controller.enqueue(Buffer.from([0x50, 0x4b, 0x03, 0x04])),
      cancel: () => { cancelled = true; },
    });
    const fetchMock = vi.fn(async (_url: string, options?: RequestInit) => {
      expect(new Headers(options?.headers).get('authorization')).toBe('Bearer backup-token');
      return new Response(body, { status: 200 });
    });

    await expect(downloadBackup(
      buildAuthConfig({ clientId: 'client', noPersist: true }),
      'item',
      directory,
      undefined,
      {
        metadata: async () => ({ id: 'item', name: 'Book.xlsx' }),
        token: async () => 'backup-token',
        fetch: fetchMock,
        deadlines: { requestMilliseconds: 50, bodyMilliseconds: 10, cleanupMilliseconds: 5 },
      },
    )).rejects.toThrow(/^Backup download body timed out$/u);

    expect(cancelled).toBe(true);
    expect((await readdir(directory)).filter((name) => name.endsWith('.partial'))).toEqual([]);
  });
});
