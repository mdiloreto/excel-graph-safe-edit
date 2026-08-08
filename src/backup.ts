import { createHash, randomBytes } from 'node:crypto';
import { createReadStream, createWriteStream, type Stats } from 'node:fs';
import { chmod, lstat, link, mkdir, open, rm, stat, unlink, type FileHandle } from 'node:fs/promises';
import { basename, dirname, join, resolve } from 'node:path';
import { pipeline } from 'node:stream/promises';
import { getAccessToken } from './auth.js';
import { type AuthConfig, BACKUP_DIR, GRAPH_BASE } from './config.js';
import { itemPathById } from './excel.js';
import { graphRequest } from './graph.js';

interface DriveItemMetadata {
  id: string;
  name?: string;
  webUrl?: string;
}

export interface LocalBackup {
  path: string;
  bytes: number;
  sha256: string;
  item: {
    id: string;
    name?: string;
    webUrl?: string;
    driveId?: string;
  };
}

const MAX_BACKUP_FILE_NAME_LENGTH = 180;
const END_OF_CENTRAL_DIRECTORY_SIGNATURE = 0x0605_4b50;
const CENTRAL_DIRECTORY_SIGNATURE = 0x0201_4b50;
const LOCAL_FILE_HEADER_SIGNATURE = 0x0403_4b50;
const MAX_EOCD_SEARCH = 65_535 + 22;
const REQUIRED_XLSX_ENTRIES = new Set(['[Content_Types].xml', 'xl/workbook.xml']);

export function backupFileName(
  name: string | undefined,
  timestamp = new Date(),
  nonce = randomBytes(12).toString('hex'),
): string {
  const safeNonce = nonce.replace(/[^A-Za-z0-9_-]/g, '').slice(0, 32) || randomBytes(12).toString('hex');
  const suffix = `.${timestamp.toISOString().replace(/[:.]/g, '-')}.${safeNonce}.backup.xlsx`;
  const sourceName = basename(name ?? 'workbook.xlsx').replace(/[^A-Za-z0-9._-]/g, '_').replace(/^\.+|\.+$/g, '');
  const maxSourceLength = MAX_BACKUP_FILE_NAME_LENGTH - suffix.length;
  const safeName = (sourceName || 'workbook.xlsx').slice(0, Math.max(1, maxSourceLength));
  return `${safeName}${suffix}`;
}

export function hasXlsxSignature(header: Uint8Array): boolean {
  return header.length >= 4 && Buffer.from(header).readUInt32LE(0) === LOCAL_FILE_HEADER_SIGNATURE;
}

async function readExactly(handle: FileHandle, length: number, position: number): Promise<Buffer> {
  const buffer = Buffer.alloc(length);
  let totalRead = 0;
  while (totalRead < length) {
    const { bytesRead } = await handle.read(buffer, totalRead, length - totalRead, position + totalRead);
    if (bytesRead === 0) throw new Error('Backup validation failed: truncated ZIP structure');
    totalRead += bytesRead;
  }
  return buffer;
}

function findEndOfCentralDirectory(tail: Buffer, fileSize: number): { record: Buffer; offset: number } {
  for (let index = tail.length - 22; index >= 0; index -= 1) {
    if (tail.readUInt32LE(index) !== END_OF_CENTRAL_DIRECTORY_SIGNATURE) continue;
    const commentLength = tail.readUInt16LE(index + 20);
    if (index + 22 + commentLength !== tail.length) continue;
    return {
      record: tail.subarray(index, index + 22),
      offset: fileSize - tail.length + index,
    };
  }
  throw new Error('Backup validation failed: ZIP end-of-central-directory record is missing or truncated');
}

async function validateXlsxZip(handle: FileHandle, fileSize: number): Promise<void> {
  if (fileSize < 22) throw new Error('Backup validation failed: downloaded file is too small to be a ZIP archive');
  const tailLength = Math.min(fileSize, MAX_EOCD_SEARCH);
  const tail = await readExactly(handle, tailLength, fileSize - tailLength);
  const { record: eocd, offset: eocdOffset } = findEndOfCentralDirectory(tail, fileSize);
  const diskNumber = eocd.readUInt16LE(4);
  const centralDirectoryDisk = eocd.readUInt16LE(6);
  const entriesOnDisk = eocd.readUInt16LE(8);
  const totalEntries = eocd.readUInt16LE(10);
  const centralDirectorySize = eocd.readUInt32LE(12);
  const centralDirectoryOffset = eocd.readUInt32LE(16);
  if (
    entriesOnDisk === 0xffff
    || totalEntries === 0xffff
    || centralDirectorySize === 0xffff_ffff
    || centralDirectoryOffset === 0xffff_ffff
  ) throw new Error('Backup validation failed: ZIP64 archives are not supported');
  if (diskNumber !== 0 || centralDirectoryDisk !== 0 || entriesOnDisk !== totalEntries) {
    throw new Error('Backup validation failed: multi-disk ZIP archives are not supported');
  }
  if (totalEntries === 0 || centralDirectoryOffset + centralDirectorySize !== eocdOffset) {
    throw new Error('Backup validation failed: invalid ZIP central-directory bounds');
  }

  const entries = new Set<string>();
  const localRanges: Array<{ start: number; end: number }> = [];
  let cursor = centralDirectoryOffset;
  const centralDirectoryEnd = centralDirectoryOffset + centralDirectorySize;
  for (let index = 0; index < totalEntries; index += 1) {
    if (cursor + 46 > centralDirectoryEnd) throw new Error('Backup validation failed: truncated ZIP central-directory entry');
    const centralHeader = await readExactly(handle, 46, cursor);
    if (centralHeader.readUInt32LE(0) !== CENTRAL_DIRECTORY_SIGNATURE) {
      throw new Error('Backup validation failed: invalid ZIP central-directory signature');
    }
    const flags = centralHeader.readUInt16LE(8);
    const compressionMethod = centralHeader.readUInt16LE(10);
    const compressedSize = centralHeader.readUInt32LE(20);
    const uncompressedSize = centralHeader.readUInt32LE(24);
    const fileNameLength = centralHeader.readUInt16LE(28);
    const extraLength = centralHeader.readUInt16LE(30);
    const commentLength = centralHeader.readUInt16LE(32);
    const entryDiskNumber = centralHeader.readUInt16LE(34);
    const localHeaderOffset = centralHeader.readUInt32LE(42);
    if (compressedSize === 0xffff_ffff || uncompressedSize === 0xffff_ffff || localHeaderOffset === 0xffff_ffff) {
      throw new Error('Backup validation failed: ZIP64 entries are not supported');
    }
    if ((flags & 0x1) !== 0) throw new Error('Backup validation failed: encrypted ZIP entries are not supported');
    if (entryDiskNumber !== 0) throw new Error('Backup validation failed: multi-disk ZIP entries are not supported');
    const centralEntryLength = 46 + fileNameLength + extraLength + commentLength;
    if (cursor + centralEntryLength > centralDirectoryEnd) throw new Error('Backup validation failed: truncated ZIP central-directory metadata');
    const fileName = (await readExactly(handle, fileNameLength, cursor + 46)).toString('utf8');
    entries.add(fileName);

    if (localHeaderOffset + 30 > centralDirectoryOffset) throw new Error('Backup validation failed: invalid ZIP local-header offset');
    const localHeader = await readExactly(handle, 30, localHeaderOffset);
    if (localHeader.readUInt32LE(0) !== LOCAL_FILE_HEADER_SIGNATURE) {
      throw new Error('Backup validation failed: invalid ZIP local-header signature');
    }
    const localFileNameLength = localHeader.readUInt16LE(26);
    const localExtraLength = localHeader.readUInt16LE(28);
    const localFlags = localHeader.readUInt16LE(6);
    const localCompressionMethod = localHeader.readUInt16LE(8);
    if (localFlags !== flags || localCompressionMethod !== compressionMethod) {
      throw new Error('Backup validation failed: ZIP local and central entry metadata differ');
    }
    if ((flags & 0x8) === 0 && (
      localHeader.readUInt32LE(18) !== compressedSize
      || localHeader.readUInt32LE(22) !== uncompressedSize
    )) throw new Error('Backup validation failed: ZIP local and central entry sizes differ');
    const localFileName = (await readExactly(handle, localFileNameLength, localHeaderOffset + 30)).toString('utf8');
    if (localFileName !== fileName) throw new Error('Backup validation failed: ZIP local and central entry names differ');
    const localEntryEnd = localHeaderOffset + 30 + localFileNameLength + localExtraLength + compressedSize;
    if (localEntryEnd > centralDirectoryOffset) throw new Error('Backup validation failed: ZIP entry data exceeds local-file area');
    localRanges.push({ start: localHeaderOffset, end: localEntryEnd });
    cursor += centralEntryLength;
  }
  if (cursor !== centralDirectoryEnd) throw new Error('Backup validation failed: ZIP central-directory size does not match its entries');
  localRanges.sort((left, right) => left.start - right.start);
  for (let index = 1; index < localRanges.length; index += 1) {
    if ((localRanges[index - 1]?.end ?? 0) > (localRanges[index]?.start ?? 0)) {
      throw new Error('Backup validation failed: ZIP local entries overlap');
    }
  }
  for (const requiredEntry of REQUIRED_XLSX_ENTRIES) {
    if (!entries.has(requiredEntry)) throw new Error(`Backup validation failed: XLSX entry is missing: ${requiredEntry}`);
  }
}

export async function validateBackupFile(path: string): Promise<{ bytes: number; sha256: string }> {
  const info = await stat(path);
  if (!info.isFile() || info.size === 0) throw new Error('Backup validation failed: downloaded file is empty or not regular');
  const handle = await open(path, 'r');
  try {
    await validateXlsxZip(handle, info.size);
  } finally {
    await handle.close();
  }
  const hash = createHash('sha256');
  for await (const chunk of createReadStream(path)) hash.update(chunk as Buffer);
  return { bytes: info.size, sha256: hash.digest('hex') };
}

async function ensureBackupDirectory(dir: string): Promise<void> {
  let info: Stats;
  try {
    info = await lstat(dir);
  } catch (error) {
    if (!(error instanceof Error && 'code' in error && error.code === 'ENOENT')) throw error;
    await mkdir(dir, { recursive: true, mode: 0o700 });
    info = await lstat(dir);
  }
  if (info.isSymbolicLink() || !info.isDirectory()) throw new Error(`Refusing unsafe backup directory: ${dir}`);
  if ((info.mode & 0o077) !== 0) {
    throw new Error(`Refusing backup directory with group or other permissions: ${dir}`);
  }
}

function isFileExistsError(error: unknown): boolean {
  return error instanceof Error && 'code' in error && error.code === 'EEXIST';
}

export async function publishBackupFile(
  tempPath: string,
  dir: string,
  name: string | undefined,
  fileNameFactory: (name: string | undefined) => string = (value) => backupFileName(value),
): Promise<string> {
  if (resolve(dirname(tempPath)) !== resolve(dir)) throw new Error('Backup temporary and final files must share a directory');
  for (let attempt = 0; attempt < 10; attempt += 1) {
    const fileName = fileNameFactory(name);
    if (basename(fileName) !== fileName) throw new Error('Backup file name must not contain path segments');
    const path = join(dir, fileName);
    try {
      await link(tempPath, path);
      await unlink(tempPath);
      return path;
    } catch (error) {
      if (!isFileExistsError(error)) throw error;
    }
  }
  throw new Error('Could not publish a unique backup file name');
}

export async function downloadBackup(
  config: AuthConfig,
  itemId: string,
  dir = BACKUP_DIR,
  driveId?: string,
): Promise<LocalBackup> {
  await ensureBackupDirectory(dir);
  const itemPath = itemPathById(itemId, driveId);
  const metadata = await graphRequest<DriveItemMetadata>(config, itemPath);
  const token = await getAccessToken(config);
  const response = await fetch(`${GRAPH_BASE}${itemPath}/content`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!response.ok || !response.body) throw new Error(`Backup download failed: ${response.status}`);

  const tempPath = join(dir, `.backup-${process.pid}-${randomBytes(12).toString('hex')}.partial`);
  try {
    await pipeline(response.body, createWriteStream(tempPath, { flags: 'wx', mode: 0o600 }));
    await chmod(tempPath, 0o600);
    const validation = await validateBackupFile(tempPath);
    const finalPath = await publishBackupFile(tempPath, dir, metadata.name);
    return {
      path: finalPath,
      bytes: validation.bytes,
      sha256: validation.sha256,
      item: { id: metadata.id, name: metadata.name, webUrl: metadata.webUrl, driveId },
    };
  } catch (error) {
    try {
      await rm(tempPath, { force: true });
    } catch (cleanupError) {
      const failure = error instanceof Error ? error.message : String(error);
      const cleanup = cleanupError instanceof Error ? cleanupError.message : String(cleanupError);
      throw new Error(`Backup failed (${failure}) and partial-file cleanup failed (${cleanup})`);
    }
    throw error;
  }
}
