import { randomBytes } from 'node:crypto';
import { type Stats } from 'node:fs';
import { chmod, lstat, link, mkdir, open, rm, unlink, type FileHandle } from 'node:fs/promises';
import { basename, dirname, join, resolve } from 'node:path';
import { getAccessToken } from './auth.js';
import { type AuthConfig, BACKUP_DIR, GRAPH_BASE } from './config.js';
import { itemPathById } from './excel.js';
import {
  consumeResponseBody,
  discardResponse,
  fetchWithDeadline,
  graphRequest,
  resolveTransportDeadlines,
  type TransportDeadlines,
} from './graph.js';
import { hasZipSignature, validateOoxmlFile } from './ooxml.js';

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
const MAX_BACKUP_BYTES = 100 * 1024 * 1024;
const BACKUP_TRANSPORT_DEADLINES: TransportDeadlines = {
  requestMilliseconds: 60_000,
  bodyMilliseconds: 5 * 60_000,
  cleanupMilliseconds: 1000,
};
const REQUIRED_XLSX_ENTRIES = new Set(['[Content_Types].xml', 'xl/workbook.xml']);

export interface BackupDownloadDependencies {
  metadata: (config: AuthConfig, path: string) => Promise<DriveItemMetadata>;
  token: (config: AuthConfig) => Promise<string>;
  fetch: (url: string, options?: RequestInit) => Promise<Response>;
  deadlines?: Partial<TransportDeadlines>;
}

const defaultBackupDependencies: BackupDownloadDependencies = {
  metadata: (config, path) => graphRequest<DriveItemMetadata>(config, path),
  token: getAccessToken,
  fetch: (url, options) => fetch(url, options),
};

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
  return hasZipSignature(header);
}

export async function validateBackupFile(path: string): Promise<{ bytes: number; sha256: string }> {
  try {
    return await validateOoxmlFile(path, REQUIRED_XLSX_ENTRIES, 'Backup');
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    throw new Error(message.replace('OOXML entry is missing:', 'XLSX entry is missing:'));
  }
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

async function writeBoundedResponse(
  path: string,
  response: Response,
  deadlines: TransportDeadlines,
): Promise<void> {
  const contentLength = response.headers.get('content-length');
  if (contentLength && /^\d+$/u.test(contentLength) && BigInt(contentLength) > BigInt(MAX_BACKUP_BYTES)) {
    await discardResponse(response, deadlines.cleanupMilliseconds);
    throw new Error(`Backup download exceeded ${MAX_BACKUP_BYTES} bytes`);
  }
  let handle: FileHandle;
  try {
    handle = await open(path, 'wx', 0o600);
  } catch (error) {
    await discardResponse(response, deadlines.cleanupMilliseconds);
    throw error;
  }
  let totalBytes = 0;
  try {
    await consumeResponseBody(response, 'Backup download', deadlines, async (reader) => {
      while (true) {
        const result = await reader.read();
        if (result.done) break;
        const chunk = Buffer.from(result.value);
        totalBytes += chunk.length;
        if (totalBytes > MAX_BACKUP_BYTES) {
          throw new Error(`Backup download exceeded ${MAX_BACKUP_BYTES} bytes`);
        }
        let written = 0;
        while (written < chunk.length) {
          const writeResult = await handle.write(chunk, written, chunk.length - written);
          if (writeResult.bytesWritten === 0) throw new Error('Backup download could not write the response body');
          written += writeResult.bytesWritten;
        }
      }
    });
  } finally {
    await handle.close();
  }
}

export async function downloadBackup(
  config: AuthConfig,
  itemId: string,
  dir = BACKUP_DIR,
  driveId?: string,
  dependencies: BackupDownloadDependencies = defaultBackupDependencies,
): Promise<LocalBackup> {
  const deadlines = resolveTransportDeadlines(BACKUP_TRANSPORT_DEADLINES, dependencies.deadlines);
  await ensureBackupDirectory(dir);
  const itemPath = itemPathById(itemId, driveId);
  const metadata = await dependencies.metadata(config, itemPath);
  const token = await dependencies.token(config);
  const response = await fetchWithDeadline(
    dependencies.fetch,
    `${GRAPH_BASE}${itemPath}/content`,
    { headers: { Authorization: `Bearer ${token}` } },
    'Backup download request',
    deadlines,
  );
  if (!response.ok) {
    await discardResponse(response, deadlines.cleanupMilliseconds);
    throw new Error(`Backup download failed: ${response.status}`);
  }

  const tempPath = join(dir, `.backup-${process.pid}-${randomBytes(12).toString('hex')}.partial`);
  try {
    await writeBoundedResponse(tempPath, response, deadlines);
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
