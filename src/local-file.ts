import { constants, type BigIntStats } from 'node:fs';
import { lstat, open, realpath, type FileHandle } from 'node:fs/promises';
import { dirname, join, parse, relative, resolve, sep } from 'node:path';
import { type AuthConfig } from './config.js';
import {
  MAX_SAFE_UPLOAD_BYTES,
  safeCreateDriveFile,
  type VerifiedDriveItem,
} from './drive.js';

export interface SecureLocalReadDependencies {
  afterReadChunk?: (totalBytes: number) => void | Promise<void>;
}

async function readBoundedFile(
  handle: FileHandle,
  maximumBytes: number,
  dependencies: SecureLocalReadDependencies,
): Promise<Buffer> {
  const chunks: Buffer[] = [];
  let totalBytes = 0;
  while (totalBytes <= maximumBytes) {
    const chunk = Buffer.allocUnsafe(Math.min(64 * 1024, maximumBytes - totalBytes + 1));
    const { bytesRead } = await handle.read(chunk, 0, chunk.length, totalBytes);
    if (bytesRead === 0) break;
    chunks.push(chunk.subarray(0, bytesRead));
    totalBytes += bytesRead;
    await dependencies.afterReadChunk?.(totalBytes);
  }
  if (totalBytes === 0 || totalBytes > maximumBytes) {
    throw new Error(`Secure local file must contain 1 through ${maximumBytes} bytes`);
  }
  return Buffer.concat(chunks, totalBytes);
}

interface PathComponentIdentity {
  path: string;
  dev: bigint;
  ino: bigint;
}

function sameFileIdentity(left: Pick<BigIntStats, 'dev' | 'ino'>, right: Pick<BigIntStats, 'dev' | 'ino'>): boolean {
  return left.dev === right.dev && left.ino === right.ino;
}

function sameFileState(
  left: Pick<BigIntStats, 'size' | 'mtimeNs' | 'ctimeNs'>,
  right: Pick<BigIntStats, 'size' | 'mtimeNs' | 'ctimeNs'>,
): boolean {
  return left.size === right.size && left.mtimeNs === right.mtimeNs && left.ctimeNs === right.ctimeNs;
}

async function validateAncestorComponents(path: string): Promise<PathComponentIdentity[]> {
  const root = parse(path).root;
  const parent = dirname(path);
  const relativeParent = relative(root, parent);
  const segments = relativeParent ? relativeParent.split(sep) : [];
  const identities: PathComponentIdentity[] = [];
  let current = root;
  for (const segment of segments) {
    current = join(current, segment);
    const info = await lstat(current, { bigint: true });
    if (info.isSymbolicLink()) throw new Error('Secure local file path must not contain symlink ancestors');
    if (!info.isDirectory()) throw new Error('Secure local file ancestors must be directories');
    identities.push({ path: current, dev: info.dev, ino: info.ino });
  }
  return identities;
}

async function revalidateOpenedPath(
  path: string,
  handleInfo: BigIntStats,
  expectedAncestors: readonly PathComponentIdentity[],
): Promise<void> {
  const actualPath = await realpath(path);
  if (actualPath !== path) throw new Error('Secure local file path changed or resolved through a symlink');
  const pathInfo = await lstat(path, { bigint: true });
  if (
    pathInfo.isSymbolicLink()
    || !pathInfo.isFile()
    || !sameFileIdentity(pathInfo, handleInfo)
    || !sameFileState(pathInfo, handleInfo)
  ) {
    throw new Error('Secure local file path no longer identifies the opened regular file');
  }
  const actualAncestors = await validateAncestorComponents(path);
  if (
    actualAncestors.length !== expectedAncestors.length
    || actualAncestors.some((component, index) => {
      const expected = expectedAncestors[index];
      return !expected || component.path !== expected.path || component.dev !== expected.dev || component.ino !== expected.ino;
    })
  ) {
    throw new Error('Secure local file ancestor changed during validation');
  }
}

export async function readSecureLocalFile(
  path: string,
  maximumBytes = MAX_SAFE_UPLOAD_BYTES,
  dependencies: SecureLocalReadDependencies = {},
): Promise<Buffer> {
  if (typeof path !== 'string' || path.length === 0 || path.includes('\u0000')) {
    throw new Error('Secure local file path must not be empty or contain NUL');
  }
  if (!Number.isSafeInteger(maximumBytes) || maximumBytes < 1 || maximumBytes > MAX_SAFE_UPLOAD_BYTES) {
    throw new Error(`Secure local file limit must be between 1 and ${MAX_SAFE_UPLOAD_BYTES} bytes`);
  }
  if (process.platform === 'win32') {
    throw new Error('Secure local file validation is not supported on Windows');
  }
  const absolutePath = resolve(path);
  const ancestors = await validateAncestorComponents(absolutePath);
  const handle = await open(absolutePath, constants.O_RDONLY | constants.O_NOFOLLOW);
  try {
    const info = await handle.stat({ bigint: true });
    if (!info.isFile() || info.size === 0n || info.size > BigInt(maximumBytes)) {
      throw new Error(`Secure local file must be regular and contain 1 through ${maximumBytes} bytes`);
    }
    if ((info.mode & 0o077n) !== 0n) {
      throw new Error('Secure local file must not grant permissions to group or other users');
    }
    if (process.getuid && info.uid !== BigInt(process.getuid())) {
      throw new Error('Secure local file must be owned by the current user');
    }
    await revalidateOpenedPath(absolutePath, info, ancestors);
    const bytes = await readBoundedFile(handle, maximumBytes, dependencies);
    const postReadInfo = await handle.stat({ bigint: true });
    if (!sameFileIdentity(postReadInfo, info) || !sameFileState(postReadInfo, info)) {
      throw new Error('Secure local file changed during read');
    }
    await revalidateOpenedPath(absolutePath, info, ancestors);
    return bytes;
  } finally {
    await handle.close();
  }
}

export interface LocalFileUploadDependencies {
  readFile: typeof readSecureLocalFile;
  upload: typeof safeCreateDriveFile;
}

export async function createAndUploadLocalFile(
  config: AuthConfig,
  input: { inputFile: string; path: string; driveId?: string },
  dependencies: LocalFileUploadDependencies = {
    readFile: readSecureLocalFile,
    upload: safeCreateDriveFile,
  },
): Promise<VerifiedDriveItem> {
  const bytes = await dependencies.readFile(input.inputFile);
  return dependencies.upload(config, {
    path: input.path,
    bytes,
    driveId: input.driveId,
  });
}
