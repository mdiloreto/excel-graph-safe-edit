import { createHash } from 'node:crypto';
import { createReadStream } from 'node:fs';
import { open, stat, type FileHandle } from 'node:fs/promises';
import { TextDecoder } from 'node:util';

const END_OF_CENTRAL_DIRECTORY_SIGNATURE = 0x0605_4b50;
const CENTRAL_DIRECTORY_SIGNATURE = 0x0201_4b50;
const LOCAL_FILE_HEADER_SIGNATURE = 0x0403_4b50;
const MAX_EOCD_SEARCH = 65_535 + 22;

export interface OoxmlValidation {
  bytes: number;
  sha256: string;
}

type ReadExactly = (length: number, position: number) => Promise<Buffer>;

export function hasZipSignature(header: Uint8Array): boolean {
  return header.length >= 4 && Buffer.from(header).readUInt32LE(0) === LOCAL_FILE_HEADER_SIGNATURE;
}

function failure(label: string, detail: string): Error {
  return new Error(`${label} validation failed: ${detail}`);
}

function decodeEntryName(value: Buffer, label: string): string {
  let name: string;
  try {
    name = new TextDecoder('utf-8', { fatal: true }).decode(value);
  } catch {
    throw failure(label, 'ZIP entry name is not valid UTF-8');
  }
  if (/\p{Cc}/u.test(name)) throw failure(label, 'ZIP entry name contains a control character');
  return name;
}

async function readFileExactly(handle: FileHandle, length: number, position: number, label: string): Promise<Buffer> {
  const buffer = Buffer.alloc(length);
  let totalRead = 0;
  while (totalRead < length) {
    const { bytesRead } = await handle.read(buffer, totalRead, length - totalRead, position + totalRead);
    if (bytesRead === 0) throw failure(label, 'truncated ZIP structure');
    totalRead += bytesRead;
  }
  return buffer;
}

function findEndOfCentralDirectory(tail: Buffer, fileSize: number, label: string): { record: Buffer; offset: number } {
  for (let index = tail.length - 22; index >= 0; index -= 1) {
    if (tail.readUInt32LE(index) !== END_OF_CENTRAL_DIRECTORY_SIGNATURE) continue;
    const commentLength = tail.readUInt16LE(index + 20);
    if (index + 22 + commentLength !== tail.length) continue;
    return {
      record: tail.subarray(index, index + 22),
      offset: fileSize - tail.length + index,
    };
  }
  throw failure(label, 'ZIP end-of-central-directory record is missing or truncated');
}

async function validateZipStructure(
  readExactly: ReadExactly,
  fileSize: number,
  requiredEntries: ReadonlySet<string>,
  label: string,
): Promise<void> {
  if (fileSize < 22) throw failure(label, 'downloaded file is too small to be a ZIP archive');
  const tailLength = Math.min(fileSize, MAX_EOCD_SEARCH);
  const tail = await readExactly(tailLength, fileSize - tailLength);
  const { record: eocd, offset: eocdOffset } = findEndOfCentralDirectory(tail, fileSize, label);
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
  ) throw failure(label, 'ZIP64 archives are not supported');
  if (diskNumber !== 0 || centralDirectoryDisk !== 0 || entriesOnDisk !== totalEntries) {
    throw failure(label, 'multi-disk ZIP archives are not supported');
  }
  if (totalEntries === 0 || centralDirectoryOffset + centralDirectorySize !== eocdOffset) {
    throw failure(label, 'invalid ZIP central-directory bounds');
  }

  const entries = new Set<string>();
  const localRanges: Array<{ start: number; end: number }> = [];
  let cursor = centralDirectoryOffset;
  const centralDirectoryEnd = centralDirectoryOffset + centralDirectorySize;
  for (let index = 0; index < totalEntries; index += 1) {
    if (cursor + 46 > centralDirectoryEnd) throw failure(label, 'truncated ZIP central-directory entry');
    const centralHeader = await readExactly(46, cursor);
    if (centralHeader.readUInt32LE(0) !== CENTRAL_DIRECTORY_SIGNATURE) {
      throw failure(label, 'invalid ZIP central-directory signature');
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
      throw failure(label, 'ZIP64 entries are not supported');
    }
    if ((flags & 0x1) !== 0) throw failure(label, 'encrypted ZIP entries are not supported');
    if (entryDiskNumber !== 0) throw failure(label, 'multi-disk ZIP entries are not supported');
    const centralEntryLength = 46 + fileNameLength + extraLength + commentLength;
    if (cursor + centralEntryLength > centralDirectoryEnd) throw failure(label, 'truncated ZIP central-directory metadata');
    const fileName = decodeEntryName(await readExactly(fileNameLength, cursor + 46), label);
    if (entries.has(fileName)) throw failure(label, 'duplicate ZIP entry');
    entries.add(fileName);

    if (localHeaderOffset + 30 > centralDirectoryOffset) throw failure(label, 'invalid ZIP local-header offset');
    const localHeader = await readExactly(30, localHeaderOffset);
    if (localHeader.readUInt32LE(0) !== LOCAL_FILE_HEADER_SIGNATURE) {
      throw failure(label, 'invalid ZIP local-header signature');
    }
    const localFileNameLength = localHeader.readUInt16LE(26);
    const localExtraLength = localHeader.readUInt16LE(28);
    const localFlags = localHeader.readUInt16LE(6);
    const localCompressionMethod = localHeader.readUInt16LE(8);
    if (localFlags !== flags || localCompressionMethod !== compressionMethod) {
      throw failure(label, 'ZIP local and central entry metadata differ');
    }
    if ((flags & 0x8) === 0 && (
      localHeader.readUInt32LE(18) !== compressedSize
      || localHeader.readUInt32LE(22) !== uncompressedSize
    )) throw failure(label, 'ZIP local and central entry sizes differ');
    const localFileName = decodeEntryName(await readExactly(localFileNameLength, localHeaderOffset + 30), label);
    if (localFileName !== fileName) throw failure(label, 'ZIP local and central entry names differ');
    const localEntryEnd = localHeaderOffset + 30 + localFileNameLength + localExtraLength + compressedSize;
    if (localEntryEnd > centralDirectoryOffset) throw failure(label, 'ZIP entry data exceeds local-file area');
    localRanges.push({ start: localHeaderOffset, end: localEntryEnd });
    cursor += centralEntryLength;
  }
  if (cursor !== centralDirectoryEnd) throw failure(label, 'ZIP central-directory size does not match its entries');
  localRanges.sort((left, right) => left.start - right.start);
  for (let index = 1; index < localRanges.length; index += 1) {
    if ((localRanges[index - 1]?.end ?? 0) > (localRanges[index]?.start ?? 0)) {
      throw failure(label, 'ZIP local entries overlap');
    }
  }
  for (const requiredEntry of requiredEntries) {
    if (!entries.has(requiredEntry)) throw failure(label, `OOXML entry is missing: ${requiredEntry}`);
  }
}

export async function validateOoxmlFile(
  path: string,
  requiredEntries: ReadonlySet<string>,
  label: string,
): Promise<OoxmlValidation> {
  const info = await stat(path);
  if (!info.isFile() || info.size === 0) throw failure(label, 'downloaded file is empty or not regular');
  const handle = await open(path, 'r');
  try {
    await validateZipStructure(
      (length, position) => readFileExactly(handle, length, position, label),
      info.size,
      requiredEntries,
      label,
    );
  } finally {
    await handle.close();
  }
  const hash = createHash('sha256');
  for await (const chunk of createReadStream(path)) hash.update(chunk as Buffer);
  return { bytes: info.size, sha256: hash.digest('hex') };
}

export async function validateOoxmlBytes(
  value: Uint8Array,
  requiredEntries: ReadonlySet<string>,
  label: string,
): Promise<OoxmlValidation> {
  const bytes = Buffer.from(value);
  await validateZipStructure(async (length, position) => {
    if (position < 0 || length < 0 || position + length > bytes.length) throw failure(label, 'truncated ZIP structure');
    return bytes.subarray(position, position + length);
  }, bytes.length, requiredEntries, label);
  return {
    bytes: bytes.length,
    sha256: createHash('sha256').update(bytes).digest('hex'),
  };
}
