import { createHash } from 'node:crypto';
import { getAccessToken } from './auth.js';
import { type AuthConfig, GRAPH_BASE } from './config.js';
import {
  discardResponse,
  fetchWithDeadline,
  graphRequest,
  GraphRequestError,
  isAmbiguousWriteError,
  readBoundedJson,
  readBoundedResponse,
  resolveTransportDeadlines,
  type TransportDeadlines,
} from './graph.js';

export const GRAPH_UPLOAD_GRANULARITY = 320 * 1024;
export const GRAPH_UPLOAD_CHUNK_BYTES = 16 * GRAPH_UPLOAD_GRANULARITY;
export const MAX_SAFE_UPLOAD_BYTES = 20 * 1024 * 1024;

const MAX_DRIVE_PATH_LENGTH = 400;
const MAX_DRIVE_SEGMENT_LENGTH = 255;
const MAX_DRIVE_ID_LENGTH = 1024;
const MAX_ITEM_ID_LENGTH = 1024;
const MAX_WEB_URL_LENGTH = 4096;
const MAX_UPLOAD_URL_LENGTH = 8192;
const MAX_UPLOAD_RESPONSE_BYTES = 64 * 1024;
const MAX_UPLOAD_ATTEMPTS = 4;
const MAX_UPLOAD_REQUEST_ATTEMPTS = 256;
const DRIVE_TRANSPORT_DEADLINES: TransportDeadlines = {
  requestMilliseconds: 120_000,
  bodyMilliseconds: 120_000,
  cleanupMilliseconds: 1000,
};
const ITEM_SELECT = '$select=id,name,size,webUrl,parentReference,file,folder';
const FORBIDDEN_NAME_CHARACTERS = /["*:<>?\\|]/u;
const RESERVED_NAME = /^(?:\.lock|con|prn|aux|nul|com[0-9]|lpt[0-9]|desktop\.ini)(?:\..*)?$/iu;

export interface DrivePathTarget {
  path: string;
  parentSegments: string[];
  fileName: string;
}

export interface VerifiedDriveItem {
  id: string;
  name: string;
  size: number;
  webUrl?: string;
  driveId?: string;
  sha256: string;
  verified: true;
}

interface DriveItemRecord {
  id: string;
  name?: string;
  size?: number;
  webUrl?: string;
  folder?: Record<string, unknown>;
  file?: Record<string, unknown>;
  parentReference?: {
    id?: string;
    driveId?: string;
    path?: string;
  };
}

interface UploadProgress {
  kind: 'progress';
  nextOffset: number;
}

interface UploadComplete {
  kind: 'complete';
  itemId: string;
}

type UploadState = UploadProgress | UploadComplete | { kind: 'expired' };

export interface DriveUploadDependencies {
  graph: (config: AuthConfig, path: string, options?: RequestInit) => Promise<unknown>;
  token: (config: AuthConfig) => Promise<string>;
  fetch: (url: string, options?: RequestInit) => Promise<Response>;
  sleep: (milliseconds: number) => Promise<void>;
  deadlines?: Partial<TransportDeadlines>;
}

const defaultDependencies: DriveUploadDependencies = {
  graph: graphRequest,
  token: getAccessToken,
  fetch: (url, options) => fetch(url, options),
  sleep: (milliseconds) => new Promise((resolve) => setTimeout(resolve, milliseconds)),
};

export class DriveItemConflictError extends Error {
  constructor(message = 'A drive item already exists at the target path') {
    super(message);
    this.name = 'DriveItemConflictError';
  }
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === 'object' && !Array.isArray(value);
}

function codePointLength(value: string): number {
  return Array.from(value).length;
}

function hasControlCharacter(value: string): boolean {
  return Array.from(value).some((character) => (character.codePointAt(0) ?? 0) < 0x20);
}

function boundedString(value: unknown, field: string, maximum: number): string {
  if (typeof value !== 'string' || value.length === 0 || codePointLength(value) > maximum) {
    throw new Error(`${field} must be a non-empty string of at most ${maximum} characters`);
  }
  if (!value.isWellFormed() || /\p{Cc}/u.test(value)) throw new Error(`${field} must not contain invalid or control characters`);
  return value;
}

export function encodeGraphPathSegment(value: string): string {
  return encodeURIComponent(value).replace(/[!'()*]/g, (character) => `%${character.charCodeAt(0).toString(16).toUpperCase()}`);
}

export function encodeOneDrivePath(path: string): string {
  return path.split('/').filter(Boolean).map(encodeGraphPathSegment).join('/');
}

function assertDriveId(driveId: string | undefined): void {
  if (driveId === undefined) return;
  if (!driveId.trim()) throw new Error('Drive ID must not be empty');
  boundedString(driveId, 'Drive ID', MAX_DRIVE_ID_LENGTH);
  if (driveId.trim() !== driveId) throw new Error('Drive ID must not have leading or trailing whitespace');
}

export function driveRootPath(driveId?: string): string {
  assertDriveId(driveId);
  return driveId === undefined ? '/me/drive/root' : `/drives/${encodeGraphPathSegment(driveId)}/root`;
}

export function driveItemPathById(itemId: string, driveId?: string): string {
  if (!itemId.trim()) throw new Error('Item ID must not be empty');
  const item = boundedString(itemId, 'Item ID', MAX_ITEM_ID_LENGTH);
  assertDriveId(driveId);
  return driveId === undefined
    ? `/me/drive/items/${encodeGraphPathSegment(item)}`
    : `/drives/${encodeGraphPathSegment(driveId)}/items/${encodeGraphPathSegment(item)}`;
}

export function driveItemPathByPath(path: string, driveId?: string): string {
  assertDriveId(driveId);
  return `${driveRootPath(driveId)}:/${encodeOneDrivePath(path)}:`;
}

export function driveChildrenPath(parentItemId: string, driveId?: string): string {
  return `${driveItemPathById(parentItemId, driveId)}/children`;
}

export function driveUploadSessionPath(parentItemId: string, fileName: string, driveId?: string): string {
  return `${driveItemPathById(parentItemId, driveId)}:/${encodeGraphPathSegment(fileName)}:/createUploadSession`;
}

function assertDriveName(name: string, index: number): void {
  const label = `OneDrive path segment ${index + 1}`;
  if (codePointLength(name) > MAX_DRIVE_SEGMENT_LENGTH) throw new Error(`${label} exceeds ${MAX_DRIVE_SEGMENT_LENGTH} characters`);
  if (!name || name === '.' || name === '..') throw new Error(`${label} is empty or reserved`);
  if (name.trim() !== name) throw new Error(`${label} must not have leading or trailing whitespace`);
  if (FORBIDDEN_NAME_CHARACTERS.test(name) || hasControlCharacter(name)) {
    throw new Error(`${label} contains a character OneDrive does not allow`);
  }
  if (name.endsWith('.')) throw new Error(`${label} must not end with a period`);
  if (name.startsWith('~') || name.startsWith('~$')) throw new Error(`${label} uses a reserved OneDrive prefix`);
  if (name.toLocaleLowerCase('en-US').includes('_vti_') || RESERVED_NAME.test(name)) {
    throw new Error(`${label} uses a reserved OneDrive name`);
  }
  if (index === 0 && name.toLocaleLowerCase('en-US') === 'forms') {
    throw new Error('The root-level OneDrive name forms is reserved');
  }
}

export function parseDriveFilePath(path: string, requiredExtension?: string): DrivePathTarget {
  if (typeof path !== 'string' || path.length === 0) throw new Error('OneDrive path must not be empty');
  if (path.trim() !== path) throw new Error('OneDrive path must not have leading or trailing whitespace');
  if (codePointLength(path) > MAX_DRIVE_PATH_LENGTH) throw new Error(`OneDrive path exceeds ${MAX_DRIVE_PATH_LENGTH} characters`);
  const withoutRootSlash = path.startsWith('/') ? path.slice(1) : path;
  if (!withoutRootSlash || withoutRootSlash.endsWith('/') || withoutRootSlash.includes('//')) {
    throw new Error('OneDrive path must identify a file with no empty segments');
  }
  const segments = withoutRootSlash.split('/');
  segments.forEach(assertDriveName);
  const fileName = segments.at(-1);
  if (!fileName) throw new Error('OneDrive path must identify a file');
  if (requiredExtension && !fileName.toLocaleLowerCase('en-US').endsWith(requiredExtension.toLocaleLowerCase('en-US'))) {
    throw new Error(`OneDrive target filename must end with ${requiredExtension}`);
  }
  return {
    path: segments.join('/'),
    parentSegments: segments.slice(0, -1),
    fileName,
  };
}

function parseDriveItem(value: unknown, context: string): DriveItemRecord {
  if (!isRecord(value)) throw new Error(`${context} response must be an object`);
  const item: DriveItemRecord = { id: boundedString(value.id, `${context} item id`, MAX_ITEM_ID_LENGTH) };
  if (value.name !== undefined) item.name = boundedString(value.name, `${context} item name`, MAX_DRIVE_SEGMENT_LENGTH);
  if (value.webUrl !== undefined) {
    const webUrl = boundedString(value.webUrl, `${context} web URL`, MAX_WEB_URL_LENGTH);
    const parsed = new URL(webUrl);
    if (parsed.protocol !== 'https:' || parsed.username || parsed.password) throw new Error(`${context} web URL must be HTTPS without credentials`);
    item.webUrl = webUrl;
  }
  if (value.size !== undefined) {
    if (!Number.isSafeInteger(value.size) || (value.size as number) < 0) {
      throw new Error(`${context} item size must be a safe nonnegative integer`);
    }
    item.size = value.size as number;
  }
  if (value.folder !== undefined) {
    if (!isRecord(value.folder)) throw new Error(`${context} folder facet must be an object`);
    item.folder = value.folder;
  }
  if (value.file !== undefined) {
    if (!isRecord(value.file)) throw new Error(`${context} file facet must be an object`);
    item.file = value.file;
  }
  if (value.parentReference !== undefined) {
    if (!isRecord(value.parentReference)) throw new Error(`${context} parent reference must be an object`);
    const parentReference: NonNullable<DriveItemRecord['parentReference']> = {};
    if (value.parentReference.id !== undefined) {
      parentReference.id = boundedString(value.parentReference.id, `${context} parent item id`, MAX_ITEM_ID_LENGTH);
    }
    if (value.parentReference.driveId !== undefined) {
      parentReference.driveId = boundedString(value.parentReference.driveId, `${context} parent drive id`, MAX_DRIVE_ID_LENGTH);
    }
    if (value.parentReference.path !== undefined) {
      parentReference.path = boundedString(value.parentReference.path, `${context} parent path`, MAX_WEB_URL_LENGTH);
    }
    item.parentReference = parentReference;
  }
  return item;
}

function assertFolder(value: unknown, expectedName: string | undefined, context: string): DriveItemRecord {
  const item = parseDriveItem(value, context);
  if (!item.folder) throw new Error(`${context} resolved to a file instead of a folder`);
  if (expectedName !== undefined && item.name !== expectedName) {
    throw new Error(`${context} returned a different folder name`);
  }
  return item;
}

async function getItemByPath(
  config: AuthConfig,
  path: string,
  driveId: string | undefined,
  dependencies: DriveUploadDependencies,
): Promise<DriveItemRecord | null> {
  try {
    return parseDriveItem(await dependencies.graph(config, `${driveItemPathByPath(path, driveId)}?${ITEM_SELECT}`), 'Drive lookup');
  } catch (error) {
    if (error instanceof GraphRequestError && error.status === 404) return null;
    throw error;
  }
}

export async function ensureDriveParentFolder(
  config: AuthConfig,
  parentSegments: readonly string[],
  driveId?: string,
  dependencies: DriveUploadDependencies = defaultDependencies,
): Promise<{ id: string }> {
  const root = assertFolder(
    await dependencies.graph(config, `${driveRootPath(driveId)}?${ITEM_SELECT}`),
    undefined,
    'Drive root',
  );
  let parent = root;
  const accumulated: string[] = [];
  for (const [index, segment] of parentSegments.entries()) {
    assertDriveName(segment, index);
    accumulated.push(segment);
    const path = accumulated.join('/');
    const existing = await getItemByPath(config, path, driveId, dependencies);
    if (existing) {
      parent = assertFolder(existing, segment, 'Parent folder lookup');
      continue;
    }
    try {
      parent = assertFolder(await dependencies.graph(config, driveChildrenPath(parent.id, driveId), {
        method: 'POST',
        body: JSON.stringify({
          name: segment,
          folder: {},
          '@microsoft.graph.conflictBehavior': 'fail',
        }),
      }), segment, 'Parent folder creation');
    } catch (error) {
      if (!(error instanceof GraphRequestError && error.status === 409) && !isAmbiguousWriteError(error)) throw error;
      const reconciled = await getItemByPath(config, path, driveId, dependencies);
      if (!reconciled) {
        throw new Error('Parent folder creation had an ambiguous result; rerun to reconcile safely');
      }
      parent = assertFolder(reconciled, segment, 'Parent folder reconciliation');
    }
  }
  return { id: parent.id };
}

function nextExpectedOffset(value: unknown, totalBytes: number, context: string): number {
  if (!isRecord(value) || !Array.isArray(value.nextExpectedRanges) || value.nextExpectedRanges.length === 0 || value.nextExpectedRanges.length > 100) {
    throw new Error(`${context} did not include bounded nextExpectedRanges`);
  }
  const starts: number[] = [];
  for (const range of value.nextExpectedRanges) {
    if (typeof range !== 'string' || range.length > 50) throw new Error(`${context} included an invalid expected range`);
    const match = /^(\d+)-(?:\d+)?$/u.exec(range);
    const start = match?.[1] ? Number(match[1]) : Number.NaN;
    if (!Number.isSafeInteger(start) || start < 0 || start > totalBytes) throw new Error(`${context} included an invalid expected range`);
    starts.push(start);
  }
  return Math.min(...starts);
}

function completeItemId(value: unknown, context: string): string {
  if (!isRecord(value)) throw new Error(`${context} response must be an object`);
  return boundedString(value.id, `${context} item id`, MAX_ITEM_ID_LENGTH);
}

function parseUploadUrl(value: unknown): string {
  if (!isRecord(value)) throw new Error('Upload session response must be an object');
  const uploadUrl = boundedString(value.uploadUrl, 'Upload session URL', MAX_UPLOAD_URL_LENGTH);
  let parsed: URL;
  try {
    parsed = new URL(uploadUrl);
  } catch {
    throw new Error('Upload session URL must be a valid preauthenticated HTTPS URL');
  }
  if (parsed.protocol !== 'https:' || parsed.username || parsed.password || parsed.hash) {
    throw new Error('Upload session URL must be a preauthenticated HTTPS URL without credentials or a fragment');
  }
  return uploadUrl;
}

function isRetryableUploadStatus(status: number): boolean {
  return status === 408 || status === 429 || status === 500 || status === 502 || status === 503 || status === 504;
}

function retryDelay(response: Response | undefined, attempt: number): number {
  const retryAfter = response?.headers.get('retry-after');
  if (retryAfter && /^\d+$/u.test(retryAfter)) return Math.min(Number(retryAfter) * 1000, 120_000);
  return Math.min(500 * (2 ** attempt), 8000);
}

async function uploadStatus(
  uploadUrl: string,
  totalBytes: number,
  dependencies: DriveUploadDependencies,
  consumeRequestAttempt: () => void,
  deadlines: TransportDeadlines,
): Promise<UploadState> {
  for (let attempt = 0; attempt < MAX_UPLOAD_ATTEMPTS; attempt += 1) {
    consumeRequestAttempt();
    let response: Response;
    try {
      response = await fetchWithDeadline(
        dependencies.fetch,
        uploadUrl,
        { method: 'GET', headers: { Accept: 'application/json' } },
        'Upload-session status request',
        deadlines,
      );
    } catch {
      if (attempt === MAX_UPLOAD_ATTEMPTS - 1) throw new Error('Upload-session status could not be reconciled after a transport failure');
      await dependencies.sleep(retryDelay(undefined, attempt));
      continue;
    }
    if (response.status === 404) {
      await discardResponse(response, deadlines.cleanupMilliseconds);
      return { kind: 'expired' };
    }
    if (response.ok) {
      const payload = await readBoundedJson(response, MAX_UPLOAD_RESPONSE_BYTES, 'Upload-session status', deadlines);
      if (isRecord(payload) && payload.id !== undefined) return { kind: 'complete', itemId: completeItemId(payload, 'Upload-session status') };
      return { kind: 'progress', nextOffset: nextExpectedOffset(payload, totalBytes, 'Upload-session status') };
    }
    if ((response.status === 429 || response.status >= 500) && attempt < MAX_UPLOAD_ATTEMPTS - 1) {
      await discardResponse(response, deadlines.cleanupMilliseconds);
      await dependencies.sleep(retryDelay(response, attempt));
      continue;
    }
    await discardResponse(response, deadlines.cleanupMilliseconds);
    throw new Error(`Upload-session status failed with HTTP ${response.status}`);
  }
  throw new Error('Upload-session status could not be reconciled');
}

export async function uploadBytesToSession(
  uploadUrl: string,
  value: Uint8Array,
  options: {
    dependencies?: DriveUploadDependencies;
    chunkBytes?: number;
    findCompletedItem?: () => Promise<unknown | null>;
  } = {},
): Promise<{ itemId: string; transport: 'confirmed' | 'reconciled' }> {
  const dependencies = options.dependencies ?? defaultDependencies;
  const deadlines = resolveTransportDeadlines(DRIVE_TRANSPORT_DEADLINES, dependencies.deadlines);
  const bytes = Buffer.from(value);
  if (bytes.length === 0 || bytes.length > MAX_SAFE_UPLOAD_BYTES) {
    throw new Error(`Upload must contain 1 through ${MAX_SAFE_UPLOAD_BYTES} bytes`);
  }
  const chunkBytes = options.chunkBytes ?? GRAPH_UPLOAD_CHUNK_BYTES;
  if (!Number.isSafeInteger(chunkBytes) || chunkBytes <= 0 || chunkBytes >= 60 * 1024 * 1024 || chunkBytes % GRAPH_UPLOAD_GRANULARITY !== 0) {
    throw new Error(`Upload chunk size must be a positive multiple of ${GRAPH_UPLOAD_GRANULARITY} below 60 MiB`);
  }
  parseUploadUrl({ uploadUrl });

  let offset = 0;
  let highestConfirmedOffset = 0;
  let reconciled = false;
  let attemptsForOffset = 0;
  let requestAttempts = 0;
  const consumeRequestAttempt = (): void => {
    requestAttempts += 1;
    if (requestAttempts > MAX_UPLOAD_REQUEST_ATTEMPTS) {
      throw new Error('Upload-session request limit reached before completion');
    }
  };
  while (offset < bytes.length) {
    const endExclusive = Math.min(offset + chunkBytes, bytes.length);
    const chunk = bytes.subarray(offset, endExclusive);
    consumeRequestAttempt();
    let response: Response | undefined;
    try {
      response = await fetchWithDeadline(
        dependencies.fetch,
        uploadUrl,
        {
          method: 'PUT',
          headers: {
            'Content-Length': String(chunk.length),
            'Content-Range': `bytes ${offset}-${endExclusive - 1}/${bytes.length}`,
          },
          body: chunk,
        },
        'Upload fragment request',
        deadlines,
      );
    } catch {
      response = undefined;
    }

    if (response?.status === 200 || response?.status === 201) {
      if (endExclusive !== bytes.length) throw new Error('Upload session completed before all bytes were sent');
      return {
        itemId: completeItemId(
          await readBoundedJson(response, MAX_UPLOAD_RESPONSE_BYTES, 'Upload completion', deadlines),
          'Upload completion',
        ),
        transport: reconciled ? 'reconciled' : 'confirmed',
      };
    }
    if (response?.status === 202) {
      const nextOffset = nextExpectedOffset(
        await readBoundedJson(response, MAX_UPLOAD_RESPONSE_BYTES, 'Upload fragment', deadlines),
        bytes.length,
        'Upload fragment',
      );
      if (nextOffset < highestConfirmedOffset) throw new Error('Upload fragment response regressed below the confirmed offset');
      if (nextOffset !== endExclusive) throw new Error('Upload fragment response did not confirm the sequential byte range');
      offset = nextOffset;
      highestConfirmedOffset = nextOffset;
      attemptsForOffset = 0;
      continue;
    }
    if (response?.status === 409) {
      await discardResponse(response, deadlines.cleanupMilliseconds);
      if (options.findCompletedItem) {
        const completed = await options.findCompletedItem();
        if (completed) return { itemId: completeItemId(completed, 'Upload reconciliation'), transport: 'reconciled' };
      }
      throw new DriveItemConflictError();
    }
    if (response && response.status !== 416 && !isRetryableUploadStatus(response.status)) {
      await discardResponse(response, deadlines.cleanupMilliseconds);
      throw new Error(`Upload fragment failed with HTTP ${response.status}`);
    }

    const backedOffBeforeStatus = response !== undefined && (response.status === 429 || response.status >= 500);
    if (response) {
      await discardResponse(response, deadlines.cleanupMilliseconds);
      if (backedOffBeforeStatus) await dependencies.sleep(retryDelay(response, attemptsForOffset));
    }

    let status: UploadState;
    try {
      status = await uploadStatus(uploadUrl, bytes.length, dependencies, consumeRequestAttempt, deadlines);
    } catch (error) {
      if (endExclusive === bytes.length && options.findCompletedItem) {
        const completed = await options.findCompletedItem();
        if (completed) return { itemId: completeItemId(completed, 'Upload reconciliation'), transport: 'reconciled' };
      }
      throw error;
    }
    reconciled = true;
    if (status.kind === 'complete') return { itemId: status.itemId, transport: 'reconciled' };
    if (status.kind === 'expired') {
      if (endExclusive === bytes.length && options.findCompletedItem) {
        const completed = await options.findCompletedItem();
        if (completed) return { itemId: completeItemId(completed, 'Upload reconciliation'), transport: 'reconciled' };
      }
      throw new Error('Upload session expired before completion');
    }
    if (status.nextOffset < highestConfirmedOffset) throw new Error('Upload-session status regressed below the confirmed offset');
    if (status.nextOffset > endExclusive) throw new Error('Upload-session status reported a non-sequential range');
    highestConfirmedOffset = Math.max(highestConfirmedOffset, status.nextOffset);
    if (status.nextOffset === bytes.length && options.findCompletedItem) {
      const completed = await options.findCompletedItem();
      if (completed) return { itemId: completeItemId(completed, 'Upload reconciliation'), transport: 'reconciled' };
    }
    if (status.nextOffset !== offset) {
      offset = status.nextOffset;
      attemptsForOffset = 0;
      continue;
    }
    attemptsForOffset += 1;
    if (attemptsForOffset >= MAX_UPLOAD_ATTEMPTS) throw new Error('Upload fragment retry limit reached after status reconciliation');
    if (!backedOffBeforeStatus) await dependencies.sleep(retryDelay(response, attemptsForOffset - 1));
  }
  throw new Error('Upload session ended without a completed drive item');
}

const targetWriteQueues = new Map<string, Promise<unknown>>();

export async function serializeDriveTargetWrite<T>(targetKey: string, operation: () => Promise<T>): Promise<T> {
  const previous = targetWriteQueues.get(targetKey) ?? Promise.resolve();
  const current = previous.catch(() => undefined).then(operation);
  targetWriteQueues.set(targetKey, current);
  try {
    return await current;
  } finally {
    if (targetWriteQueues.get(targetKey) === current) targetWriteQueues.delete(targetKey);
  }
}

async function downloadDriveItemBytes(
  config: AuthConfig,
  itemId: string,
  expectedBytes: number,
  driveId: string | undefined,
  dependencies: DriveUploadDependencies,
): Promise<Buffer> {
  const deadlines = resolveTransportDeadlines(DRIVE_TRANSPORT_DEADLINES, dependencies.deadlines);
  const token = await dependencies.token(config);
  const response = await fetchWithDeadline(
    dependencies.fetch,
    `${GRAPH_BASE}${driveItemPathById(itemId, driveId)}/content`,
    { headers: { Authorization: `Bearer ${token}` } },
    'Uploaded file verification request',
    deadlines,
  );
  if (!response.ok) {
    await discardResponse(response, deadlines.cleanupMilliseconds);
    throw new Error(`Uploaded file verification download failed with HTTP ${response.status}`);
  }
  const downloaded = await readBoundedResponse(
    response,
    Math.min(expectedBytes + 1, MAX_SAFE_UPLOAD_BYTES),
    'Uploaded file verification download',
    deadlines,
  );
  if (downloaded.length !== expectedBytes) throw new Error('Uploaded file verification size mismatch');
  return downloaded;
}

interface ExpectedDriveFile {
  target: DrivePathTarget;
  parentId: string;
  driveId?: string;
  localBytes: Buffer;
  localSha256: string;
}

function parentReferenceMatchesTarget(
  item: DriveItemRecord,
  parentId: string,
  driveId: string | undefined,
): boolean {
  const reference = item.parentReference;
  return (reference?.id === undefined || reference.id === parentId)
    && (driveId === undefined || reference?.driveId === undefined || reference.driveId === driveId);
}

function parentReferencesAgree(left: DriveItemRecord, right: DriveItemRecord): boolean {
  const leftReference = left.parentReference;
  const rightReference = right.parentReference;
  return (leftReference?.id === undefined || rightReference?.id === undefined || leftReference.id === rightReference.id)
    && (leftReference?.driveId === undefined || rightReference?.driveId === undefined || leftReference.driveId === rightReference.driveId)
    && (leftReference?.path === undefined || rightReference?.path === undefined || leftReference.path === rightReference.path);
}

function isRequestedRegularFile(item: DriveItemRecord, expected: ExpectedDriveFile): boolean {
  return item.file !== undefined
    && item.folder === undefined
    && item.name === expected.target.fileName
    && item.size === expected.localBytes.length
    && parentReferenceMatchesTarget(item, expected.parentId, expected.driveId);
}

type VerificationMismatch = 'metadata' | 'content' | 'path';

async function verifyDriveItemAtRequestedPath(
  config: AuthConfig,
  candidate: DriveItemRecord,
  expected: ExpectedDriveFile,
  dependencies: DriveUploadDependencies,
  mismatchError: (mismatch: VerificationMismatch) => Error,
): Promise<DriveItemRecord> {
  if (!isRequestedRegularFile(candidate, expected)) throw mismatchError('metadata');
  let downloaded: Buffer;
  try {
    downloaded = await downloadDriveItemBytes(
      config,
      candidate.id,
      expected.localBytes.length,
      expected.driveId,
      dependencies,
    );
  } catch (error) {
    if (error instanceof Error && /maximum response size|verification size mismatch/u.test(error.message)) {
      throw mismatchError('content');
    }
    throw error;
  }
  const remoteSha256 = createHash('sha256').update(downloaded).digest('hex');
  if (remoteSha256 !== expected.localSha256 || !downloaded.equals(expected.localBytes)) {
    throw mismatchError('content');
  }

  const rebound = await getItemByPath(config, expected.target.path, expected.driveId, dependencies);
  if (
    !rebound
    || rebound.id !== candidate.id
    || !isRequestedRegularFile(rebound, expected)
    || !parentReferencesAgree(candidate, rebound)
  ) {
    throw mismatchError('path');
  }
  return candidate;
}

function verifiedDriveItem(item: DriveItemRecord, sha256: string, driveId?: string): VerifiedDriveItem {
  if (item.name === undefined || item.size === undefined) throw new Error('Verified drive item metadata was incomplete');
  return {
    id: item.id,
    name: item.name,
    size: item.size,
    ...(item.webUrl ? { webUrl: item.webUrl } : {}),
    ...(driveId ? { driveId } : {}),
    sha256,
    verified: true,
  };
}

async function findIdenticalTarget(
  config: AuthConfig,
  expected: ExpectedDriveFile,
  dependencies: DriveUploadDependencies,
): Promise<DriveItemRecord | null> {
  const candidate = await getItemByPath(config, expected.target.path, expected.driveId, dependencies);
  if (!candidate) return null;
  return verifyDriveItemAtRequestedPath(
    config,
    candidate,
    expected,
    dependencies,
    () => new DriveItemConflictError(),
  );
}

export async function safeCreateDriveFile(
  config: AuthConfig,
  input: {
    path: string;
    bytes: Uint8Array;
    driveId?: string;
    requiredExtension?: string;
  },
  dependencies: DriveUploadDependencies = defaultDependencies,
): Promise<VerifiedDriveItem> {
  const target = parseDriveFilePath(input.path, input.requiredExtension);
  assertDriveId(input.driveId);
  const localBytes = Buffer.from(input.bytes);
  if (localBytes.length === 0 || localBytes.length > MAX_SAFE_UPLOAD_BYTES) {
    throw new Error(`Upload must contain 1 through ${MAX_SAFE_UPLOAD_BYTES} bytes`);
  }
  const localSha256 = createHash('sha256').update(localBytes).digest('hex');
  const targetKey = `${input.driveId ?? 'me'}:${target.path.toLocaleLowerCase('en-US')}`;

  return serializeDriveTargetWrite(targetKey, async () => {
    const parent = await ensureDriveParentFolder(config, target.parentSegments, input.driveId, dependencies);
    const expected: ExpectedDriveFile = {
      target,
      parentId: parent.id,
      driveId: input.driveId,
      localBytes,
      localSha256,
    };
    const existing = await findIdenticalTarget(config, expected, dependencies);
    if (existing) return verifiedDriveItem(existing, localSha256, input.driveId);

    let sessionPayload: unknown;
    try {
      sessionPayload = await dependencies.graph(config, driveUploadSessionPath(parent.id, target.fileName, input.driveId), {
        method: 'POST',
        body: JSON.stringify({
          item: {
            '@microsoft.graph.conflictBehavior': 'fail',
            name: target.fileName,
          },
        }),
      });
    } catch (error) {
      const conflict = error instanceof GraphRequestError && error.status === 409;
      if (conflict || isAmbiguousWriteError(error)) {
        const reconciled = await findIdenticalTarget(config, expected, dependencies);
        if (reconciled) return verifiedDriveItem(reconciled, localSha256, input.driveId);
        if (conflict) throw new DriveItemConflictError();
        throw new Error('Upload-session creation had an ambiguous result and no identical target could be verified; rerun to reconcile safely');
      }
      throw error;
    }
    const uploadUrl = parseUploadUrl(sessionPayload);
    const uploaded = await uploadBytesToSession(uploadUrl, localBytes, {
      dependencies,
      findCompletedItem: () => findIdenticalTarget(config, expected, dependencies),
    });

    const metadataPath = `${driveItemPathById(uploaded.itemId, input.driveId)}?${ITEM_SELECT}`;
    const metadata = parseDriveItem(await dependencies.graph(config, metadataPath), 'Uploaded item verification');
    const verified = await verifyDriveItemAtRequestedPath(
      config,
      metadata,
      expected,
      dependencies,
      (mismatch) => new Error(mismatch === 'content'
        ? 'Uploaded file verification hash or byte comparison failed'
        : 'Uploaded item did not match the requested regular file at the requested Graph path'),
    );
    return verifiedDriveItem(verified, localSha256, input.driveId);
  });
}
