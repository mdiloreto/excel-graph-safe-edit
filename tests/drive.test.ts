import { createHash } from 'node:crypto';
import { describe, expect, it, vi } from 'vitest';
import { buildAuthConfig, type AuthConfig } from '../src/config.js';
import {
  DriveItemConflictError,
  ensureDriveParentFolder,
  GRAPH_UPLOAD_GRANULARITY,
  MAX_SAFE_UPLOAD_BYTES,
  parseDriveFilePath,
  safeCreateDriveFile,
  serializeDriveTargetWrite,
  uploadBytesToSession,
  driveItemPathByPath,
  driveUploadSessionPath,
  type DriveUploadDependencies,
} from '../src/drive.js';
import { GraphRequestError } from '../src/graph.js';

const authConfig = buildAuthConfig({ clientId: 'client-id', noPersist: true });
const uploadUrl = 'https://uploads.example.test/preauthenticated-secret';

function jsonResponse(value: unknown, status: number): Response {
  return new Response(JSON.stringify(value), {
    status,
    headers: { 'Content-Type': 'application/json' },
  });
}

function dependencies(overrides: Partial<DriveUploadDependencies> = {}): DriveUploadDependencies {
  return {
    graph: async () => ({}),
    token: async () => 'bearer-secret',
    fetch: async () => { throw new Error('Unexpected fetch'); },
    sleep: async () => undefined,
    ...overrides,
  };
}

describe('generic drive path helpers', () => {
  it('validates target paths and escapes each Graph path segment', () => {
    expect(parseDriveFilePath('/Class notes/Week #1 % complete (final).docx', '.docx')).toEqual({
      path: 'Class notes/Week #1 % complete (final).docx',
      parentSegments: ['Class notes'],
      fileName: 'Week #1 % complete (final).docx',
    });
    expect(driveItemPathByPath('Class notes/Week #1 % complete (final).docx', 'drive/id')).toBe(
      '/drives/drive%2Fid/root:/Class%20notes/Week%20%231%20%25%20complete%20%28final%29.docx:',
    );
    expect(driveUploadSessionPath('parent/id', 'Week #1.docx', 'drive/id')).toBe(
      '/drives/drive%2Fid/items/parent%2Fid:/Week%20%231.docx:/createUploadSession',
    );
  });

  it('rejects ambiguous, reserved, invalid, and non-DOCX target names', () => {
    for (const path of [
      '',
      'Folder//file.docx',
      'Folder/../file.docx',
      'Folder/bad?.docx',
      'Folder/trailing .docx ',
      'CON.docx',
      '~temporary.docx',
      'Forms/file.docx',
      'Folder/file.txt',
    ]) {
      expect(() => parseDriveFilePath(path, '.docx')).toThrow();
    }
  });
});

describe('folder ensure and upload-session safety', () => {
  it('accepts a root folder size above the upload limit', async () => {
    const graph = vi.fn(async () => ({
      id: 'root',
      name: 'root',
      size: 150 * 1024 * 1024,
      folder: {},
    }));

    await expect(ensureDriveParentFolder(authConfig, [], undefined, dependencies({ graph }))).resolves.toEqual({ id: 'root' });
  });

  it('reconciles a folder create conflict without renaming or retrying the write', async () => {
    let lookupCount = 0;
    const graph = vi.fn(async (_config: AuthConfig, path: string, options?: RequestInit): Promise<unknown> => {
      if (path.startsWith('/me/drive/root?')) return { id: 'root', name: 'root', folder: {} };
      if (path.includes('root:/Class%20notes:')) {
        lookupCount += 1;
        if (lookupCount === 1) throw new GraphRequestError('not found', 404);
        return { id: 'folder', name: 'Class notes', folder: {} };
      }
      if (options?.method === 'POST') throw new GraphRequestError('conflict', 409);
      throw new Error(`Unexpected graph path ${path}`);
    });

    await expect(ensureDriveParentFolder(authConfig, ['Class notes'], undefined, dependencies({ graph }))).resolves.toEqual({ id: 'folder' });
    const createCalls = graph.mock.calls.filter((call) => call[2]?.method === 'POST');
    expect(createCalls).toHaveLength(1);
    expect(JSON.parse(String(createCalls[0]?.[2]?.body))).toEqual({
      name: 'Class notes',
      folder: {},
      '@microsoft.graph.conflictBehavior': 'fail',
    });
  });

  it('uploads sequential Graph-valid chunks and never sends bearer auth to the upload URL', async () => {
    const bytes = Buffer.alloc(GRAPH_UPLOAD_GRANULARITY + 17, 7);
    const calls: RequestInit[] = [];
    const fetchUpload = vi.fn(async (_url: string, options?: RequestInit) => {
      calls.push(options ?? {});
      if (calls.length === 1) {
        return jsonResponse({ nextExpectedRanges: [`${GRAPH_UPLOAD_GRANULARITY}-`] }, 202);
      }
      return jsonResponse({ id: 'uploaded-item' }, 201);
    });

    await expect(uploadBytesToSession(uploadUrl, bytes, {
      dependencies: dependencies({ fetch: fetchUpload }),
      chunkBytes: GRAPH_UPLOAD_GRANULARITY,
    })).resolves.toEqual({ itemId: 'uploaded-item', transport: 'confirmed' });

    expect(calls.map((call) => new Headers(call.headers).get('content-range'))).toEqual([
      `bytes 0-${GRAPH_UPLOAD_GRANULARITY - 1}/${bytes.length}`,
      `bytes ${GRAPH_UPLOAD_GRANULARITY}-${bytes.length - 1}/${bytes.length}`,
    ]);
    for (const call of calls) expect(new Headers(call.headers).has('authorization')).toBe(false);
  });

  it('reconciles an ambiguous accepted chunk through upload-session status before continuing', async () => {
    const bytes = Buffer.alloc(GRAPH_UPLOAD_GRANULARITY + 1, 3);
    const methods: string[] = [];
    const fetchUpload = vi.fn(async (_url: string, options?: RequestInit) => {
      const method = options?.method ?? 'GET';
      methods.push(method);
      if (methods.length === 1) throw new TypeError('socket closed');
      if (method === 'GET') return jsonResponse({ nextExpectedRanges: [`${GRAPH_UPLOAD_GRANULARITY}-`] }, 200);
      return jsonResponse({ id: 'uploaded-item' }, 201);
    });

    await expect(uploadBytesToSession(uploadUrl, bytes, {
      dependencies: dependencies({ fetch: fetchUpload }),
      chunkBytes: GRAPH_UPLOAD_GRANULARITY,
    })).resolves.toEqual({ itemId: 'uploaded-item', transport: 'reconciled' });
    expect(methods).toEqual(['PUT', 'GET', 'PUT']);
    for (const call of fetchUpload.mock.calls) expect(new Headers(call[1]?.headers).has('authorization')).toBe(false);
  });

  it('rejects alternating status offsets below the highest confirmed offset', async () => {
    const bytes = Buffer.alloc(GRAPH_UPLOAD_GRANULARITY + 1, 3);
    const methods: string[] = [];
    const fetchUpload = vi.fn(async (_url: string, options?: RequestInit) => {
      const method = options?.method ?? 'GET';
      methods.push(method);
      if (method === 'PUT') throw new TypeError('socket closed');
      return jsonResponse({
        nextExpectedRanges: [methods.filter((entry) => entry === 'GET').length === 1 ? `${GRAPH_UPLOAD_GRANULARITY}-` : '0-'],
      }, 200);
    });

    await expect(uploadBytesToSession(uploadUrl, bytes, {
      dependencies: dependencies({ fetch: fetchUpload }),
      chunkBytes: GRAPH_UPLOAD_GRANULARITY,
    })).rejects.toThrow(/regressed below the confirmed offset/);
    expect(methods).toEqual(['PUT', 'GET', 'PUT', 'GET']);
  });

  it('caps total PUT and status attempts across the entire upload session', async () => {
    const bytes = Buffer.alloc(MAX_SAFE_UPLOAD_BYTES, 4);
    let confirmedOffset = 0;
    let failedPutsAtOffset = 0;
    const fetchUpload = vi.fn(async (_url: string, options?: RequestInit) => {
      const method = options?.method ?? 'GET';
      if (method === 'GET') return jsonResponse({ nextExpectedRanges: [`${confirmedOffset}-`] }, 200);
      if (failedPutsAtOffset < 3) {
        failedPutsAtOffset += 1;
        throw new TypeError('socket closed');
      }
      failedPutsAtOffset = 0;
      confirmedOffset = Math.min(confirmedOffset + GRAPH_UPLOAD_GRANULARITY, bytes.length);
      return jsonResponse({ nextExpectedRanges: [`${confirmedOffset}-`] }, 202);
    });

    await expect(uploadBytesToSession(uploadUrl, bytes, {
      dependencies: dependencies({ fetch: fetchUpload }),
      chunkBytes: GRAPH_UPLOAD_GRANULARITY,
    })).rejects.toThrow(/request limit reached/);
    expect(fetchUpload).toHaveBeenCalledTimes(256);
  });

  it('bounds never-resolving upload PUT and status requests even when fetch ignores abort', async () => {
    const signals: AbortSignal[] = [];
    const fetchUpload = vi.fn((_url: string, options?: RequestInit) => {
      if (options?.signal) signals.push(options.signal);
      return new Promise<Response>(() => undefined);
    });

    let error: unknown;
    try {
      await uploadBytesToSession(uploadUrl, Buffer.from('docx'), {
        dependencies: dependencies({
          fetch: fetchUpload,
          deadlines: { requestMilliseconds: 10, bodyMilliseconds: 10, cleanupMilliseconds: 5 },
        }),
      });
    } catch (caught) {
      error = caught;
    }
    expect(error).toBeInstanceOf(Error);
    expect((error as Error).message).toMatch(/status could not be reconciled/);
    expect((error as Error).message).not.toContain(uploadUrl);
    expect(fetchUpload).toHaveBeenCalledTimes(5);
    expect(signals).toHaveLength(5);
    expect(signals.every((signal) => signal.aborted)).toBe(true);
  });

  it('cancels a slow-drip upload response body at its deadline', async () => {
    let cancelled = false;
    const body = new ReadableStream<Uint8Array>({
      start: (controller) => controller.enqueue(Buffer.from('{"nextExpectedRanges":')),
      cancel: () => { cancelled = true; },
    });
    const fetchUpload = vi.fn(async () => new Response(body, {
      status: 202,
      headers: { 'Content-Type': 'application/json' },
    }));

    await expect(uploadBytesToSession(uploadUrl, Buffer.from('docx'), {
      dependencies: dependencies({
        fetch: fetchUpload,
        deadlines: { requestMilliseconds: 50, bodyMilliseconds: 10, cleanupMilliseconds: 5 },
      }),
    })).rejects.toThrow(/^Upload fragment body timed out$/u);
    expect(cancelled).toBe(true);
  });

  it('backs off and retries a 5xx fragment only after status proves the range is still missing', async () => {
    const methods: string[] = [];
    const sleep = vi.fn(async () => undefined);
    const fetchUpload = vi.fn(async (_url: string, options?: RequestInit) => {
      const method = options?.method ?? 'GET';
      methods.push(method);
      if (methods.length === 1) return jsonResponse({ error: 'temporary' }, 503);
      if (method === 'GET') return jsonResponse({ nextExpectedRanges: ['0-'] }, 200);
      return jsonResponse({ id: 'uploaded-item' }, 201);
    });

    await expect(uploadBytesToSession(uploadUrl, Buffer.from('docx'), {
      dependencies: dependencies({ fetch: fetchUpload, sleep }),
    })).resolves.toEqual({ itemId: 'uploaded-item', transport: 'reconciled' });
    expect(methods).toEqual(['PUT', 'GET', 'PUT']);
    expect(sleep).toHaveBeenCalledOnce();
  });

  it('reconciles an ambiguous final response by target lookup when the completed session is gone', async () => {
    const fetchUpload = vi.fn()
      .mockRejectedValueOnce(new TypeError('socket closed'))
      .mockResolvedValueOnce(jsonResponse({ error: 'session gone' }, 404));
    const findCompletedItem = vi.fn(async () => ({ id: 'uploaded-item' }));

    await expect(uploadBytesToSession(uploadUrl, Buffer.from('docx'), {
      dependencies: dependencies({ fetch: fetchUpload }),
      findCompletedItem,
    })).resolves.toEqual({ itemId: 'uploaded-item', transport: 'reconciled' });
    expect(findCompletedItem).toHaveBeenCalledOnce();
  });

  it('does not retry a non-ambiguous fragment conflict', async () => {
    const fetchUpload = vi.fn(async () => jsonResponse({ error: 'nameAlreadyExists' }, 409));
    await expect(uploadBytesToSession(uploadUrl, Buffer.from('docx'), {
      dependencies: dependencies({ fetch: fetchUpload }),
    })).rejects.toBeInstanceOf(DriveItemConflictError);
    expect(fetchUpload).toHaveBeenCalledOnce();
  });

  it('does not retry or query status for a non-retryable client error', async () => {
    const fetchUpload = vi.fn(async () => jsonResponse({ error: 'bad request' }, 400));
    await expect(uploadBytesToSession(uploadUrl, Buffer.from('docx'), {
      dependencies: dependencies({ fetch: fetchUpload }),
    })).rejects.toThrow(/HTTP 400/);
    expect(fetchUpload).toHaveBeenCalledOnce();
  });
});

describe('create-or-verify-identical upload verification', () => {
  it('uses conflictBehavior fail, rereads by returned id, redownloads, and returns allowlisted verified metadata', async () => {
    const bytes = Buffer.from('bounded-docx-bytes');
    let targetLookupCount = 0;
    const graph = vi.fn(async (_config: AuthConfig, path: string, options?: RequestInit): Promise<unknown> => {
      if (path.startsWith('/drives/drive/root?')) return { id: 'root', name: 'root', folder: {} };
      if (path.includes('root:/Class%20notes:')) return { id: 'folder', name: 'Class notes', folder: {} };
      if (path.startsWith('/drives/drive/root:/Class%20notes/transcript.docx:')) {
        targetLookupCount += 1;
        if (targetLookupCount === 1) throw new GraphRequestError('not found', 404);
        return {
          id: 'uploaded',
          name: 'transcript.docx',
          size: bytes.length,
          file: {},
          parentReference: { id: 'folder', driveId: 'drive', path: '/drive/root:/Class notes' },
        };
      }
      if (path.endsWith(':/transcript.docx:/createUploadSession') && options?.method === 'POST') return { uploadUrl };
      if (path.startsWith('/drives/drive/items/uploaded?')) {
        return {
          id: 'uploaded',
          name: 'transcript.docx',
          size: bytes.length,
          webUrl: 'https://onedrive.example.test/transcript.docx',
          file: { mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' },
          parentReference: { id: 'folder', driveId: 'drive', path: '/drive/root:/Class notes' },
          ignoredSecret: 'not returned',
        };
      }
      throw new Error(`Unexpected graph path ${path}`);
    });
    const fetchMock = vi.fn(async (url: string, options?: RequestInit) => {
      if (url === uploadUrl) {
        expect(new Headers(options?.headers).has('authorization')).toBe(false);
        return jsonResponse({ id: 'uploaded' }, 201);
      }
      expect(url).toContain('/drives/drive/items/uploaded/content');
      expect(new Headers(options?.headers).get('authorization')).toBe('Bearer bearer-secret');
      return new Response(bytes, { status: 200 });
    });

    await expect(safeCreateDriveFile(authConfig, {
      path: 'Class notes/transcript.docx',
      bytes,
      driveId: 'drive',
      requiredExtension: '.docx',
    }, dependencies({ graph, fetch: fetchMock }))).resolves.toEqual({
      id: 'uploaded',
      name: 'transcript.docx',
      size: bytes.length,
      webUrl: 'https://onedrive.example.test/transcript.docx',
      driveId: 'drive',
      sha256: createHash('sha256').update(bytes).digest('hex'),
      verified: true,
    });

    const sessionCall = graph.mock.calls.find((call) => call[1].endsWith('/createUploadSession'));
    expect(JSON.parse(String(sessionCall?.[2]?.body))).toEqual({
      item: {
        '@microsoft.graph.conflictBehavior': 'fail',
        name: 'transcript.docx',
      },
    });
    expect(graph.mock.calls.some((call) => call[1].startsWith('/drives/drive/items/uploaded?'))).toBe(true);
    expect(targetLookupCount).toBe(2);
  });

  it('adopts an identical regular file already at the exact target without issuing a write', async () => {
    const bytes = Buffer.from('existing-identical');
    const item = {
      id: 'existing',
      name: 'transcript.vtt',
      size: bytes.length,
      webUrl: 'https://onedrive.example.test/transcript.vtt',
      file: {},
      parentReference: { id: 'root', path: '/drive/root:' },
    };
    const graph = vi.fn(async (_config: AuthConfig, path: string, options?: RequestInit): Promise<unknown> => {
      expect(options?.method).toBeUndefined();
      if (path.startsWith('/me/drive/root?')) return { id: 'root', folder: {} };
      if (path.startsWith('/me/drive/root:/transcript.vtt:')) return item;
      throw new Error(`Unexpected graph path ${path}`);
    });
    const fetchMock = vi.fn(async () => new Response(bytes, { status: 200 }));

    await expect(safeCreateDriveFile(authConfig, {
      path: 'transcript.vtt',
      bytes,
    }, dependencies({ graph, fetch: fetchMock }))).resolves.toEqual({
      id: 'existing',
      name: 'transcript.vtt',
      size: bytes.length,
      webUrl: 'https://onedrive.example.test/transcript.vtt',
      sha256: createHash('sha256').update(bytes).digest('hex'),
      verified: true,
    });
    expect(graph.mock.calls.some((call) => call[2]?.method !== undefined)).toBe(false);
    expect(fetchMock).toHaveBeenCalledOnce();
  });

  it('keeps a pre-existing target conflict when its downloaded bytes differ', async () => {
    const bytes = Buffer.from('local');
    const graph = vi.fn(async (_config: AuthConfig, path: string, options?: RequestInit): Promise<unknown> => {
      expect(options?.method).toBeUndefined();
      if (path.startsWith('/me/drive/root?')) return { id: 'root', folder: {} };
      if (path.startsWith('/me/drive/root:/transcript.vtt:')) {
        return { id: 'existing', name: 'transcript.vtt', size: bytes.length, file: {}, parentReference: { id: 'root' } };
      }
      throw new Error(`Unexpected graph path ${path}`);
    });
    const fetchMock = vi.fn(async () => new Response(Buffer.from('other'), { status: 200 }));

    await expect(safeCreateDriveFile(authConfig, {
      path: 'transcript.vtt',
      bytes,
    }, dependencies({ graph, fetch: fetchMock }))).rejects.toBeInstanceOf(DriveItemConflictError);
    expect(graph.mock.calls.some((call) => call[2]?.method !== undefined)).toBe(false);
  });

  it.each([
    ['conflict', new GraphRequestError('conflict', 409)],
    ['ambiguous transport failure', new TypeError('socket closed')],
  ])('adopts identical content after upload-session creation %s without retrying a write', async (_label, creationError) => {
    const bytes = Buffer.from('reconciled-identical');
    let targetLookups = 0;
    const item = {
      id: 'existing',
      name: 'transcript.vtt',
      size: bytes.length,
      file: {},
      parentReference: { id: 'root' },
    };
    const graph = vi.fn(async (_config: AuthConfig, path: string, options?: RequestInit): Promise<unknown> => {
      if (path.startsWith('/me/drive/root?')) return { id: 'root', folder: {} };
      if (path.startsWith('/me/drive/root:/transcript.vtt:')) {
        targetLookups += 1;
        if (targetLookups === 1) throw new GraphRequestError('not found', 404);
        return item;
      }
      if (options?.method === 'POST') throw creationError;
      throw new Error(`Unexpected graph path ${path}`);
    });
    const fetchMock = vi.fn(async () => new Response(bytes, { status: 200 }));

    await expect(safeCreateDriveFile(authConfig, {
      path: 'transcript.vtt',
      bytes,
    }, dependencies({ graph, fetch: fetchMock }))).resolves.toMatchObject({ id: 'existing', verified: true });
    expect(graph.mock.calls.filter((call) => call[2]?.method === 'POST')).toHaveLength(1);
    expect(fetchMock).toHaveBeenCalledOnce();
  });

  it('adopts only exact bytes when final upload completion is ambiguous', async () => {
    const bytes = Buffer.from('final-identical');
    let targetLookups = 0;
    const item = {
      id: 'uploaded',
      name: 'transcript.vtt',
      size: bytes.length,
      file: {},
      parentReference: { id: 'root' },
    };
    const graph = vi.fn(async (_config: AuthConfig, path: string, options?: RequestInit): Promise<unknown> => {
      if (path.startsWith('/me/drive/root?')) return { id: 'root', folder: {} };
      if (path.startsWith('/me/drive/root:/transcript.vtt:')) {
        targetLookups += 1;
        if (targetLookups === 1) throw new GraphRequestError('not found', 404);
        return item;
      }
      if (options?.method === 'POST') return { uploadUrl };
      if (path.startsWith('/me/drive/items/uploaded?')) return item;
      throw new Error(`Unexpected graph path ${path}`);
    });
    const uploadMethods: string[] = [];
    const fetchMock = vi.fn(async (url: string, options?: RequestInit) => {
      if (url !== uploadUrl) return new Response(bytes, { status: 200 });
      const method = options?.method ?? 'GET';
      uploadMethods.push(method);
      throw new TypeError('socket closed');
    });

    await expect(safeCreateDriveFile(authConfig, {
      path: 'transcript.vtt',
      bytes,
    }, dependencies({ graph, fetch: fetchMock }))).resolves.toMatchObject({ id: 'uploaded', verified: true });
    expect(uploadMethods).toEqual(['PUT', 'GET', 'GET', 'GET', 'GET']);
    expect(targetLookups).toBe(4);
  });

  it('fails when the uploaded item is replaced at the requested path before final verification', async () => {
    const bytes = Buffer.from('bounded-docx-bytes');
    let targetLookups = 0;
    const graph = vi.fn(async (_config: AuthConfig, path: string, options?: RequestInit): Promise<unknown> => {
      if (path.startsWith('/me/drive/root?')) return { id: 'root', folder: {} };
      if (path.startsWith('/me/drive/root:/transcript.docx:')) {
        targetLookups += 1;
        if (targetLookups === 1) throw new GraphRequestError('not found', 404);
        return { id: 'replacement', name: 'transcript.docx', size: bytes.length, file: {}, parentReference: { id: 'root' } };
      }
      if (options?.method === 'POST') return { uploadUrl };
      if (path.startsWith('/me/drive/items/uploaded?')) {
        return { id: 'uploaded', name: 'transcript.docx', size: bytes.length, file: {}, parentReference: { id: 'root' } };
      }
      throw new Error(`Unexpected graph path ${path}`);
    });
    const fetchMock = vi.fn(async (url: string) => (
      url === uploadUrl ? jsonResponse({ id: 'uploaded' }, 201) : new Response(bytes, { status: 200 })
    ));

    await expect(safeCreateDriveFile(authConfig, {
      path: 'transcript.docx',
      bytes,
      requiredExtension: '.docx',
    }, dependencies({ graph, fetch: fetchMock }))).rejects.toThrow(/requested Graph path/);
  });

  it('fails when the uploaded item is moved away from the requested path before final verification', async () => {
    const bytes = Buffer.from('bounded-docx-bytes');
    const graph = vi.fn(async (_config: AuthConfig, path: string, options?: RequestInit): Promise<unknown> => {
      if (path.startsWith('/me/drive/root?')) return { id: 'root', folder: {} };
      if (path.startsWith('/me/drive/root:/transcript.docx:')) throw new GraphRequestError('not found', 404);
      if (options?.method === 'POST') return { uploadUrl };
      if (path.startsWith('/me/drive/items/uploaded?')) {
        return { id: 'uploaded', name: 'transcript.docx', size: bytes.length, file: {}, parentReference: { id: 'root' } };
      }
      throw new Error(`Unexpected graph path ${path}`);
    });
    const fetchMock = vi.fn(async (url: string) => (
      url === uploadUrl ? jsonResponse({ id: 'uploaded' }, 201) : new Response(bytes, { status: 200 })
    ));

    await expect(safeCreateDriveFile(authConfig, {
      path: 'transcript.docx',
      bytes,
      requiredExtension: '.docx',
    }, dependencies({ graph, fetch: fetchMock }))).rejects.toThrow(/requested Graph path/);
  });

  it('fails closed when redownloaded bytes do not match the local SHA-256', async () => {
    const bytes = Buffer.from('local-docx');
    const graph = vi.fn(async (_config: AuthConfig, path: string, options?: RequestInit): Promise<unknown> => {
      if (path.startsWith('/me/drive/root?')) return { id: 'root', folder: {} };
      if (path.startsWith('/me/drive/root:/transcript.docx:')) throw new GraphRequestError('not found', 404);
      if (options?.method === 'POST') return { uploadUrl };
      if (path.startsWith('/me/drive/items/uploaded?')) return { id: 'uploaded', name: 'transcript.docx', size: bytes.length, file: {} };
      throw new Error(`Unexpected graph path ${path}`);
    });
    const fetchMock = vi.fn(async (url: string) => (
      url === uploadUrl
        ? jsonResponse({ id: 'uploaded' }, 201)
        : new Response(Buffer.from('other-docx'), { status: 200 })
    ));

    await expect(safeCreateDriveFile(authConfig, {
      path: 'transcript.docx',
      bytes,
      requiredExtension: '.docx',
    }, dependencies({ graph, fetch: fetchMock }))).rejects.toThrow(/hash or byte comparison/);
  });

  it('aborts and bounds a verification content request when fetch never resolves', async () => {
    const bytes = Buffer.from('local-docx');
    const verificationSignals: AbortSignal[] = [];
    const graph = vi.fn(async (_config: AuthConfig, path: string, options?: RequestInit): Promise<unknown> => {
      if (path.startsWith('/me/drive/root?')) return { id: 'root', folder: {} };
      if (path.startsWith('/me/drive/root:/transcript.docx:')) throw new GraphRequestError('not found', 404);
      if (options?.method === 'POST') return { uploadUrl };
      if (path.startsWith('/me/drive/items/uploaded?')) {
        return { id: 'uploaded', name: 'transcript.docx', size: bytes.length, file: {}, parentReference: { id: 'root' } };
      }
      throw new Error(`Unexpected graph path ${path}`);
    });
    const fetchMock = vi.fn((url: string, options?: RequestInit): Promise<Response> => {
      if (url === uploadUrl) return Promise.resolve(jsonResponse({ id: 'uploaded' }, 201));
      if (options?.signal) verificationSignals.push(options.signal);
      return new Promise<Response>(() => undefined);
    });

    await expect(safeCreateDriveFile(authConfig, {
      path: 'transcript.docx',
      bytes,
      requiredExtension: '.docx',
    }, dependencies({
      graph,
      fetch: fetchMock,
      deadlines: { requestMilliseconds: 10, bodyMilliseconds: 10, cleanupMilliseconds: 5 },
    }))).rejects.toThrow(/^Uploaded file verification request timed out$/u);
    expect(verificationSignals).toHaveLength(1);
    expect(verificationSignals[0]?.aborted).toBe(true);
  });

  it('serializes writes to the same normalized target in process', async () => {
    const events: string[] = [];
    let release: (() => void) | undefined;
    const first = serializeDriveTargetWrite('me:folder/file.docx', async () => {
      events.push('first-start');
      await new Promise<void>((resolve) => { release = resolve; });
      events.push('first-end');
      return 1;
    });
    const second = serializeDriveTargetWrite('me:folder/file.docx', async () => {
      events.push('second-start');
      return 2;
    });
    await new Promise<void>((resolve) => setImmediate(resolve));
    expect(events).toEqual(['first-start']);
    release?.();
    await expect(Promise.all([first, second])).resolves.toEqual([1, 2]);
    expect(events).toEqual(['first-start', 'first-end', 'second-start']);
  });
});
