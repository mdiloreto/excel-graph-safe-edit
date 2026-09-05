import { type AuthConfig, GRAPH_BASE } from './config.js';
import { getAccessToken } from './auth.js';

export const MAX_GRAPH_RESPONSE_BYTES = 8 * 1024 * 1024;

export interface TransportDeadlines {
  requestMilliseconds: number;
  bodyMilliseconds: number;
  cleanupMilliseconds: number;
}

export const DEFAULT_GRAPH_TRANSPORT_DEADLINES: TransportDeadlines = {
  requestMilliseconds: 30_000,
  bodyMilliseconds: 30_000,
  cleanupMilliseconds: 1000,
};

export class GraphRequestError extends Error {
  constructor(
    message: string,
    readonly status: number,
  ) {
    super(message);
    this.name = 'GraphRequestError';
  }
}

export function isAmbiguousWriteError(error: unknown): boolean {
  return !(error instanceof GraphRequestError) || error.status === 408 || error.status === 429 || error.status >= 500;
}

function assertDeadline(value: number, field: string): number {
  if (!Number.isSafeInteger(value) || value < 1 || value > 30 * 60 * 1000) {
    throw new Error(`${field} must be an integer from 1 through 1800000 milliseconds`);
  }
  return value;
}

export function resolveTransportDeadlines(
  defaults: TransportDeadlines,
  overrides: Partial<TransportDeadlines> = {},
): TransportDeadlines {
  return {
    requestMilliseconds: assertDeadline(
      overrides.requestMilliseconds ?? defaults.requestMilliseconds,
      'Transport request deadline',
    ),
    bodyMilliseconds: assertDeadline(
      overrides.bodyMilliseconds ?? defaults.bodyMilliseconds,
      'Transport body deadline',
    ),
    cleanupMilliseconds: assertDeadline(
      overrides.cleanupMilliseconds ?? defaults.cleanupMilliseconds,
      'Transport cleanup deadline',
    ),
  };
}

async function settleWithin(promise: Promise<unknown>, milliseconds: number): Promise<void> {
  await new Promise<void>((resolve) => {
    let settled = false;
    const finish = (): void => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      resolve();
    };
    const timer = setTimeout(finish, milliseconds);
    void promise.then(finish, finish);
  });
}

export async function discardResponse(response: Response, cleanupMilliseconds = DEFAULT_GRAPH_TRANSPORT_DEADLINES.cleanupMilliseconds): Promise<void> {
  if (!response.body) return;
  let cancellation: Promise<unknown>;
  try {
    cancellation = response.body.cancel();
  } catch {
    return;
  }
  await settleWithin(cancellation, cleanupMilliseconds);
}

export async function fetchWithDeadline(
  fetcher: (url: string, options?: RequestInit) => Promise<Response>,
  url: string,
  options: RequestInit,
  context: string,
  deadlines: TransportDeadlines,
): Promise<Response> {
  const controller = new AbortController();
  const signal = options.signal
    ? AbortSignal.any([options.signal, controller.signal])
    : controller.signal;
  const request = Promise.resolve().then(() => fetcher(url, { ...options, signal }));
  return new Promise<Response>((resolve, reject) => {
    let settled = false;
    const timer = setTimeout(() => {
      if (settled) return;
      settled = true;
      reject(new Error(`${context} timed out`));
      controller.abort();
    }, deadlines.requestMilliseconds);
    void request.then(
      (response) => {
        if (settled) {
          void discardResponse(response, deadlines.cleanupMilliseconds);
          return;
        }
        settled = true;
        clearTimeout(timer);
        resolve(response);
      },
      (error: unknown) => {
        if (settled) return;
        settled = true;
        clearTimeout(timer);
        reject(error);
      },
    );
  });
}

export async function consumeResponseBody<T>(
  response: Response,
  context: string,
  deadlines: TransportDeadlines,
  consumer: (reader: ReadableStreamDefaultReader<Uint8Array>) => Promise<T>,
): Promise<T> {
  if (!response.body) throw new Error(`${context} did not include a response body`);
  const reader = response.body.getReader();
  const operation = Promise.resolve().then(() => consumer(reader));
  const deadline = new Promise<never>((_resolve, reject) => {
    const timer = setTimeout(() => reject(new Error(`${context} body timed out`)), deadlines.bodyMilliseconds);
    void operation.then(() => clearTimeout(timer), () => clearTimeout(timer));
  });
  try {
    return await Promise.race([operation, deadline]);
  } catch (error) {
    let cancellation: Promise<unknown>;
    try {
      cancellation = reader.cancel();
    } catch {
      throw error;
    }
    await settleWithin(cancellation, deadlines.cleanupMilliseconds);
    throw error;
  } finally {
    try {
      reader.releaseLock();
    } catch {
      // A non-cooperative injected stream may retain a pending read after the deadline.
    }
  }
}

export async function readBoundedResponse(
  response: Response,
  maximumBytes: number,
  context: string,
  deadlines: TransportDeadlines = DEFAULT_GRAPH_TRANSPORT_DEADLINES,
): Promise<Buffer> {
  const contentLength = response.headers.get('content-length');
  if (contentLength && /^\d+$/u.test(contentLength) && BigInt(contentLength) > BigInt(maximumBytes)) {
    await discardResponse(response, deadlines.cleanupMilliseconds);
    throw new Error(`${context} exceeded the maximum response size`);
  }
  if (!response.body) return Buffer.alloc(0);
  return consumeResponseBody(response, context, deadlines, async (reader) => {
    const chunks: Buffer[] = [];
    let bytes = 0;
    while (true) {
      const result = await reader.read();
      if (result.done) break;
      const chunk = Buffer.from(result.value);
      bytes += chunk.length;
      if (bytes > maximumBytes) {
        throw new Error(`${context} exceeded the maximum response size`);
      }
      chunks.push(chunk);
    }
    return Buffer.concat(chunks, bytes);
  });
}

export async function readBoundedJson(
  response: Response,
  maximumBytes: number,
  context: string,
  deadlines: TransportDeadlines = DEFAULT_GRAPH_TRANSPORT_DEADLINES,
): Promise<unknown> {
  const bytes = await readBoundedResponse(response, maximumBytes, context, deadlines);
  if (bytes.length === 0) return {};
  let text: string;
  try {
    text = new TextDecoder('utf-8', { fatal: true }).decode(bytes);
  } catch {
    throw new Error(`${context} returned invalid UTF-8`);
  }
  try {
    return JSON.parse(text) as unknown;
  } catch {
    throw new Error(`${context} returned invalid JSON`);
  }
}

export async function graphRequest<T = unknown>(
  config: AuthConfig,
  path: string,
  options: RequestInit = {},
  deadlineOverrides: Partial<TransportDeadlines> = {},
): Promise<T> {
  const deadlines = resolveTransportDeadlines(DEFAULT_GRAPH_TRANSPORT_DEADLINES, deadlineOverrides);
  const token = await getAccessToken(config);
  const response = await fetchWithDeadline((url, requestOptions) => fetch(url, requestOptions), `${GRAPH_BASE}${path}`, {
    ...options,
    headers: {
      ...(options.body ? { 'Content-Type': 'application/json' } : {}),
      ...(options.headers ?? {}),
      Authorization: `Bearer ${token}`,
    },
  }, 'Graph request', deadlines);
  const bytes = await readBoundedResponse(response, MAX_GRAPH_RESPONSE_BYTES, 'Graph response', deadlines);
  if (!response.ok) {
    throw new GraphRequestError(`Graph request failed with HTTP ${response.status}`, response.status);
  }
  const contentType = response.headers.get('content-type') ?? '';
  let text: string;
  try {
    text = new TextDecoder('utf-8', { fatal: true }).decode(bytes);
  } catch {
    throw new Error('Graph response returned invalid UTF-8');
  }
  let payload: unknown = text;
  if (contentType.includes('application/json')) {
    try {
      payload = JSON.parse(text) as unknown;
    } catch {
      throw new Error('Graph response returned invalid JSON');
    }
  }
  return payload as T;
}
