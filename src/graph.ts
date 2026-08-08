import { type AuthConfig, GRAPH_BASE } from './config.js';
import { getAccessToken } from './auth.js';

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

export async function graphRequest<T = unknown>(config: AuthConfig, path: string, options: RequestInit = {}): Promise<T> {
  const token = await getAccessToken(config);
  const response = await fetch(`${GRAPH_BASE}${path}`, {
    ...options,
    headers: {
      ...(options.body ? { 'Content-Type': 'application/json' } : {}),
      ...(options.headers ?? {}),
      Authorization: `Bearer ${token}`,
    },
  });
  const contentType = response.headers.get('content-type') ?? '';
  const payload = contentType.includes('application/json') ? await response.json() : await response.text();
  if (!response.ok) {
    throw new GraphRequestError(
      `Graph request failed ${response.status}: ${typeof payload === 'string' ? payload : JSON.stringify(payload)}`,
      response.status,
    );
  }
  return payload as T;
}
