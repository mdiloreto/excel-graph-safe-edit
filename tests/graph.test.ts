import { afterEach, describe, expect, it, vi } from 'vitest';
import { buildAuthConfig } from '../src/config.js';
import { graphRequest, GraphRequestError, MAX_GRAPH_RESPONSE_BYTES, readBoundedResponse } from '../src/graph.js';

vi.mock('../src/auth.js', () => ({
  getAccessToken: vi.fn(async () => 'test-token'),
}));

const authConfig = buildAuthConfig({ clientId: 'client-id', noPersist: true });
const shortDeadlines = {
  requestMilliseconds: 15,
  bodyMilliseconds: 15,
  cleanupMilliseconds: 5,
};

afterEach(() => {
  vi.restoreAllMocks();
});

describe('bounded Graph responses', () => {
  it('parses a bounded authenticated JSON success response', async () => {
    const fetchMock = vi.spyOn(globalThis, 'fetch').mockImplementation(async (_url, options) => {
      expect(new Headers(options?.headers).get('authorization')).toBe('Bearer test-token');
      return new Response(JSON.stringify({ id: 'item' }), {
        status: 200,
        headers: { 'Content-Type': 'application/json' },
      });
    });

    await expect(graphRequest(authConfig, '/me/drive')).resolves.toEqual({ id: 'item' });
    expect(fetchMock).toHaveBeenCalledOnce();
  });

  it('cancels a success response whose declared size exceeds the bound', async () => {
    let cancelled = false;
    const body = new ReadableStream<Uint8Array>({
      cancel: () => { cancelled = true; },
    });
    vi.spyOn(globalThis, 'fetch').mockResolvedValue(new Response(body, {
      status: 200,
      headers: {
        'Content-Length': String(MAX_GRAPH_RESPONSE_BYTES + 1),
        'Content-Type': 'application/json',
      },
    }));

    await expect(graphRequest(authConfig, '/me/drive')).rejects.toThrow(/maximum response size/);
    expect(cancelled).toBe(true);
  });

  it('cancels a streamed response that crosses its byte limit', async () => {
    let cancelled = false;
    const body = new ReadableStream<Uint8Array>({
      start: (controller) => controller.enqueue(Buffer.from('four')),
      cancel: () => { cancelled = true; },
    });

    await expect(readBoundedResponse(new Response(body), 3, 'Test response')).rejects.toThrow(/maximum response size/);
    expect(cancelled).toBe(true);
  });

  it('bounds error responses without exposing raw bodies or terminal controls', async () => {
    const canary = 'server-secret\u001b[31mCANARY\u0007';
    vi.spyOn(globalThis, 'fetch').mockResolvedValue(new Response(canary, { status: 500 }));

    let error: unknown;
    try {
      await graphRequest(authConfig, '/me/drive');
    } catch (caught) {
      error = caught;
    }
    expect(error).toBeInstanceOf(GraphRequestError);
    expect(error).toMatchObject({ status: 500, message: 'Graph request failed with HTTP 500' });
    expect((error as Error).message).not.toContain('server-secret');
    expect(Array.from((error as Error).message).some((character) => {
      const codePoint = character.codePointAt(0) ?? 0;
      return codePoint < 0x20 || codePoint === 0x7f;
    })).toBe(false);
  });

  it('aborts and rejects when fetch never resolves even if it ignores the signal', async () => {
    const signals: AbortSignal[] = [];
    vi.spyOn(globalThis, 'fetch').mockImplementation((_url, options) => {
      if (options?.signal) signals.push(options.signal);
      return new Promise<Response>(() => undefined);
    });

    await expect(graphRequest(authConfig, '/me/drive', {}, shortDeadlines)).rejects.toThrow(/^Graph request timed out$/u);
    expect(signals).toHaveLength(1);
    expect(signals[0]?.aborted).toBe(true);
  });

  it.each([200, 500])('cancels a slow-drip Graph body on HTTP %i', async (status) => {
    let cancelled = false;
    const body = new ReadableStream<Uint8Array>({
      start: (controller) => controller.enqueue(Buffer.from('{"id":')),
      cancel: () => { cancelled = true; },
    });
    vi.spyOn(globalThis, 'fetch').mockResolvedValue(new Response(body, {
      status,
      headers: { 'Content-Type': 'application/json' },
    }));

    await expect(graphRequest(authConfig, '/me/drive', {}, shortDeadlines)).rejects.toThrow(/^Graph response body timed out$/u);
    expect(cancelled).toBe(true);
  });
});
