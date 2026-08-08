import { afterEach, describe, expect, it, vi } from 'vitest';
import { parseArgs } from '../src/args.js';
import { patchRange } from '../src/cli.js';
import { buildAuthConfig } from '../src/config.js';
import { type LocalBackup } from '../src/backup.js';

const authConfig = buildAuthConfig({ clientId: 'client-id', noPersist: true });
const backup: LocalBackup = {
  path: '/safe/backup.xlsx',
  bytes: 123,
  sha256: 'abc123',
  item: { id: 'item', driveId: 'drive' },
};

function mutationArgs() {
  return parseArgs([
    'patch-range',
    '--item-id', 'item',
    '--drive-id', 'drive',
    '--sheet', 'Sheet 1',
    '--address', 'A1:B1',
    '--values-json', '[[null,"changed"]]',
  ]);
}

afterEach(() => {
  vi.restoreAllMocks();
});

describe('patch workflow', () => {
  it('reconciles an ambiguous transport failure without retrying PATCH', async () => {
    vi.spyOn(console, 'error').mockImplementation(() => undefined);
    const backupDownload = vi.fn(async () => backup);
    const request = vi.fn(async (_config, path: string, options?: RequestInit) => {
      expect(path).toContain('/drives/drive/items/item/');
      if (options?.method === 'PATCH') throw new TypeError('socket closed');
      return { address: 'Sheet 1!A1:B1', values: [['unchanged', 'changed']] };
    });

    await expect(patchRange(authConfig, mutationArgs(), { backup: backupDownload, request })).resolves.toMatchObject({
      backup,
      patch: { transport: 'reconciled', verificationMatched: true },
    });
    expect(request.mock.calls.filter((call) => call[2]?.method === 'PATCH')).toHaveLength(1);
    expect(request).toHaveBeenCalledTimes(2);
  });

  it('fails with backup metadata when post-PATCH verification mismatches', async () => {
    vi.spyOn(console, 'error').mockImplementation(() => undefined);
    const request = vi.fn(async (_config, _path: string, options?: RequestInit) => {
      if (options?.method === 'PATCH') return {};
      return { address: 'Sheet 1!A1:B1', values: [['unchanged', 'wrong']] };
    });

    await expect(patchRange(authConfig, mutationArgs(), {
      backup: async () => backup,
      request,
    })).rejects.toThrow('backup=/safe/backup.xlsx; verificationMatched=false');
    expect(request).toHaveBeenCalledTimes(2);
  });

  it('does not issue a mutation when backup creation fails', async () => {
    const request = vi.fn(async () => ({}));
    await expect(patchRange(authConfig, mutationArgs(), {
      backup: async () => { throw new Error('invalid backup content'); },
      request,
    })).rejects.toThrow('invalid backup content');
    expect(request).not.toHaveBeenCalled();
  });
});
