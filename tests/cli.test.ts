import { afterEach, describe, expect, it, vi } from 'vitest';
import { parseArgs } from '../src/args.js';
import { commandAllowsInteractiveAuthentication, patchRange, uploadDocx, uploadFile } from '../src/cli.js';
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

describe('noninteractive command policy', () => {
  it('disables interactive authentication only for upload automation commands', () => {
    const environment = { EXCEL_GRAPH_NONINTERACTIVE: '1' };
    expect(commandAllowsInteractiveAuthentication('upload-docx', environment)).toBe(false);
    expect(commandAllowsInteractiveAuthentication('upload-file', environment)).toBe(false);
    expect(commandAllowsInteractiveAuthentication('whoami', environment)).toBe(false);
    expect(commandAllowsInteractiveAuthentication('login', environment)).toBe(true);
    expect(commandAllowsInteractiveAuthentication('upload-docx', {})).toBe(true);
  });
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

describe('upload-docx workflow', () => {
  it('requires one secure input path and one target path', async () => {
    await expect(uploadDocx(authConfig, parseArgs(['upload-docx', '--path', 'file.docx']), {
      readModel: vi.fn(),
      upload: vi.fn(),
    })).rejects.toThrow(/exactly one --input-json/);
  });

  it('passes validated file content to the high-level upload API and returns only its safe result', async () => {
    const model = {
      schemaVersion: 1 as const,
      title: 'Transcript',
      sections: [{ heading: 'Class', paragraphs: ['Text'] }],
    };
    const result = {
      id: 'item',
      name: 'transcript.docx',
      size: 10,
      driveId: 'drive',
      sha256: 'a'.repeat(64),
      verified: true as const,
    };
    const readModel = vi.fn(async () => model);
    const upload = vi.fn(async () => result);
    const args = parseArgs([
      'upload-docx',
      '--input-json', '/private/model.json',
      '--path', 'Class/transcript.docx',
      '--drive-id', 'drive',
    ]);

    await expect(uploadDocx(authConfig, args, { readModel, upload })).resolves.toEqual(result);
    expect(readModel).toHaveBeenCalledWith('/private/model.json');
    expect(upload).toHaveBeenCalledWith(authConfig, {
      model,
      path: 'Class/transcript.docx',
      driveId: 'drive',
    });
  });
});

describe('upload-file workflow', () => {
  it('requires exactly one secure input path and target path', async () => {
    await expect(uploadFile(authConfig, parseArgs(['upload-file', '--path', 'file.vtt']), {
      upload: vi.fn(),
    })).rejects.toThrow(/exactly one --input-file/);
  });

  it('invokes the generic verified upload dependency with the selected drive', async () => {
    const result = {
      id: 'item',
      name: 'transcript.vtt',
      size: 100,
      driveId: 'drive',
      sha256: 'b'.repeat(64),
      verified: true as const,
    };
    const upload = vi.fn(async () => result);
    const args = parseArgs([
      'upload-file',
      '--input-file', '/private/transcript.vtt',
      '--path', 'Class/transcript.vtt',
      '--drive-id', 'drive',
    ]);

    await expect(uploadFile(authConfig, args, { upload })).resolves.toEqual(result);
    expect(upload).toHaveBeenCalledWith(authConfig, {
      inputFile: '/private/transcript.vtt',
      path: 'Class/transcript.vtt',
      driveId: 'drive',
    });
  });
});
