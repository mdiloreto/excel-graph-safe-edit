import { describe, expect, it } from 'vitest';
import { parseArgs } from '../src/args.js';
import { itemTarget, main } from '../src/cli.js';
import { buildAuthConfig, normalizeAuthority } from '../src/config.js';

describe('CLI argument parsing', () => {
  it('rejects unknown options and boolean values', () => {
    expect(() => parseArgs(['whoami', '--unknown'])).toThrow(/Unknown option/);
    expect(() => parseArgs(['whoami', '--json=false'])).toThrow(/does not accept a value/);
  });

  it('rejects empty option values instead of falling back to default-drive behavior', () => {
    for (const option of ['drive-id', 'item-id', 'path', 'sheet', 'address', 'client-id', 'authority']) {
      expect(() => parseArgs(['metadata', `--${option}=`])).toThrow(`Empty value for --${option}`);
    }
    expect(() => parseArgs(['metadata', '--item-id', 'item', '--drive-id='])).toThrow('Empty value for --drive-id');
  });

  it('supports help and the positional delimiter', () => {
    expect(parseArgs(['--help'])).toMatchObject({ help: true });
    expect(parseArgs(['search', '--', '--literal-query'])).toMatchObject({ _: ['search', '--literal-query'] });
    expect(parseArgs(['patch-range', '--values-json=[["a=b"]]']).values_json).toBe('[["a=b"]]');
  });

  it('accepts only a secure input path and target path for upload-docx and rejects duplicates', () => {
    expect(parseArgs([
      'upload-docx',
      '--input-json', '/private/model.json',
      '--path', 'Class notes/transcript.docx',
      '--drive-id', 'drive',
    ])).toMatchObject({
      _: ['upload-docx'],
      input_json: '/private/model.json',
      path: 'Class notes/transcript.docx',
      drive_id: 'drive',
    });
    expect(() => parseArgs(['upload-docx', '--input-json', 'one.json', '--input-json', 'two.json'])).toThrow(/only once/);
    expect(() => parseArgs(['upload-docx', '--path', 'one.docx', '--path', 'two.docx'])).toThrow(/only once/);
    for (const option of ['client-id', 'authority', 'port']) {
      expect(() => parseArgs(['upload-docx', `--${option}`, 'one', `--${option}`, 'two'])).toThrow(/only once/);
    }
  });

  it('accepts strict upload-file options and rejects duplicate inputs or targets', () => {
    expect(parseArgs([
      'upload-file',
      '--input-file', '/private/transcript.vtt',
      '--path', 'Class notes/transcript.vtt',
      '--drive-id', 'drive',
    ])).toMatchObject({
      _: ['upload-file'],
      input_file: '/private/transcript.vtt',
      path: 'Class notes/transcript.vtt',
      drive_id: 'drive',
    });
    expect(() => parseArgs(['upload-file', '--input-file', 'one.vtt', '--input-file', 'two.vtt'])).toThrow(/only once/);
    expect(() => parseArgs(['upload-file', '--path', 'one.vtt', '--path', 'two.vtt'])).toThrow(/only once/);
    expect(() => parseArgs(['upload-file', '--drive-id', 'one', '--drive-id', 'two'])).toThrow(/only once/);
    expect(() => parseArgs(['upload-file', '--client-id', 'one', '--client-id', 'two'])).toThrow(/only once/);
  });

  it('validates callback ports', () => {
    expect(() => buildAuthConfig({ clientId: 'client', port: -1 })).toThrow(/port/);
    expect(() => buildAuthConfig({ clientId: 'client', port: 65_536 })).toThrow(/port/);
    expect(buildAuthConfig({ clientId: 'client', port: 0 }).port).toBe(0);
  });

  it('builds a separate encrypted cache profile from absolute paths', () => {
    const config = buildAuthConfig({
      clientId: 'client',
      cacheKeyFile: '/private/material-archive.key',
      cachePath: '/private/material-archive.enc',
    });
    expect(config.cacheKeyFile).toBe('/private/material-archive.key');
    expect(config.cachePath).toBe('/private/material-archive.enc');
    expect(() => buildAuthConfig({ clientId: 'client', cacheKeyFile: 'relative.key' })).toThrow(/absolute/u);
    expect(() => buildAuthConfig({
      clientId: 'client',
      cacheKeyFile: '/private/key',
      noPersist: true,
    })).toThrow(/no-persist/u);
  });

  it('splits and adds grouped scope values to required defaults before no-persist filtering', () => {
    const parsed = parseArgs(['login', '--scope', 'openid offline_access scope']);
    const persistentScopes = buildAuthConfig({ clientId: 'client', scopes: parsed.scope }).scopes;
    expect(persistentScopes).toEqual(expect.arrayContaining([
      'openid',
      'profile',
      'offline_access',
      'scope',
      'https://graph.microsoft.com/User.Read',
      'https://graph.microsoft.com/Files.ReadWrite',
    ]));
    const oneShotScopes = buildAuthConfig({ clientId: 'client', scopes: parsed.scope, noPersist: true }).scopes;
    expect(oneShotScopes).not.toContain('offline_access');
    expect(oneShotScopes).toContain('scope');
  });

  it('keeps every required default when adding a repeatable custom scope', () => {
    const parsed = parseArgs(['login', '--scope', 'Sites.ReadWrite.All']);
    const persistentScopes = buildAuthConfig({ clientId: 'client', scopes: parsed.scope }).scopes;
    expect(persistentScopes).toEqual(expect.arrayContaining([
      'openid',
      'profile',
      'offline_access',
      'Sites.ReadWrite.All',
      'https://graph.microsoft.com/User.Read',
      'https://graph.microsoft.com/Files.ReadWrite',
    ]));
    expect(persistentScopes).toHaveLength(6);
  });

  it('allows only the global Microsoft identity authority host', () => {
    expect(normalizeAuthority('https://login.microsoftonline.com/tenant/')).toBe('https://login.microsoftonline.com/tenant');
    expect(() => normalizeAuthority('https://example.com/tenant')).toThrow(/host must be one of/);
    expect(() => normalizeAuthority('http://login.microsoftonline.com/tenant')).toThrow(/HTTPS/);
    expect(() => normalizeAuthority('https://login.microsoftonline.com:444/tenant')).toThrow(/host must be one of/);
  });

  it.each([
    'https://login.microsoftonline.us/tenant',
    'https://login.partner.microsoftonline.cn/tenant',
    'https://login.chinacloudapi.cn/tenant',
    'https://login.microsoftonline.de/tenant',
  ])('fails early for unsupported sovereign authority %s', (authority) => {
    expect(() => buildAuthConfig({ clientId: 'client', authority })).toThrow(/host must be one of/);
  });

  it('rejects conflicting item, path, and drive selectors', () => {
    expect(() => itemTarget(parseArgs(['metadata', '--item-id', 'item', '--path', 'book.xlsx']), true)).toThrow(/exactly one/);
    expect(() => itemTarget(parseArgs(['metadata', '--path', 'book.xlsx', '--drive-id', 'drive']), true)).toThrow(/requires --item-id/);
    expect(() => itemTarget(parseArgs(['range', '--path', 'book.xlsx']))).toThrow(/only supported by metadata/);
    expect(itemTarget(parseArgs(['metadata', '--item-id', 'item', '--drive-id', 'drive']), true)).toMatchObject({
      kind: 'item',
      graphPath: '/drives/drive/items/item',
    });
  });

  it('rejects options and positional arguments that do not apply to a command', async () => {
    await expect(main(['metadata', '--client-id', 'client', '--item-id', 'item', '--sheet', 'ignored'])).rejects.toThrow(/not valid for metadata/);
    await expect(main(['whoami', 'extra', '--client-id', 'client'])).rejects.toThrow(/does not accept positional/);
    await expect(main(['upload-docx', '--values-json', '{}', '--path', 'file.docx'])).rejects.toThrow(/not valid for upload-docx/);
    await expect(main(['upload-docx', '{"schemaVersion":1}', '--path', 'file.docx'])).rejects.toThrow(/does not accept positional/);
    await expect(main(['upload-file', '--input-json', 'model.json', '--path', 'file.vtt'])).rejects.toThrow(/not valid for upload-file/);
    await expect(main(['upload-docx', '--input-file', 'file.vtt', '--path', 'file.docx'])).rejects.toThrow(/not valid for upload-docx/);
    await expect(main(['upload-file', '--path', 'file.vtt'])).rejects.toThrow(/exactly one --input-file/);
  });
});
