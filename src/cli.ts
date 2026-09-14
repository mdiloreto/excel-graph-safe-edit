#!/usr/bin/env node
import { realpathSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { parseArgs, type CliArgs } from './args.js';
import { getAccessToken } from './auth.js';
import { downloadBackup, type LocalBackup } from './backup.js';
import { buildAuthConfig, CACHE_PATH, ENCRYPTED_CACHE_PATH } from './config.js';
import { clearEncryptedMsalCache, createEncryptedCacheKey } from './encrypted-msal-cache.js';
import {
  assertBoundedWriteRange,
  itemPathByDrivePath,
  itemPathById,
  parseRangeMutation,
  searchPath,
  serializeWorkbookMutation,
  verifyRangeMutation,
  workbookRangePath,
  worksheetsPath,
  worksheetTablesPath,
  type MutationVerification,
  type RangeMutation,
} from './excel.js';
import { graphRequest, isAmbiguousWriteError } from './graph.js';
import { clearTokenCache } from './token-cache.js';
import {
  createAndUploadWordDocument,
  readSecureWordDocumentModel,
  type WordDocumentModel,
} from './word.js';
import { type VerifiedDriveItem } from './drive.js';
import { createAndUploadLocalFile } from './local-file.js';

function usage(): void {
  console.log(`Usage: excel-graph-safe-edit <command> [options]

Commands:
  init-cache-key --cache-key-file <path>     Create a private encryption key for persistent MSAL auth
  login [--no-persist]                       Authenticate and optionally cache the token
  logout                                    Remove cached credentials
  whoami [--no-persist]                      Show signed-in account and default drive metadata
  search <query>                             Search the current user's OneDrive (not other drives)
  metadata --item-id <id> [--drive-id <id>] Show driveItem metadata by item id
  metadata --path <path>                     Show metadata by current-user OneDrive path
  worksheets --item-id <id> [--drive-id <id>]
  tables --item-id <id> [--drive-id <id>] --sheet <sheet>
  range --item-id <id> [--drive-id <id>] --sheet <sheet> --address <A1:B2>
  backup --item-id <id> [--drive-id <id>] [--dir <directory>]
  patch-range --item-id <id> [--drive-id <id>] --sheet <sheet> --address <A1:B2>
               (--values-json <json> | --formulas-json <json>) [--backup-dir <directory>]
                                              Back up, patch, and verify the exact bounded range
  upload-docx --input-json <secure-local-file> --path <OneDrive/path.docx> [--drive-id <id>]
                                               Safely create, upload, redownload, and verify a DOCX
  upload-file --input-file <secure-local-file> --path <OneDrive/path.vtt> [--drive-id <id>]
                                               Safely create, upload, redownload, and verify a file

Options:
  --client-id <id>       Defaults to EXCEL_GRAPH_CLIENT_ID, then MICROSOFT_CLIENT_ID
  --authority <url>      Defaults to EXCEL_GRAPH_AUTHORITY or the consumers authority
  --scope <scope>        Repeatable; adds scopes while retaining all required defaults
  --port <0-65535>       Localhost callback port; 0 (the default) selects a free port
  --cache-key-file <path> Use encrypted MSAL cache and Device Code authentication
  --cache-path <path>     Override encrypted MSAL cache path
  --no-persist           Never read or write the token cache; required on Windows
  --json                 Emit compact JSON where supported
  --help, -h             Show this help
  --                     Treat all remaining arguments as positional values`);
}

function jsonOut(value: unknown, compact = false): void {
  console.log(JSON.stringify(value, null, compact ? 0 : 2));
}

function assertNoResourceSelector(command: string, args: CliArgs): void {
  if (args.item_id || args.path || args.drive_id) {
    throw new Error(`${command} does not accept --item-id, --path, or --drive-id`);
  }
}

const GLOBAL_OPTIONS = new Set<keyof CliArgs>([
  'client_id', 'authority', 'scope', 'port', 'no_persist', 'json', 'help', 'cache_key_file', 'cache_path',
]);
const COMMAND_OPTIONS: Record<string, ReadonlySet<keyof CliArgs>> = {
  'init-cache-key': new Set(['cache_key_file']),
  login: new Set(),
  logout: new Set(),
  whoami: new Set(),
  search: new Set(),
  metadata: new Set(['item_id', 'drive_id', 'path']),
  worksheets: new Set(['item_id', 'drive_id']),
  tables: new Set(['item_id', 'drive_id', 'sheet']),
  range: new Set(['item_id', 'drive_id', 'sheet', 'address']),
  backup: new Set(['item_id', 'drive_id', 'dir']),
  'patch-range': new Set(['item_id', 'drive_id', 'sheet', 'address', 'backup_dir', 'values_json', 'formulas_json']),
  'upload-docx': new Set(['input_json', 'path', 'drive_id']),
  'upload-file': new Set(['input_file', 'path', 'drive_id']),
};

function assertCommandShape(command: string, args: CliArgs): void {
  const allowed = COMMAND_OPTIONS[command];
  if (!allowed) throw new Error(`Unknown command: ${command}`);
  for (const key of Object.keys(args) as Array<keyof CliArgs>) {
    if (key !== '_' && args[key] !== undefined && !GLOBAL_OPTIONS.has(key) && !allowed.has(key)) {
      throw new Error(`--${key.replaceAll('_', '-')} is not valid for ${command}`);
    }
  }
  const expectedPositionals = command === 'search' ? 2 : 1;
  if (args._.length < expectedPositionals) {
    if (command === 'search') throw new Error('Expected search query');
    throw new Error(`Expected command: ${command}`);
  }
  if (command !== 'search' && args._.length > expectedPositionals) {
    throw new Error(`${command} does not accept positional arguments`);
  }
  if (command === 'upload-docx' && (!args.input_json || !args.path)) {
    throw new Error('Expected exactly one --input-json <secure-local-file> and one --path <OneDrive/path.docx>');
  }
  if (command === 'upload-file' && (!args.input_file || !args.path)) {
    throw new Error('Expected exactly one --input-file <secure-local-file> and one --path <OneDrive/path>');
  }
}

type ItemTarget =
  | { kind: 'item'; itemId: string; driveId?: string; graphPath: string }
  | { kind: 'path'; path: string; graphPath: string };

export function itemTarget(args: CliArgs, allowPath = false): ItemTarget {
  if (args.item_id && args.path) throw new Error('Use exactly one of --item-id or --path');
  if (args.drive_id && !args.item_id) throw new Error('--drive-id requires --item-id');
  if (args.path) {
    if (!allowPath) throw new Error('--path is only supported by metadata');
    return { kind: 'path', path: args.path, graphPath: itemPathByDrivePath(args.path) };
  }
  if (!args.item_id) throw new Error('Expected --item-id <id>');
  return {
    kind: 'item',
    itemId: args.item_id,
    driveId: args.drive_id,
    graphPath: itemPathById(args.item_id, args.drive_id),
  };
}

function itemIdTarget(args: CliArgs): Extract<ItemTarget, { kind: 'item' }> {
  const target = itemTarget(args);
  if (target.kind !== 'item') throw new Error('Expected --item-id <id>');
  return target;
}

function parsePort(value: string | undefined): number | undefined {
  if (value === undefined) return undefined;
  if (!/^\d+$/.test(value)) throw new Error('OAuth callback port must be an integer from 0 through 65535');
  return Number(value);
}

export function commandAllowsInteractiveAuthentication(
  command: string,
  environment: NodeJS.ProcessEnv = process.env,
): boolean {
  return environment.EXCEL_GRAPH_NONINTERACTIVE !== '1' || command === 'login';
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

async function readAndVerify(
  config: ReturnType<typeof buildAuthConfig>,
  targetPath: string,
  address: string,
  mutation: RangeMutation,
  request: PatchRangeDependencies['request'],
): Promise<{ payload: unknown; result: MutationVerification }> {
  const payload = await request(config, targetPath);
  return { payload, result: verifyRangeMutation(address, mutation, payload) };
}

interface PatchRangeDependencies {
  request: (config: ReturnType<typeof buildAuthConfig>, path: string, options?: RequestInit) => Promise<unknown>;
  backup: typeof downloadBackup;
}

export async function patchRange(
  config: ReturnType<typeof buildAuthConfig>,
  args: CliArgs,
  dependencies: PatchRangeDependencies = { request: graphRequest, backup: downloadBackup },
): Promise<{ backup: LocalBackup; patch: { transport: 'confirmed' | 'reconciled'; verificationMatched: true }; verification: unknown }> {
  const target = itemIdTarget(args);
  if (!args.sheet || !args.address) {
    throw new Error('Expected --item-id <id> --sheet <sheet> --address <range>');
  }
  const { address, sheet } = args;
  assertBoundedWriteRange(address);
  const mutation = parseRangeMutation({
    address,
    valuesJson: args.values_json,
    formulasJson: args.formulas_json,
  });
  const workbookKey = `${target.driveId ?? 'me'}:${target.itemId}`;
  return serializeWorkbookMutation(workbookKey, async () => {
    const backup = await dependencies.backup(config, target.itemId, args.backup_dir, target.driveId);
    console.error(`Backup ready before PATCH: ${backup.path} (${backup.bytes} bytes, sha256 ${backup.sha256})`);
    const rangePath = workbookRangePath(target.itemId, sheet, address, target.driveId);
    try {
      await dependencies.request(config, rangePath, {
        method: 'PATCH',
        body: JSON.stringify({ [mutation.kind]: mutation.matrix }),
      });
    } catch (error) {
      if (!isAmbiguousWriteError(error)) {
        throw new Error(`PATCH rejected; backup=${backup.path}; verificationMatched=not-run; ${errorMessage(error)}`);
      }
      try {
        const verification = await readAndVerify(config, rangePath, address, mutation, dependencies.request);
        if (!verification.result.matched) {
          throw new Error(`PATCH transport was ambiguous; backup=${backup.path}; verificationMatched=false; mismatches=${verification.result.mismatches.length}`);
        }
        return {
          backup,
          patch: { transport: 'reconciled' as const, verificationMatched: true as const },
          verification: verification.payload,
        };
      } catch (verificationError) {
        const message = errorMessage(verificationError);
        if (message.includes(`backup=${backup.path}`)) throw verificationError;
        throw new Error(`PATCH transport was ambiguous; backup=${backup.path}; verificationMatched=unknown; ${message}`);
      }
    }

    try {
      const verification = await readAndVerify(config, rangePath, address, mutation, dependencies.request);
      if (!verification.result.matched) {
        throw new Error(`PATCH verification mismatch; backup=${backup.path}; verificationMatched=false; mismatches=${verification.result.mismatches.length}`);
      }
      return {
        backup,
        patch: { transport: 'confirmed' as const, verificationMatched: true as const },
        verification: verification.payload,
      };
    } catch (verificationError) {
      const message = errorMessage(verificationError);
      if (message.includes(`backup=${backup.path}`)) throw verificationError;
      throw new Error(`PATCH completed but verification failed; backup=${backup.path}; verificationMatched=unknown; ${message}`);
    }
  });
}

interface UploadDocxDependencies {
  readModel: (path: string) => Promise<WordDocumentModel>;
  upload: typeof createAndUploadWordDocument;
}

export async function uploadDocx(
  config: ReturnType<typeof buildAuthConfig>,
  args: CliArgs,
  dependencies: UploadDocxDependencies = {
    readModel: readSecureWordDocumentModel,
    upload: createAndUploadWordDocument,
  },
): Promise<VerifiedDriveItem> {
  if (!args.input_json || !args.path) {
    throw new Error('Expected exactly one --input-json <secure-local-file> and one --path <OneDrive/path.docx>');
  }
  const model = await dependencies.readModel(args.input_json);
  return dependencies.upload(config, { model, path: args.path, driveId: args.drive_id });
}

interface UploadFileDependencies {
  upload: typeof createAndUploadLocalFile;
}

export async function uploadFile(
  config: ReturnType<typeof buildAuthConfig>,
  args: CliArgs,
  dependencies: UploadFileDependencies = { upload: createAndUploadLocalFile },
): Promise<VerifiedDriveItem> {
  if (!args.input_file || !args.path) {
    throw new Error('Expected exactly one --input-file <secure-local-file> and one --path <OneDrive/path>');
  }
  return dependencies.upload(config, {
    inputFile: args.input_file,
    path: args.path,
    driveId: args.drive_id,
  });
}

export async function main(argv = process.argv.slice(2)): Promise<void> {
  const args = parseArgs(argv);
  const command = args._[0];
  if (!command || args.help || command === 'help') {
    usage();
    return;
  }
  const knownCommands = new Set(['init-cache-key', 'login', 'logout', 'whoami', 'search', 'metadata', 'worksheets', 'tables', 'range', 'backup', 'patch-range', 'upload-docx', 'upload-file']);
  if (!knownCommands.has(command)) throw new Error(`Unknown command: ${command}`);
  assertCommandShape(command, args);
  const callbackPort = parsePort(args.port);

  if (command === 'init-cache-key') {
    assertNoResourceSelector(command, args);
    if (!args.cache_key_file) throw new Error('Expected --cache-key-file <absolute-path>');
    await createEncryptedCacheKey(args.cache_key_file);
    jsonOut({ created: true, keyFile: args.cache_key_file }, args.json);
    return;
  }
  if (command === 'logout') {
    assertNoResourceSelector(command, args);
    if (args.no_persist) throw new Error('logout cannot be combined with --no-persist');
    const effectiveKeyFile = args.cache_key_file ?? process.env.EXCEL_GRAPH_CACHE_KEY_FILE;
    if (effectiveKeyFile) {
      const effectiveCachePath = args.cache_path ?? process.env.EXCEL_GRAPH_CACHE_PATH ?? ENCRYPTED_CACHE_PATH;
      if (!effectiveKeyFile.startsWith('/') || !effectiveCachePath.startsWith('/')) {
        throw new Error('Encrypted cache paths must be absolute');
      }
      await clearEncryptedMsalCache(effectiveCachePath);
      jsonOut({ loggedOut: true, cachePath: effectiveCachePath }, args.json);
    } else {
      await clearTokenCache();
      jsonOut({ loggedOut: true, cachePath: CACHE_PATH }, args.json);
    }
    return;
  }

  const config = buildAuthConfig({
    clientId: args.client_id,
    authority: args.authority,
    scopes: args.scope,
    noPersist: args.no_persist,
    port: callbackPort,
    cacheKeyFile: args.cache_key_file,
    cachePath: args.cache_path,
    allowInteractive: commandAllowsInteractiveAuthentication(command),
  });

  if (command === 'login') {
    assertNoResourceSelector(command, args);
    await getAccessToken(config);
    jsonOut({
      authenticated: true,
      persisted: config.persist,
      cachePath: config.persist ? (config.cachePath ?? CACHE_PATH) : null,
    }, args.json);
    return;
  }
  if (command === 'whoami') {
    assertNoResourceSelector(command, args);
    const [drive, me] = await Promise.all([graphRequest(config, '/me/drive'), graphRequest(config, '/me')]);
    jsonOut({ me, drive }, args.json);
    return;
  }
  if (command === 'search') {
    assertNoResourceSelector(command, args);
    const query = args._.slice(1).join(' ');
    if (!query) throw new Error('Expected search query');
    jsonOut(await graphRequest(config, searchPath(query)), args.json);
    return;
  }
  if (command === 'metadata') {
    jsonOut(await graphRequest(config, itemTarget(args, true).graphPath), args.json);
    return;
  }
  if (command === 'worksheets') {
    const target = itemIdTarget(args);
    jsonOut(await graphRequest(config, worksheetsPath(target.itemId, target.driveId)), args.json);
    return;
  }
  if (command === 'tables') {
    const target = itemIdTarget(args);
    if (!args.sheet) throw new Error('Expected --sheet <sheet>');
    jsonOut(await graphRequest(config, worksheetTablesPath(target.itemId, args.sheet, target.driveId)), args.json);
    return;
  }
  if (command === 'range') {
    const target = itemIdTarget(args);
    if (!args.sheet || !args.address) throw new Error('Expected --sheet <sheet> --address <range>');
    jsonOut(await graphRequest(config, workbookRangePath(target.itemId, args.sheet, args.address, target.driveId)), args.json);
    return;
  }
  if (command === 'backup') {
    const target = itemIdTarget(args);
    jsonOut(await downloadBackup(config, target.itemId, args.dir, target.driveId), args.json);
    return;
  }
  if (command === 'patch-range') {
    jsonOut(await patchRange(config, args), args.json);
    return;
  }
  if (command === 'upload-docx') {
    jsonOut(await uploadDocx(config, args), args.json);
    return;
  }
  if (command === 'upload-file') {
    jsonOut(await uploadFile(config, args), args.json);
    return;
  }
}

function isMainModule(): boolean {
  if (!process.argv[1]) return false;
  try {
    return realpathSync(process.argv[1]) === fileURLToPath(import.meta.url);
  } catch {
    return false;
  }
}

if (isMainModule()) {
  main().catch((error: unknown) => {
    console.error(errorMessage(error));
    process.exitCode = 1;
  });
}
