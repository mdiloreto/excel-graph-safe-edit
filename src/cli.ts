#!/usr/bin/env node
import { realpathSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { parseArgs, type CliArgs } from './args.js';
import { getAccessToken } from './auth.js';
import { downloadBackup, type LocalBackup } from './backup.js';
import { buildAuthConfig, CACHE_PATH } from './config.js';
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

function usage(): void {
  console.log(`Usage: excel-graph-safe-edit <command> [options]

Commands:
  login [--no-persist]                       Open browser login and optionally cache the token
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

Options:
  --client-id <id>       Defaults to EXCEL_GRAPH_CLIENT_ID, then MICROSOFT_CLIENT_ID
  --authority <url>      Defaults to EXCEL_GRAPH_AUTHORITY or the consumers authority
  --scope <scope>        Repeatable; adds scopes while retaining all required defaults
  --port <0-65535>       Localhost callback port; 0 (the default) selects a free port
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

const GLOBAL_OPTIONS = new Set<keyof CliArgs>(['client_id', 'authority', 'scope', 'port', 'no_persist', 'json', 'help']);
const COMMAND_OPTIONS: Record<string, ReadonlySet<keyof CliArgs>> = {
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

export async function main(argv = process.argv.slice(2)): Promise<void> {
  const args = parseArgs(argv);
  const command = args._[0];
  if (!command || args.help || command === 'help') {
    usage();
    return;
  }
  const knownCommands = new Set(['login', 'logout', 'whoami', 'search', 'metadata', 'worksheets', 'tables', 'range', 'backup', 'patch-range']);
  if (!knownCommands.has(command)) throw new Error(`Unknown command: ${command}`);
  assertCommandShape(command, args);
  const callbackPort = parsePort(args.port);

  if (command === 'logout') {
    assertNoResourceSelector(command, args);
    if (args.no_persist) throw new Error('logout cannot be combined with --no-persist');
    await clearTokenCache();
    jsonOut({ loggedOut: true, cachePath: CACHE_PATH }, args.json);
    return;
  }

  const config = buildAuthConfig({
    clientId: args.client_id,
    authority: args.authority,
    scopes: args.scope,
    noPersist: args.no_persist,
    port: callbackPort,
  });

  if (command === 'login') {
    assertNoResourceSelector(command, args);
    await getAccessToken(config);
    jsonOut({ authenticated: true, persisted: config.persist, cachePath: config.persist ? CACHE_PATH : null }, args.json);
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
