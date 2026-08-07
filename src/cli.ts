#!/usr/bin/env node
import { clearTokenCache } from './token-cache.js';
import { buildAuthConfig, CACHE_PATH } from './config.js';
import { getAccessToken } from './auth.js';
import { graphRequest } from './graph.js';
import { downloadBackup } from './backup.js';
import {
  assertBoundedWriteRange,
  assertMatrixMatchesRange,
  itemPathByDrivePath,
  itemPathById,
  searchPath,
  workbookRangePath,
  worksheetsPath,
  worksheetTablesPath,
} from './excel.js';
import { parseArgs } from './args.js';

function usage(): void {
  console.log(`Usage: excel-graph-safe-edit <command> [options]

Commands:
  login [--no-persist]              Open browser login and cache delegated Graph token by default
  logout                            Remove cached credentials
  whoami [--no-persist]             Show signed-in drive/account metadata
  search <query>                    Search OneDrive items
  metadata --item-id <id>           Show driveItem metadata
  metadata --path <path>            Show driveItem metadata by OneDrive path
  worksheets --item-id <id>         List workbook worksheets
  tables --item-id <id> --sheet <s> List worksheet tables
  range --item-id <id> --sheet <s> --address <A1:B2>
                                    Read workbook range values/formulas
  backup --item-id <id> [--dir <d>] Download a timestamped local .xlsx backup
  patch-range --item-id <id> --sheet <s> --address <A1:B2> --values-json <json>
  patch-range --item-id <id> --sheet <s> --address <A1:B2> --formulas-json <json>
                                    Back up, patch a bounded range, then re-read it

Auth/config options:
  --client-id <id>                  Defaults to EXCEL_GRAPH_CLIENT_ID, then MICROSOFT_CLIENT_ID
  --authority <url>                 Defaults to EXCEL_GRAPH_AUTHORITY or consumers authority
  --scope <scope>                   Repeatable. Defaults to Files.ReadWrite delegated scope
  --port <port>                     Localhost callback port; default random
  --no-persist                      Do not read/write token cache; browser login for this invocation
  --json                            Emit compact JSON where supported`);
}

function jsonOut(value: unknown, compact = false): void {
  console.log(JSON.stringify(value, null, compact ? 0 : 2));
}

function itemSegment(args: { item_id?: string; path?: string }): string {
  if (args.item_id) return itemPathById(args.item_id);
  if (args.path) return itemPathByDrivePath(args.path);
  throw new Error('Expected --item-id <id> or --path <path>');
}

async function main(): Promise<void> {
  const args = parseArgs(process.argv.slice(2));
  const command = args._[0];
  if (!command || ['help', '-h', '--help'].includes(command)) {
    usage();
    return;
  }
  if (command === 'logout') {
    await clearTokenCache();
    jsonOut({ loggedOut: true, cachePath: CACHE_PATH }, args.json);
    return;
  }

  const config = buildAuthConfig({
    clientId: args.client_id,
    authority: args.authority,
    scopes: args.scope,
    noPersist: args.no_persist,
    port: args.port ? Number(args.port) : undefined,
  });

  if (command === 'login') {
    await getAccessToken(config);
    jsonOut({ authenticated: true, persisted: config.persist, cachePath: config.persist ? CACHE_PATH : null }, args.json);
    return;
  }
  if (command === 'whoami') {
    const [drive, me] = await Promise.all([graphRequest(config, '/me/drive'), graphRequest(config, '/me')]);
    jsonOut({ me, drive }, args.json);
    return;
  }
  if (command === 'search') {
    const query = args._.slice(1).join(' ');
    if (!query) throw new Error('Expected search query');
    const payload = await graphRequest<{ value?: unknown[] }>(config, searchPath(query));
    jsonOut(payload.value ?? payload, args.json);
    return;
  }
  if (command === 'metadata') {
    jsonOut(await graphRequest(config, itemSegment(args)), args.json);
    return;
  }
  if (command === 'worksheets') {
    if (!args.item_id) throw new Error('Expected --item-id <id>');
    jsonOut(await graphRequest(config, worksheetsPath(args.item_id)), args.json);
    return;
  }
  if (command === 'tables') {
    if (!args.item_id || !args.sheet) throw new Error('Expected --item-id <id> --sheet <sheet>');
    jsonOut(await graphRequest(config, worksheetTablesPath(args.item_id, args.sheet)), args.json);
    return;
  }
  if (command === 'range') {
    if (!args.item_id || !args.sheet || !args.address) throw new Error('Expected --item-id <id> --sheet <sheet> --address <range>');
    jsonOut(await graphRequest(config, workbookRangePath(args.item_id, args.sheet, args.address)), args.json);
    return;
  }
  if (command === 'backup') {
    if (!args.item_id) throw new Error('Expected --item-id <id>');
    jsonOut(await downloadBackup(config, args.item_id, args.dir), args.json);
    return;
  }
  if (command === 'patch-range') {
    if (!args.item_id || !args.sheet || !args.address) throw new Error('Expected --item-id <id> --sheet <sheet> --address <range>');
    assertBoundedWriteRange(args.address);
    const body = args.values_json
      ? { values: JSON.parse(args.values_json) as unknown }
      : args.formulas_json
        ? { formulas: JSON.parse(args.formulas_json) as unknown }
        : null;
    if (!body) throw new Error('Expected --values-json <json> or --formulas-json <json>');
    assertMatrixMatchesRange(args.address, 'values' in body ? body.values : body.formulas);
    const backup = await downloadBackup(config, args.item_id, args.backup_dir);
    await graphRequest(config, workbookRangePath(args.item_id, args.sheet, args.address), {
      method: 'PATCH',
      body: JSON.stringify(body),
    });
    const verification = await graphRequest(config, workbookRangePath(args.item_id, args.sheet, args.address));
    jsonOut({ backup, verification }, args.json);
    return;
  }
  throw new Error(`Unknown command: ${command}`);
}

main().catch((error: unknown) => {
  console.error(error instanceof Error ? error.message : String(error));
  process.exitCode = 1;
});
