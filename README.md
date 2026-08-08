# Excel Graph Safe Edit

Dependency-free TypeScript CLI for backup-first, bounded, verified Microsoft Graph Excel edits in OneDrive and SharePoint drives.

## Install and build

Node.js 22 or newer is required.

```bash
npm ci
npm run build
node dist/src/cli.js --help
```

The package runs `npm run build` during `prepare`, so Git and package installs produce `dist/src/cli.js` without Unix-only lifecycle commands. Published package contents are limited to compiled JavaScript, declarations, the README, and the license.

## Authentication

Register a Microsoft public client application that permits localhost redirects and delegated Microsoft Graph access. Supply its client ID and, when needed, a tenant authority:

```bash
export EXCEL_GRAPH_CLIENT_ID="<public-client-app-id>"
export EXCEL_GRAPH_AUTHORITY="https://login.microsoftonline.com/<tenant-id>"
node dist/src/cli.js login
```

The default authority is `https://login.microsoftonline.com/consumers`. Only the global Microsoft identity host `login.microsoftonline.com` and global `graph.microsoft.com` API are currently supported; sovereign and national clouds fail before login. Tenant paths are supported. Required defaults always include OIDC scopes plus fully-qualified `User.Read` and `Files.ReadWrite`. Repeatable `--scope` values are additive, so `--scope Sites.ReadWrite.All` retains every required default; whitespace-separated scopes within one value are normalized individually.

Authentication uses authorization-code flow with PKCE and a loopback callback. If the browser cannot be opened, the CLI prints a one-time authorization URL to open manually. The callback rejects mismatched OAuth state and times out after five minutes.

Tokens are reused only for the same client ID, normalized authority, normalized scope set, and persistence mode. The persistent cache is stored at `~/.local/state/opencode-excel-graph/token-cache.json` with restricted permissions. Malformed cache files are quarantined, and logout removes active, quarantined, and abandoned temporary cache artifacts. Because POSIX modes do not enforce a private Windows DACL, persistent authentication fails closed on Windows; use `--no-persist` there. `--no-persist` removes `offline_access` from defaults and does not read or write the token cache.

## OneDrive and SharePoint drives

Item-ID commands use the signed-in user's default OneDrive unless `--drive-id` is supplied:

```bash
# Current user's OneDrive: /me/drive/items/{itemId}
node dist/src/cli.js metadata --item-id '<item-id>'
node dist/src/cli.js range --item-id '<item-id>' --sheet 'Sheet 1' --address 'A1:B2'

# SharePoint or another known drive: /drives/{driveId}/items/{itemId}
node dist/src/cli.js metadata --drive-id '<drive-id>' --item-id '<item-id>'
node dist/src/cli.js worksheets --drive-id '<drive-id>' --item-id '<item-id>'
node dist/src/cli.js tables --drive-id '<drive-id>' --item-id '<item-id>' --sheet 'Sheet 1'
node dist/src/cli.js range --drive-id '<drive-id>' --item-id '<item-id>' --sheet 'Sheet 1' --address 'A1:B2'
node dist/src/cli.js backup --drive-id '<drive-id>' --item-id '<item-id>'
```

`metadata --path` and `search` are intentionally limited to the current user's OneDrive. `--drive-id` requires `--item-id`; it cannot be combined with `--path`. Search preserves the full Graph response envelope, including `@odata.nextLink` when Graph returns it.

## Verified mutations

`patch-range` requires exactly one JSON matrix whose dimensions match the bounded target range:

```bash
node dist/src/cli.js patch-range \
  --item-id '<item-id>' \
  --sheet 'Sheet 1' \
  --address 'A1:B2' \
  --values-json '[["North",10],[null,20]]'

node dist/src/cli.js patch-range \
  --drive-id '<sharepoint-drive-id>' \
  --item-id '<item-id>' \
  --sheet 'Sheet 1' \
  --address 'C2:C3' \
  --formulas-json '[["=A2+B2"],["=A3+B3"]]'
```

Cells may be strings, booleans, `null`, or finite numbers. Nested values, ragged matrices, non-finite numbers, and all-null payloads are rejected. Requested `null` cells are left unchanged and ignored during verification. Ranges must stay within Excel's `XFD`/`1048576` grid, be ordered and syntactically bounded, and contain at most 10,000 cells per write.

Before PATCH, the CLI downloads the workbook to an exclusive temporary file, validates its ZIP directory and required XLSX entries, computes SHA-256 metadata, and publishes a restricted-permission backup with an atomic no-replace hard link. It reports the backup path before mutation. After PATCH, it rereads the exact range and compares every requested non-null value or formula. A mismatch exits nonzero and includes the backup path. Ambiguous PATCH transport failures are reconciled with one read; writes are never blindly retried.

## Safety limitations

- Backup validation checks classic ZIP EOCD/central/local-header structure and requires `[Content_Types].xml` and `xl/workbook.xml`. It rejects ZIP64 and does not decompress entries, validate every CRC/XML relationship, or perform malware scanning.
- Backups are local files and are not automatically restored.
- New backup directories are created with mode `0700`; existing custom directories with group/other permissions are rejected without changing their mode.
- Verification proves requested non-null cells matched the immediate Graph read. It cannot prevent later edits by another user or process.
- Patch serialization applies within one CLI process only; independent processes and external editors are not locked.
- Formula results may recalculate asynchronously, but formula verification compares the requested formula text.
- No credentials, access tokens, or authorization headers are intentionally logged.

Run `node dist/src/cli.js --help` for all commands and options.
