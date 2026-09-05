# Excel Graph Safe Edit

TypeScript CLI and ESM package for backup-first Excel edits and create-or-verify-identical file uploads through Microsoft Graph.

## Install and build

Node.js 22 or newer is required.

```bash
npm ci
npm run check
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

Set `EXCEL_GRAPH_NONINTERACTIVE=1` for unattended `upload-docx` and `upload-file` automation. Those commands may reuse a valid in-memory or persistent token and may refresh a matching cached token, but they never launch a browser or print an authorization URL; missing, invalid, or unrefreshable cache state fails with a generic authentication error. The explicit `login` command remains interactive even when the flag is set.

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

## Safe generic file upload

`upload-file` securely reads a bounded local file and uses the same no-overwrite upload session, conflict handling, serialization, and byte-for-byte redownload verification as DOCX uploads. The command is generic and does not impose an extension, with VTT as the intended companion artifact:

```bash
chmod 600 transcript.vtt
node dist/src/cli.js upload-file \
  --input-file ./transcript.vtt \
  --path 'Class notes/transcript.vtt'

# Target another known drive
node dist/src/cli.js upload-file \
  --input-file ./transcript.vtt \
  --path 'Class notes/transcript.vtt' \
  --drive-id '<drive-id>'
```

Input must be a non-symlink regular file reached without symlink ancestors, owned by the current user, with no group or other permissions, and no larger than `MAX_SAFE_UPLOAD_BYTES`. Programmatic callers can compose `readSecureLocalFile` with `safeCreateDriveFile`, or call `createAndUploadLocalFile` directly.

## Safe DOCX creation

`upload-docx` accepts document content only from a private regular local JSON file (not from JSON on the command line), creates a DOCX locally, and uploads it without replacing an existing drive item:

```bash
chmod 600 transcript.json
node dist/src/cli.js upload-docx \
  --input-json ./transcript.json \
  --path 'Class notes/transcript.docx'

# Target another known drive
node dist/src/cli.js upload-docx \
  --input-json ./transcript.json \
  --path 'Class notes/transcript.docx' \
  --drive-id '<drive-id>'
```

The generic, exact-key input model is bounded by exported limits and has this shape:

```json
{
  "schemaVersion": 1,
  "title": "Class transcript",
  "metadata": [{ "label": "Date", "value": "2026-09-04" }],
  "sections": [
    {
      "heading": "Overview",
      "paragraphs": ["Opening discussion."],
      "bullets": ["First point", "Second point"]
    }
  ]
}
```

Each section must contain non-empty `paragraphs` and/or `bullets`. The command validates the DOCX OOXML package and local SHA-256, ensures parent folders segment-by-segment, creates an upload session with conflict behavior `fail`, uploads sequential 320-KiB-aligned chunks, then rereads the returned item by ID and verifies a bounded redownload byte-for-byte at the originally requested Graph path. If that exact path already contains a regular file, the operation succeeds only when its exact name, size, downloaded SHA-256, and bytes match the local input; otherwise it reports a conflict. This create-or-verify-identical behavior remains no-overwrite and never updates, deletes, or versions the existing item. Output remains limited to allowlisted item metadata, SHA-256, and `verified: true`.

DOCX generation is deterministic for the same validated model: core-property timestamps, ZIP entry order, ZIP timestamps, compression, and permissions are canonicalized before final OOXML validation. This lets a retry reproduce the exact bytes needed by create-or-verify-identical reconciliation.

`#` and `%` are encoded as Graph path data rather than rejected unconditionally. Microsoft filename restrictions can vary by tenant, so callers should confirm portability requirements for every target tenant.

For a two-hour timestamped transcript, aggregate adjacent VTT cues into logical timestamped paragraphs rather than creating one paragraph per cue. Keep each Word section at no more than 100 combined paragraphs and bullets; for one transcript section, 100 roughly 72-second intervals cover two hours while retaining timestamps. The existing per-block and total-character limits still apply.

Secure local input checks require POSIX ownership, mode, and no-follow file semantics, so both CLI upload commands fail closed on Windows. Programmatic callers can still validate an in-memory model with `parseWordDocumentModel` or upload already-trusted in-memory bytes with `safeCreateDriveFile`.

The package exposes declarations and an ESM entry point without deep imports. A CommonJS consumer can use dynamic import:

```js
const {
  buildAuthConfig,
  createAndUploadLocalFile,
  createAndUploadWordDocument,
  createWordDocument,
  parseWordDocumentModel,
  readSecureLocalFile,
} = await import('excel-graph-safe-edit');
```

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
- Backup response bodies are capped at 100 MiB before OOXML validation.
- Graph metadata requests have 30-second request/body deadlines, upload and verification transport has 120-second deadlines, and backup downloads have a 60-second request deadline plus a five-minute body deadline. Timed-out requests are aborted and response readers are cancelled with bounded cleanup.
- DOCX validation uses the same bounded classic-ZIP structural checks and requires `[Content_Types].xml`, `_rels/.rels`, `docProps/core.xml`, and `word/document.xml`; it has the same ZIP64, CRC, relationship, and malware-scanning limitations.
- Backups are local files and are not automatically restored.
- New backup directories are created with mode `0700`; existing custom directories with group/other permissions are rejected without changing their mode.
- Verification proves requested non-null cells matched the immediate Graph read. It cannot prevent later edits by another user or process.
- Patch serialization applies within one CLI process only; independent processes and external editors are not locked.
- Create-or-verify-identical target writes are serialized only within one process. Conflict checks and Graph's `fail` behavior prevent intentional overwrite, but external actors can still race folder and file creation.
- Secure local reads reject observed symlink ancestors, open the final component with `O_NOFOLLOW`, and revalidate component identity plus file inode, size, mtime, and ctime after reading. Node does not expose an `openat2`-style atomic traversal on all supported systems, so parent directories must still be trusted against TOCTOU replacement.
- Formula results may recalculate asynchronously, but formula verification compares the requested formula text.
- No credentials, access tokens, or authorization headers are intentionally logged.

Run `node dist/src/cli.js --help` for all commands and options.
