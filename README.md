# Excel Graph Safe Edit

Safe, agent-friendly Microsoft Graph Excel workbook editing for OneDrive and SharePoint workbooks.

This project focuses on backup-first, bounded, verified edits rather than raw Graph passthrough.

## MVP

- Browser login with OAuth auth-code + PKCE.
- Optional one-shot mode with `--no-persist`.
- OneDrive workbook discovery and metadata reads.
- Worksheet, table, and range inspection.
- Local `.xlsx` backup before writes.
- Bounded range patching with post-write verification.

## Usage

```bash
npm install
npm run build
export EXCEL_GRAPH_CLIENT_ID="<public-client-app-id>"
node dist/src/cli.js login
node dist/src/cli.js search "budget.xlsx"
```

The CLI requires a Microsoft public client application id. Pass it with `--client-id` or set `EXCEL_GRAPH_CLIENT_ID`. The default authority is `https://login.microsoftonline.com/consumers` for personal Microsoft accounts; override it with `--authority` or `EXCEL_GRAPH_AUTHORITY` for another tenant.

## Safety

- Writes create a local backup first.
- Unbounded ranges are rejected for writes.
- Tokens and Authorization headers are never logged.
- Persistent token cache lives under `~/.local/state/opencode-excel-graph/` with strict file permissions.
