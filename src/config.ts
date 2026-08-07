import { homedir } from 'node:os';
import { join } from 'node:path';

export const DEFAULT_AUTHORITY = 'https://login.microsoftonline.com/consumers';
export const DEFAULT_SCOPES = [
  'openid',
  'profile',
  'offline_access',
  'https://graph.microsoft.com/User.Read',
  'https://graph.microsoft.com/Files.ReadWrite',
] as const;

export const STATE_DIR = join(homedir(), '.local/state/opencode-excel-graph');
export const CACHE_PATH = join(STATE_DIR, 'token-cache.json');
export const BACKUP_DIR = join(STATE_DIR, 'backups');
export const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';

export interface AuthConfig {
  clientId: string;
  authority: string;
  scopes: string[];
  persist: boolean;
  port: number;
}

export function buildAuthConfig(options: {
  clientId?: string;
  authority?: string;
  scopes?: string[];
  noPersist?: boolean;
  port?: number;
}): AuthConfig {
  const scopes = options.scopes?.length ? options.scopes : [...DEFAULT_SCOPES];
  const clientId = options.clientId ?? process.env.EXCEL_GRAPH_CLIENT_ID ?? process.env.MICROSOFT_CLIENT_ID;
  if (!clientId) throw new Error('Missing Microsoft public client id. Pass --client-id or set EXCEL_GRAPH_CLIENT_ID.');
  return {
    clientId,
    authority: options.authority ?? process.env.EXCEL_GRAPH_AUTHORITY ?? process.env.MICROSOFT_AUTHORITY ?? DEFAULT_AUTHORITY,
    scopes: options.noPersist ? scopes.filter((scope) => scope !== 'offline_access') : scopes,
    persist: !options.noPersist,
    port: options.port ?? 0,
  };
}
