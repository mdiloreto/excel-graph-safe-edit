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

export const SUPPORTED_AUTHORITY_HOSTS = [
  'login.microsoftonline.com',
] as const;

export function normalizeAuthority(authority: string): string {
  const url = new URL(authority);
  if (url.protocol !== 'https:') throw new Error('Microsoft authority must use HTTPS');
  if (!SUPPORTED_AUTHORITY_HOSTS.some((host) => host === url.hostname) || url.port) {
    throw new Error(`Microsoft authority host must be one of: ${SUPPORTED_AUTHORITY_HOSTS.join(', ')}`);
  }
  if (url.username || url.password || url.search || url.hash) {
    throw new Error('Microsoft authority must not include credentials, a query, or a fragment');
  }
  const pathname = url.pathname.replace(/\/+$/, '');
  return `${url.origin}${pathname}`;
}

export function normalizeScopes(scopes: readonly string[]): string[] {
  return [...new Set(scopes.flatMap((scope) => scope.split(/[\t\n\v\f\r ]+/).filter(Boolean)))].sort();
}

export function authConfigKey(config: Pick<AuthConfig, 'clientId' | 'authority' | 'scopes'>): string {
  return JSON.stringify([config.clientId, normalizeAuthority(config.authority), normalizeScopes(config.scopes)]);
}

export function authConfigsMatch(
  left: Pick<AuthConfig, 'clientId' | 'authority' | 'scopes'>,
  right: Pick<AuthConfig, 'clientId' | 'authority' | 'scopes'>,
): boolean {
  return authConfigKey(left) === authConfigKey(right);
}

export function buildAuthConfig(options: {
  clientId?: string;
  authority?: string;
  scopes?: string[];
  noPersist?: boolean;
  port?: number;
}): AuthConfig {
  const authority = normalizeAuthority(options.authority ?? process.env.EXCEL_GRAPH_AUTHORITY ?? process.env.MICROSOFT_AUTHORITY ?? DEFAULT_AUTHORITY);
  const normalizedScopes = normalizeScopes([...DEFAULT_SCOPES, ...(options.scopes ?? [])]);
  const scopes = options.noPersist ? normalizedScopes.filter((scope) => scope !== 'offline_access') : normalizedScopes;
  const clientId = (options.clientId ?? process.env.EXCEL_GRAPH_CLIENT_ID ?? process.env.MICROSOFT_CLIENT_ID)?.trim();
  if (!clientId) throw new Error('Missing Microsoft public client id. Pass --client-id or set EXCEL_GRAPH_CLIENT_ID.');
  const port = options.port ?? 0;
  if (!Number.isInteger(port) || port < 0 || port > 65_535) {
    throw new Error('OAuth callback port must be an integer from 0 through 65535');
  }
  return {
    clientId,
    authority,
    scopes,
    persist: !options.noPersist,
    port,
  };
}
