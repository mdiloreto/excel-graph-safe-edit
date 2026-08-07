import { createWriteStream } from 'node:fs';
import { chmod, mkdir, stat } from 'node:fs/promises';
import { basename, join } from 'node:path';
import { pipeline } from 'node:stream/promises';
import { getAccessToken } from './auth.js';
import { AuthConfig, BACKUP_DIR, GRAPH_BASE } from './config.js';
import { graphRequest } from './graph.js';

interface DriveItemMetadata {
  id: string;
  name?: string;
  webUrl?: string;
}

export interface LocalBackup {
  path: string;
  bytes: number;
  item: {
    id: string;
    name?: string;
    webUrl?: string;
  };
}

export function backupFileName(name: string | undefined, timestamp = new Date()): string {
  const safeName = basename(name ?? 'workbook.xlsx').replace(/[^A-Za-z0-9._-]/g, '_');
  return `${safeName}.${timestamp.toISOString().replace(/[:.]/g, '-')}.backup.xlsx`;
}

export async function downloadBackup(config: AuthConfig, itemId: string, dir = BACKUP_DIR): Promise<LocalBackup> {
  const token = await getAccessToken(config);
  await mkdir(dir, { recursive: true, mode: 0o700 });
  await chmod(dir, 0o700).catch(() => undefined);
  const metadata = await graphRequest<DriveItemMetadata>(config, `/me/drive/items/${encodeURIComponent(itemId)}`);
  const path = join(dir, backupFileName(metadata.name));
  const response = await fetch(`${GRAPH_BASE}/me/drive/items/${encodeURIComponent(itemId)}/content`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!response.ok || !response.body) throw new Error(`Backup download failed: ${response.status}`);
  await pipeline(response.body, createWriteStream(path, { mode: 0o600 }));
  await chmod(path, 0o600).catch(() => undefined);
  const info = await stat(path);
  return { path, bytes: info.size, item: { id: metadata.id, name: metadata.name, webUrl: metadata.webUrl } };
}
