export function encodeOneDrivePath(path: string): string {
  return path.split('/').filter(Boolean).map(encodeURIComponent).join('/');
}

export function odataString(value: string): string {
  return value.replaceAll("'", "''");
}

export function itemPathById(itemId: string): string {
  return `/me/drive/items/${encodeURIComponent(itemId)}`;
}

export function itemPathByDrivePath(path: string): string {
  return `/me/drive/root:/${encodeOneDrivePath(path)}:`;
}

export function workbookRangePath(itemId: string, sheet: string, address: string): string {
  return `${itemPathById(itemId)}/workbook/worksheets/${encodeURIComponent(sheet)}/range(address='${odataString(address)}')`;
}

export function worksheetTablesPath(itemId: string, sheet: string): string {
  return `${itemPathById(itemId)}/workbook/worksheets/${encodeURIComponent(sheet)}/tables`;
}

export function worksheetsPath(itemId: string): string {
  return `${itemPathById(itemId)}/workbook/worksheets`;
}

export function searchPath(query: string): string {
  return `/me/drive/root/search(q='${odataString(query)}')?$select=id,name,size,webUrl,parentReference,file,folder`;
}

export function assertBoundedWriteRange(address: string): void {
  const bare = address.includes('!') ? address.slice(address.lastIndexOf('!') + 1) : address;
  const normalized = bare.replaceAll('$', '').toUpperCase();
  if (/^[A-Z]+:[A-Z]+$/.test(normalized) || /^\d+:\d+$/.test(normalized)) {
    throw new Error(`Refusing to write unbounded range: ${address}`);
  }
}

export function columnNumber(column: string): number {
  return column.toUpperCase().split('').reduce((total, char) => total * 26 + char.charCodeAt(0) - 64, 0);
}

export function rangeDimensions(address: string): { rows: number; columns: number } {
  const bare = address.includes('!') ? address.slice(address.lastIndexOf('!') + 1) : address;
  const normalized = bare.replaceAll('$', '').toUpperCase();
  const [start, end = start] = normalized.split(':');
  const startMatch = /^([A-Z]+)(\d+)$/.exec(start ?? '');
  const endMatch = /^([A-Z]+)(\d+)$/.exec(end ?? '');
  if (!startMatch || !endMatch) throw new Error(`Cannot determine dimensions for range: ${address}`);
  return {
    rows: Math.abs(Number(endMatch[2]) - Number(startMatch[2])) + 1,
    columns: Math.abs(columnNumber(endMatch[1] ?? '') - columnNumber(startMatch[1] ?? '')) + 1,
  };
}

export function assertMatrixMatchesRange(address: string, matrix: unknown): void {
  if (!Array.isArray(matrix) || !matrix.every((row) => Array.isArray(row))) {
    throw new Error('Range payload must be a 2D array');
  }
  const dimensions = rangeDimensions(address);
  const rows = matrix.length;
  const columns = rows === 0 ? 0 : (matrix[0] as unknown[]).length;
  if (rows !== dimensions.rows || !matrix.every((row) => (row as unknown[]).length === dimensions.columns) || columns !== dimensions.columns) {
    throw new Error(`Range payload dimensions ${rows}x${columns} do not match ${address} (${dimensions.rows}x${dimensions.columns})`);
  }
}
