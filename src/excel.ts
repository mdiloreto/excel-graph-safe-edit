export type GraphCell = string | boolean | null | number;

export type RangeMutation =
  | { kind: 'values'; matrix: GraphCell[][] }
  | { kind: 'formulas'; matrix: GraphCell[][] };

export interface MutationVerification {
  matched: boolean;
  field: RangeMutation['kind'];
  address: string;
  mismatches: Array<{ row: number; column: number; expected: Exclude<GraphCell, null>; actual: unknown }>;
}

const MAX_EXCEL_COLUMNS = 16_384;
const MAX_EXCEL_ROWS = 1_048_576;
export const MAX_WRITE_CELLS = 10_000;

function encodePathSegment(value: string): string {
  return encodeURIComponent(value).replace(/[!'()*]/g, (character) => `%${character.charCodeAt(0).toString(16).toUpperCase()}`);
}

export function encodeOneDrivePath(path: string): string {
  return path.split('/').filter(Boolean).map(encodePathSegment).join('/');
}

export function odataString(value: string): string {
  return encodePathSegment(value.replaceAll("'", "''"));
}

export function itemPathById(itemId: string, driveId?: string): string {
  if (!itemId.trim()) throw new Error('Item ID must not be empty');
  if (driveId !== undefined && !driveId.trim()) throw new Error('Drive ID must not be empty');
  const item = encodePathSegment(itemId);
  return driveId !== undefined
    ? `/drives/${encodePathSegment(driveId)}/items/${item}`
    : `/me/drive/items/${item}`;
}

export function itemPathByDrivePath(path: string): string {
  return `/me/drive/root:/${encodeOneDrivePath(path)}:`;
}

export function workbookRangePath(itemId: string, sheet: string, address: string, driveId?: string): string {
  return `${itemPathById(itemId, driveId)}/workbook/worksheets/${encodePathSegment(sheet)}/range(address='${odataString(address)}')`;
}

export function worksheetTablesPath(itemId: string, sheet: string, driveId?: string): string {
  return `${itemPathById(itemId, driveId)}/workbook/worksheets/${encodePathSegment(sheet)}/tables`;
}

export function worksheetsPath(itemId: string, driveId?: string): string {
  return `${itemPathById(itemId, driveId)}/workbook/worksheets`;
}

export function searchPath(query: string): string {
  return `/me/drive/root/search(q='${odataString(query)}')?$select=id,name,size,webUrl,parentReference,file,folder`;
}

export function assertBoundedWriteRange(address: string): void {
  const dimensions = rangeDimensions(address);
  const cells = dimensions.rows * dimensions.columns;
  if (cells > MAX_WRITE_CELLS) throw new Error(`Refusing to write ${cells} cells; maximum is ${MAX_WRITE_CELLS}`);
}

export function columnNumber(column: string): number {
  return column.toUpperCase().split('').reduce((total, char) => total * 26 + char.charCodeAt(0) - 64, 0);
}

function bareRangeAddress(address: string): string {
  const bare = address.includes('!') ? address.slice(address.lastIndexOf('!') + 1) : address;
  return bare.replaceAll('$', '').toUpperCase();
}

export function rangeDimensions(address: string): { rows: number; columns: number } {
  const normalized = bareRangeAddress(address);
  if (/^[A-Z]+:[A-Z]+$/.test(normalized) || /^\d+:\d+$/.test(normalized)) {
    throw new Error(`Refusing unbounded range: ${address}`);
  }
  const parts = normalized.split(':');
  if (parts.length > 2) throw new Error(`Range must contain at most one colon: ${address}`);
  const [start, end = start] = parts;
  const startMatch = /^([A-Z]+)(\d+)$/.exec(start ?? '');
  const endMatch = /^([A-Z]+)(\d+)$/.exec(end ?? '');
  if (!startMatch || !endMatch) throw new Error(`Cannot determine dimensions for range: ${address}`);
  const startColumn = columnNumber(startMatch[1] ?? '');
  const endColumn = columnNumber(endMatch[1] ?? '');
  const startRow = Number(startMatch[2]);
  const endRow = Number(endMatch[2]);
  if (startColumn < 1 || startColumn > MAX_EXCEL_COLUMNS || endColumn < 1 || endColumn > MAX_EXCEL_COLUMNS) {
    throw new Error(`Range column must be between A and XFD: ${address}`);
  }
  if (startRow < 1 || startRow > MAX_EXCEL_ROWS || endRow < 1 || endRow > MAX_EXCEL_ROWS) {
    throw new Error(`Range row must be between 1 and ${MAX_EXCEL_ROWS}: ${address}`);
  }
  if (endColumn < startColumn || endRow < startRow) throw new Error(`Range must not be reversed: ${address}`);
  if (startColumn === 1 && endColumn === MAX_EXCEL_COLUMNS) {
    throw new Error(`Refusing whole-grid column span: ${address}`);
  }
  if (startRow === 1 && endRow === MAX_EXCEL_ROWS) {
    throw new Error(`Refusing whole-grid row span: ${address}`);
  }
  return {
    rows: endRow - startRow + 1,
    columns: endColumn - startColumn + 1,
  };
}

export function assertMatrixMatchesRange(address: string, matrix: unknown): asserts matrix is unknown[][] {
  if (!Array.isArray(matrix) || !matrix.every((row) => Array.isArray(row))) {
    throw new Error('Range payload must be a 2D array');
  }
  const dimensions = rangeDimensions(address);
  const rows = matrix.length;
  const columns = rows === 0 ? 0 : matrix[0]?.length ?? 0;
  if (rows !== dimensions.rows || !matrix.every((row) => row.length === dimensions.columns) || columns !== dimensions.columns) {
    throw new Error(`Range payload dimensions ${rows}x${columns} do not match ${address} (${dimensions.rows}x${dimensions.columns})`);
  }
}

function isGraphCell(value: unknown): value is GraphCell {
  return value === null
    || typeof value === 'string'
    || typeof value === 'boolean'
    || (typeof value === 'number' && Number.isFinite(value));
}

function parseMatrixJson(option: '--values-json' | '--formulas-json', raw: string, address: string): GraphCell[][] {
  let value: unknown;
  try {
    value = JSON.parse(raw) as unknown;
  } catch (error) {
    const detail = error instanceof SyntaxError ? error.message : 'invalid JSON';
    throw new Error(`${option} must contain valid JSON: ${detail}`);
  }
  assertMatrixMatchesRange(address, value);
  for (let rowIndex = 0; rowIndex < value.length; rowIndex += 1) {
    const row = value[rowIndex] ?? [];
    for (let columnIndex = 0; columnIndex < row.length; columnIndex += 1) {
      if (!isGraphCell(row[columnIndex])) {
        throw new Error(`${option} cell [${rowIndex}][${columnIndex}] must be a string, boolean, null, or finite number`);
      }
    }
  }
  const matrix = value as GraphCell[][];
  if (!matrix.some((row) => row.some((cell) => cell !== null))) {
    throw new Error(`${option} must change at least one cell; null cells are left unchanged`);
  }
  return matrix;
}

export function parseRangeMutation(input: {
  address: string;
  valuesJson?: string;
  formulasJson?: string;
}): RangeMutation {
  const hasValues = input.valuesJson !== undefined;
  const hasFormulas = input.formulasJson !== undefined;
  if (hasValues === hasFormulas) throw new Error('Expected exactly one of --values-json or --formulas-json');
  return hasValues
    ? { kind: 'values', matrix: parseMatrixJson('--values-json', input.valuesJson ?? '', input.address) }
    : { kind: 'formulas', matrix: parseMatrixJson('--formulas-json', input.formulasJson ?? '', input.address) };
}

export function verifyRangeMutation(address: string, mutation: RangeMutation, response: unknown): MutationVerification {
  if (!response || typeof response !== 'object' || Array.isArray(response)) {
    throw new Error('Graph verification response must be an object');
  }
  const record = response as Record<string, unknown>;
  if (typeof record.address !== 'string') throw new Error('Graph verification response did not include an address');
  if (bareRangeAddress(record.address) !== bareRangeAddress(address)) {
    throw new Error(`Graph verification returned ${record.address}, expected ${address}`);
  }
  const returnedMatrix = record[mutation.kind];
  assertMatrixMatchesRange(address, returnedMatrix);
  const mismatches: MutationVerification['mismatches'] = [];
  for (let rowIndex = 0; rowIndex < mutation.matrix.length; rowIndex += 1) {
    const requestedRow = mutation.matrix[rowIndex] ?? [];
    const returnedRow = returnedMatrix[rowIndex] ?? [];
    for (let columnIndex = 0; columnIndex < requestedRow.length; columnIndex += 1) {
      const expected = requestedRow[columnIndex];
      if (expected !== null && returnedRow[columnIndex] !== expected) {
        mismatches.push({ row: rowIndex, column: columnIndex, expected, actual: returnedRow[columnIndex] });
      }
    }
  }
  return {
    matched: mismatches.length === 0,
    field: mutation.kind,
    address: record.address,
    mismatches,
  };
}

const workbookMutationQueues = new Map<string, Promise<unknown>>();

export async function serializeWorkbookMutation<T>(workbookKey: string, operation: () => Promise<T>): Promise<T> {
  const previous = workbookMutationQueues.get(workbookKey) ?? Promise.resolve();
  const current = previous.catch(() => undefined).then(operation);
  workbookMutationQueues.set(workbookKey, current);
  try {
    return await current;
  } finally {
    if (workbookMutationQueues.get(workbookKey) === current) workbookMutationQueues.delete(workbookKey);
  }
}
