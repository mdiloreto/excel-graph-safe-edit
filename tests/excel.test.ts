import { describe, expect, it } from 'vitest';
import {
  assertBoundedWriteRange,
  assertMatrixMatchesRange,
  itemPathByDrivePath,
  odataString,
  rangeDimensions,
  searchPath,
  workbookRangePath,
} from '../src/excel.js';

describe('excel graph paths', () => {
  it('escapes OData strings and OneDrive paths', () => {
    expect(odataString("Bob's Budget")).toBe("Bob''s Budget");
    expect(itemPathByDrivePath('/Folder A/Budget 2026.xlsx')).toBe('/me/drive/root:/Folder%20A/Budget%202026.xlsx:');
  });

  it('builds workbook range paths', () => {
    expect(workbookRangePath('item id', 'Dashboard', "A1:B2")).toBe("/me/drive/items/item%20id/workbook/worksheets/Dashboard/range(address='A1:B2')");
    expect(searchPath("Bob's")).toContain("Bob''s");
  });
});

describe('range safety', () => {
  it('rejects unbounded write ranges', () => {
    expect(() => assertBoundedWriteRange('A:A')).toThrow(/unbounded/);
    expect(() => assertBoundedWriteRange('1:1')).toThrow(/unbounded/);
    expect(() => assertBoundedWriteRange('Sheet1!A:F')).toThrow(/unbounded/);
    expect(() => assertBoundedWriteRange('A1:B2')).not.toThrow();
  });

  it('computes range dimensions', () => {
    expect(rangeDimensions('A1:A1')).toEqual({ rows: 1, columns: 1 });
    expect(rangeDimensions('B2:D5')).toEqual({ rows: 4, columns: 3 });
    expect(rangeDimensions('Sheet1!AA10:AB11')).toEqual({ rows: 2, columns: 2 });
  });

  it('requires write payload dimensions to match the target range', () => {
    expect(() => assertMatrixMatchesRange('A1:B2', [[1, 2], [3, 4]])).not.toThrow();
    expect(() => assertMatrixMatchesRange('A1:B2', [[1, 2]])).toThrow(/do not match/);
    expect(() => assertMatrixMatchesRange('A1:B2', [[1], [2]])).toThrow(/do not match/);
    expect(() => assertMatrixMatchesRange('A1:B2', [1, 2])).toThrow(/2D array/);
  });
});
