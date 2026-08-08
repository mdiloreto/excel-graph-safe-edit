import { describe, expect, it } from 'vitest';
import { GRAPH_BASE } from '../src/config.js';
import {
  assertBoundedWriteRange,
  assertMatrixMatchesRange,
  itemPathByDrivePath,
  itemPathById,
  odataString,
  parseRangeMutation,
  rangeDimensions,
  searchPath,
  serializeWorkbookMutation,
  verifyRangeMutation,
  workbookRangePath,
  worksheetsPath,
  worksheetTablesPath,
} from '../src/excel.js';

describe('excel graph paths', () => {
  it('escapes OData strings and OneDrive path segments', () => {
    expect(odataString("Bob's #1? 東京")).toBe("Bob%27%27s%20%231%3F%20%E6%9D%B1%E4%BA%AC");
    expect(itemPathByDrivePath('/Folder #1/Budget ? 2026.xlsx')).toBe('/me/drive/root:/Folder%20%231/Budget%20%3F%202026.xlsx:');
  });

  it('builds default-drive and explicit-drive item paths', () => {
    expect(itemPathById('item id')).toBe('/me/drive/items/item%20id');
    expect(itemPathById('item id', 'drive/id')).toBe('/drives/drive%2Fid/items/item%20id');
    expect(worksheetsPath('item', 'drive')).toBe('/drives/drive/items/item/workbook/worksheets');
    expect(worksheetTablesPath('item', 'Sheet #1', 'drive')).toBe('/drives/drive/items/item/workbook/worksheets/Sheet%20%231/tables');
    expect(() => itemPathById('item', '')).toThrow(/Drive ID must not be empty/);
  });

  it('builds drive-aware workbook range paths', () => {
    expect(workbookRangePath('item id', 'Dashboard', 'A1:B2')).toBe("/me/drive/items/item%20id/workbook/worksheets/Dashboard/range(address='A1%3AB2')");
    expect(workbookRangePath('item id', 'Dashboard', 'A1:B2', 'drive id')).toBe("/drives/drive%20id/items/item%20id/workbook/worksheets/Dashboard/range(address='A1%3AB2')");
  });

  it('does not let OData values create fragments or unintended queries', () => {
    const rangeUrl = new URL(`${GRAPH_BASE}${workbookRangePath('item?#', 'Bob\'s #? 東京', "A1:B2#?' 東京", 'drive?#')}`);
    expect(rangeUrl.hash).toBe('');
    expect(rangeUrl.search).toBe('');
    expect(rangeUrl.pathname).toContain('drive%3F%23');
    expect(rangeUrl.pathname).toContain('Bob%27s%20%23%3F%20%E6%9D%B1%E4%BA%AC');
    expect(rangeUrl.pathname).toContain('A1%3AB2%23%3F%27%27%20%E6%9D%B1%E4%BA%AC');

    const searchUrl = new URL(`${GRAPH_BASE}${searchPath("Bob's #? 東京")}`);
    expect(searchUrl.hash).toBe('');
    expect(searchUrl.searchParams.get('$select')).toBe('id,name,size,webUrl,parentReference,file,folder');
    expect(searchUrl.pathname).toContain('Bob%27%27s%20%23%3F%20%E6%9D%B1%E4%BA%AC');
  });
});

describe('range safety', () => {
  it('rejects unbounded write ranges', () => {
    expect(() => assertBoundedWriteRange('A:A')).toThrow(/unbounded/);
    expect(() => assertBoundedWriteRange('1:1')).toThrow(/unbounded/);
    expect(() => assertBoundedWriteRange('Sheet1!A:F')).toThrow(/unbounded/);
    expect(() => assertBoundedWriteRange('A1:B2')).not.toThrow();
  });

  it('rejects malformed, out-of-grid, reversed, and oversized write ranges', () => {
    expect(() => assertBoundedWriteRange('A1:XFD1')).toThrow(/whole-grid column span/);
    expect(() => assertBoundedWriteRange('A1:A1048576')).toThrow(/whole-grid row span/);
    expect(() => assertBoundedWriteRange('A0')).toThrow(/row must be between/);
    expect(() => assertBoundedWriteRange('A1048577')).toThrow(/row must be between/);
    expect(() => assertBoundedWriteRange('XFE1')).toThrow(/column must be between/);
    expect(() => assertBoundedWriteRange('A1:B2:C3')).toThrow(/at most one colon/);
    expect(() => assertBoundedWriteRange('B2:A1')).toThrow(/reversed/);
    expect(() => assertBoundedWriteRange('A2:A1')).toThrow(/reversed/);
    expect(() => assertBoundedWriteRange('A1:A10001')).toThrow(/maximum is 10000/);
    expect(() => assertBoundedWriteRange('A1:B2')).not.toThrow();
    expect(() => assertBoundedWriteRange('A1:J1000')).not.toThrow();
    expect(() => assertBoundedWriteRange('XFD1048576')).not.toThrow();
  });

  it('computes range dimensions', () => {
    expect(rangeDimensions('A1:A1')).toEqual({ rows: 1, columns: 1 });
    expect(rangeDimensions('B2:D5')).toEqual({ rows: 4, columns: 3 });
    expect(rangeDimensions('Sheet1!AA10:AB11')).toEqual({ rows: 2, columns: 2 });
  });

  it('requires rectangular payload dimensions to match the target range', () => {
    expect(() => assertMatrixMatchesRange('A1:B2', [[1, 2], [3, 4]])).not.toThrow();
    expect(() => assertMatrixMatchesRange('A1:B2', [[1, 2]])).toThrow(/do not match/);
    expect(() => assertMatrixMatchesRange('A1:B2', [[1], [2]])).toThrow(/do not match/);
    expect(() => assertMatrixMatchesRange('A1:B2', [1, 2])).toThrow(/2D array/);
  });
});

describe('range mutation validation and verification', () => {
  it('requires exactly one valid scalar matrix and rejects no-op payloads', () => {
    expect(parseRangeMutation({ address: 'A1:B1', valuesJson: '[[1,"ok"]]' })).toEqual({ kind: 'values', matrix: [[1, 'ok']] });
    expect(() => parseRangeMutation({ address: 'A1', valuesJson: '[[1]]', formulasJson: '[["=1"]]' })).toThrow(/exactly one/);
    expect(() => parseRangeMutation({ address: 'A1' })).toThrow(/exactly one/);
    expect(() => parseRangeMutation({ address: 'A1', valuesJson: 'not-json' })).toThrow(/valid JSON/);
    expect(() => parseRangeMutation({ address: 'A1', valuesJson: '[[{}]]' })).toThrow(/cell \[0\]\[0\]/);
    expect(() => parseRangeMutation({ address: 'A1', valuesJson: '[[[]]]' })).toThrow(/cell \[0\]\[0\]/);
    expect(() => parseRangeMutation({ address: 'A1', valuesJson: '[[1e400]]' })).toThrow(/finite number/);
    expect(() => parseRangeMutation({ address: 'A1:B1', valuesJson: '[[null,null]]' })).toThrow(/at least one cell/);
  });

  it('compares requested non-null values and validates response shape', () => {
    const mutation = parseRangeMutation({ address: 'A1:B1', valuesJson: '[[null,"changed"]]' });
    expect(verifyRangeMutation('A1:B1', mutation, {
      address: "'Sheet 1'!$A$1:$B$1",
      values: [['untouched', 'changed']],
    })).toMatchObject({ matched: true, mismatches: [] });
    expect(verifyRangeMutation('A1:B1', mutation, {
      address: 'Sheet1!A1:B1',
      values: [['untouched', 'different']],
    })).toMatchObject({ matched: false, mismatches: [{ row: 0, column: 1 }] });
    expect(() => verifyRangeMutation('A1:B1', mutation, { address: 'Sheet1!A1:C1', values: [['x', 'changed']] })).toThrow(/expected/);
    expect(() => verifyRangeMutation('A1:B1', mutation, { address: 'Sheet1!A1:B1', values: [['changed']] })).toThrow(/dimensions/);
  });

  it('serializes mutation operations for the same workbook', async () => {
    const events: string[] = [];
    let releaseFirst: (() => void) | undefined;
    const first = serializeWorkbookMutation('drive:item', async () => {
      events.push('first-start');
      await new Promise<void>((resolve) => {
        releaseFirst = resolve;
      });
      events.push('first-end');
      return 1;
    });
    const second = serializeWorkbookMutation('drive:item', async () => {
      events.push('second-start');
      return 2;
    });
    await new Promise<void>((resolve) => setImmediate(resolve));
    expect(events).toEqual(['first-start']);
    releaseFirst?.();
    await expect(Promise.all([first, second])).resolves.toEqual([1, 2]);
    expect(events).toEqual(['first-start', 'first-end', 'second-start']);
  });
});
