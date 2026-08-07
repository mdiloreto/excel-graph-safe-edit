import { describe, expect, it } from 'vitest';
import { backupFileName } from '../src/backup.js';

describe('backup helpers', () => {
  it('creates safe timestamped backup names', () => {
    const name = backupFileName('Budget 2026.xlsx', new Date('2026-08-07T17:30:00.123Z'));
    expect(name).toBe('Budget_2026.xlsx.2026-08-07T17-30-00-123Z.backup.xlsx');
  });
});
