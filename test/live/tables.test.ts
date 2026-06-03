import { describe, test, expect, beforeAll, afterAll } from 'vitest';
import { getClient, skip, cleanup } from './setup';
import type { LiveTestClient } from './client';

let client: LiveTestClient;
beforeAll(async () => { client = await getClient(); }, 60000);
afterAll(cleanup);

describe('Tables', () => {
  beforeAll(async () => {
    if (skip()) return;
    await client.resetDocument();
    await client.call('word_insert_table', { rows: 3, cols: 3, data: [['A', 'B', 'C'], ['D', 'E', 'F'], ['G', 'H', 'I']] });
  });

  test('table exists with correct dimensions', async () => {
    if (skip()) return;
    const tables = await client.call('word_list_tables');
    expect(tables.count).toBe(1);
    expect(tables.tables[0].rowCount).toBe(3);
  });

  test('get table data matches input', async () => {
    if (skip()) return;
    const data = await client.call('word_get_table_data', { index: 0 });
    expect(data.values[0]).toEqual(['A', 'B', 'C']);
    expect(data.values[2][2]).toBe('I');
  });

  test('set cell and read back', async () => {
    if (skip()) return;
    await client.call('word_set_table_cell', { tableIndex: 0, row: 1, col: 1, text: 'UPDATED' });
    const data = await client.call('word_get_table_data', { index: 0 });
    expect(data.values[1][1]).toBe('UPDATED');
  });

  test('add and delete row', async () => {
    if (skip()) return;
    await client.call('word_add_table_row', { tableIndex: 0, values: ['X', 'Y', 'Z'] });
    let tables = await client.call('word_list_tables');
    expect(tables.tables[0].rowCount).toBe(4);
    await client.call('word_delete_table_row', { tableIndex: 0, rowIndex: 3 });
    tables = await client.call('word_list_tables');
    expect(tables.tables[0].rowCount).toBe(3);
  });

  test('merge cells', async () => {
    if (skip()) return;
    const r = await client.call('word_merge_table_cells', { tableIndex: 0, topRow: 0, firstCell: 0, bottomRow: 1, lastCell: 0 });
    expect(r.success).toBe(true);
  });

  test('insert paragraph at table index warns', async () => {
    if (skip()) return;
    const paras = await client.call('word_get_paragraphs');
    const tableIdx = paras.paragraphs.findIndex((p: any) => p.inTable);
    if (tableIdx === -1) return;
    const result = await client.call('word_insert_paragraph_at_index', { index: tableIdx, text: 'In cell' });
    expect(result.warning).toContain('table cell');
  });
});
