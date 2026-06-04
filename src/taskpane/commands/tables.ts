import type { CommandHandler } from './index';
import { checkHexColor, checkIndex } from './document';

declare const Word: any;

export const tableCommands: Record<string, CommandHandler> = {
  async insertTable(ctx, p) {
    if (!p.rows || p.rows <= 0) throw new Error('rows must be a positive integer (minimum 1)');
    if (!p.cols || p.cols <= 0) throw new Error('cols must be a positive integer (minimum 1)');
    if (p.cols > 63) throw new Error('cols must not exceed 63 (Word maximum column limit)');
    if (p.rows > 500) throw new Error('rows must not exceed 500 (practical limit for performance)');
    if (p.data) {
      if (!Array.isArray(p.data)) throw new Error('data must be an array of arrays');
      if (p.data.length !== p.rows)
        throw new Error(`Data rows (${p.data.length}) do not match specified rows (${p.rows}). Provide exactly ${p.rows} row(s) in the data array.`);
      for (let i = 0; i < p.data.length; i++) {
        if (!Array.isArray(p.data[i])) throw new Error(`data[${i}] must be an array`);
        if (p.data[i].length !== p.cols) throw new Error(`Data row ${i} has ${p.data[i].length} columns but expected ${p.cols}.`);
      }
    }
    const body = ctx.document.body;
    const table = body.insertTable(p.rows, p.cols, p.location || 'End', p.data || null);
    if (p.style) table.style = p.style;
    table.headerRowCount = p.headerRowCount ?? 0;
    await ctx.sync();
    const tableRange = table.getRange('End');
    tableRange.select();
    await ctx.sync();
    return { success: true };
  },

  async getTables(ctx) {
    const tables = ctx.document.body.tables;
    tables.load('rowCount,values,style,headerRowCount');
    await ctx.sync();
    const items = tables.items.map((t: any, i: number) => ({
      index: i, rowCount: t.rowCount, columnCount: (t.values && t.values[0]) ? t.values[0].length : 0, style: t.style, headerRowCount: t.headerRowCount,
    }));
    return { count: items.length, tables: items };
  },

  async getTableData(ctx, p) {
    if (p.index < 0) throw new Error('Table index must be non-negative');
    const tables = ctx.document.body.tables;
    tables.load('rowCount,values');
    await ctx.sync();
    if (p.index >= tables.items.length) throw new Error(`Table index out of range. Document has ${tables.items.length} table(s).`);
    return { values: tables.items[p.index].values };
  },

  async setTableCell(ctx, p) {
    const tables = ctx.document.body.tables;
    tables.load('rowCount');
    await ctx.sync();
    if (p.tableIndex >= tables.items.length) throw new Error(`Table index out of range. Document has ${tables.items.length} table(s).`);
    const table = tables.items[p.tableIndex];
    try {
      const cell = table.getCell(p.row, p.col);
      cell.body.clear();
      const inserted = cell.body.insertText(p.text, Word.InsertLocation.start);
      inserted.getRange('End').select();
      await ctx.sync();
    } catch (e: any) {
      if (e.message?.includes('ItemNotFound'))
        throw new Error(`Cell not found at row ${p.row}, col ${p.col}. Table has ${table.rowCount} rows. Use word_get_table_data to inspect the table.`);
      throw e;
    }
    return { success: true };
  },

  async addTableRow(ctx, p) {
    if (p.tableIndex < 0) throw new Error('Table index must be non-negative');
    const tables = ctx.document.body.tables;
    tables.load('rowCount,values');
    await ctx.sync();
    if (p.tableIndex >= tables.items.length) throw new Error('Table index out of range');
    const table = tables.items[p.tableIndex];
    const colCount = table.values[0].length;
    if (p.values && p.values.length > 0) {
      if (p.values.length > colCount) throw new Error(`values has ${p.values.length} items but table only has ${colCount} columns.`);
    }
    table.addRows(p.location || 'End', 1, p.values ? [p.values] : undefined);
    await ctx.sync();
    const tableRange = table.getRange('End');
    tableRange.select();
    await ctx.sync();
    const result: any = { success: true };
    // Warn when fewer values than columns are provided
    if (p.values && p.values.length > 0 && p.values.length < colCount) {
      result.warning = `Only ${p.values.length} of ${colCount} cells populated. Remaining cells left empty.`;
    }
    return result;
  },

  async deleteTableRow(ctx, p) {
    if (p.tableIndex < 0) throw new Error('Table index must be non-negative');
    if (p.rowIndex < 0) throw new Error('Row index must be non-negative');
    const tables = ctx.document.body.tables;
    tables.load('rowCount');
    await ctx.sync();
    if (p.tableIndex >= tables.items.length) throw new Error(`Table index out of range. Document has ${tables.items.length} table(s).`);
    const rows = tables.items[p.tableIndex].rows;
    rows.load('items');
    await ctx.sync();
    if (p.rowIndex >= rows.items.length) throw new Error(`Row index out of range. Table has ${rows.items.length} rows (0-indexed).`);
    rows.items[p.rowIndex].delete();
    await ctx.sync();
    return { success: true };
  },

  async mergeTableCells(ctx, p) {
    if (p.topRow < 0 || p.bottomRow < 0 || p.firstCell < 0 || p.lastCell < 0)
      throw new Error('All cell indices must be non-negative');
    if (p.topRow > p.bottomRow) throw new Error(`topRow (${p.topRow}) must be less than or equal to bottomRow (${p.bottomRow})`);
    if (p.firstCell > p.lastCell) throw new Error(`firstCell (${p.firstCell}) must be less than or equal to lastCell (${p.lastCell})`);
    if (p.topRow === p.bottomRow && p.firstCell === p.lastCell)
      throw new Error('Cannot merge a single cell with itself. Provide a range spanning at least 2 cells.');
    if (p.tableIndex < 0) throw new Error('Table index must be non-negative');
    const tables = ctx.document.body.tables;
    tables.load('rowCount');
    await ctx.sync();
    if (p.tableIndex >= tables.items.length) throw new Error('Table index out of range');
    const table = tables.items[p.tableIndex];
    if (p.bottomRow >= table.rowCount) throw new Error(`bottomRow (${p.bottomRow}) exceeds table row count (${table.rowCount})`);
    try {
      table.mergeCells(p.topRow, p.firstCell, p.bottomRow, p.lastCell);
      await ctx.sync();
    } catch (e: any) {
      if (e.message?.includes('InvalidArgument'))
        throw new Error(`Cannot merge cells: range is out of bounds. Table has ${table.rowCount} rows. Use word_get_table_data to inspect the table.`);
      throw e;
    }
    return { success: true };
  },

  async splitTableCell(ctx, p) {
    if (p.rowCount !== undefined && p.rowCount <= 0) throw new Error('rowCount must be a positive integer (minimum 1)');
    if (p.colCount !== undefined && p.colCount <= 0) throw new Error('colCount must be a positive integer (minimum 1)');
    if (p.tableIndex < 0) throw new Error('Table index must be non-negative');
    if (p.row < 0) throw new Error('Row index must be non-negative');
    if (p.col < 0) throw new Error('Column index must be non-negative');
    const tables = ctx.document.body.tables;
    tables.load('rowCount');
    await ctx.sync();
    if (p.tableIndex >= tables.items.length) throw new Error(`Table index out of range. Document has ${tables.items.length} table(s).`);
    try {
      const cell = tables.items[p.tableIndex].getCell(p.row, p.col);
      cell.split(p.rowCount || 1, p.colCount || 2);
      await ctx.sync();
    } catch (e: any) {
      if (e.message?.includes('ItemNotFound'))
        throw new Error(`Cell not found at row ${p.row}, col ${p.col}. Use word_get_table_data to inspect the table.`);
      throw e;
    }
    return { success: true };
  },

  async setTableStyle(ctx, p) {
    if (p.tableIndex < 0) throw new Error('Table index must be non-negative');
    const tables = ctx.document.body.tables;
    tables.load('rowCount');
    await ctx.sync();
    if (p.tableIndex >= tables.items.length) throw new Error('Table index out of range');
    try {
      tables.items[p.tableIndex].style = p.style;
      await ctx.sync();
    } catch (e: any) {
      if (e.message?.includes('InvalidArgument'))
        throw new Error(`Table style not found: "${p.style}". Use a built-in style name like "Grid Table 4 - Accent 1".`);
      throw e;
    }
    return { success: true };
  },

  async setTableCellShading(ctx, p) {
    checkHexColor(p.color, 'color');
    checkIndex(p.row, 'row');
    checkIndex(p.col, 'col');
    const tables = ctx.document.body.tables;
    tables.load('rowCount,values');
    await ctx.sync();
    if (p.tableIndex >= tables.items.length) throw new Error(`Table index out of range. Document has ${tables.items.length} table(s).`);
    const table = tables.items[p.tableIndex];
    const rowCount = table.rowCount;
    const colCount = (table.values && table.values[0]) ? table.values[0].length : 0;
    if (p.row >= rowCount)
      throw new Error(`Cell not found at row ${p.row}, col ${p.col}. Table has ${rowCount} row(s) and ${colCount} column(s).`);
    if (p.col >= colCount)
      throw new Error(`Cell not found at row ${p.row}, col ${p.col}. Table has ${rowCount} row(s) and ${colCount} column(s).`);
    try {
      const cell = table.getCell(p.row, p.col);
      cell.shadingColor = p.color;
      cell.body.getRange('End').select();
      await ctx.sync();
    } catch (e: any) {
      if (e.message?.includes('ItemNotFound'))
        throw new Error(`Cell not found at row ${p.row}, col ${p.col}. Table has ${rowCount} row(s) and ${colCount} column(s). The cell may be part of a merged range.`);
      throw e;
    }
    return { success: true };
  },
};
