import type { ToolDefinition } from '../types';
import { forwardTool } from './helpers';

export const insertTable = forwardTool(
  'word_insert_table',
  '[Tables] Insert a table with data. Provide rows (1-500), cols (1-63), and data as array of arrays.',
  {
    properties: {
      rows: { type: 'number' },
      cols: { type: 'number' },
      data: { type: 'array', items: { type: 'array', items: { type: 'string' } }, description: 'Cell values as array of row arrays' },
      location: { type: 'string', enum: ['Start', 'End'] },
      style: { type: 'string', description: 'Table style name' },
      headerRowCount: { type: 'number' },
    },
    required: ['rows', 'cols'],
  },
  'insertTable',
);

export const listTables = forwardTool(
  'word_list_tables',
  '[Tables] List table metadata: count, rowCount, columnCount, style. Does NOT return cell values.',
  { properties: {} },
  'getTables',
);

export const getTableData = forwardTool(
  'word_get_table_data',
  '[Tables] Get all cell values from a specific table as a 2D array.',
  {
    properties: {
      index: { type: 'number', description: 'Table index (0-based)' },
    },
    required: ['index'],
  },
  'getTableData',
);

export const setTableCell = forwardTool(
  'word_set_table_cell',
  '[Tables] Set text in a specific table cell.',
  {
    properties: {
      tableIndex: { type: 'number', description: 'Table index (0-based)' },
      row: { type: 'number', description: 'Row index (0-based)' },
      col: { type: 'number', description: 'Column index (0-based)' },
      text: { type: 'string' },
    },
    required: ['tableIndex', 'row', 'col', 'text'],
  },
  'setTableCell',
);

export const addTableRow = forwardTool(
  'word_add_table_row',
  '[Tables] Add a row to a table with optional cell values.',
  {
    properties: {
      tableIndex: { type: 'number', description: 'Table index (0-based)' },
      values: { type: 'array', items: { type: 'string' }, description: 'Cell values for the new row' },
      location: { type: 'string', enum: ['Start', 'End'], description: 'Default: End' },
    },
    required: ['tableIndex'],
  },
  'addTableRow',
);

export const deleteTableRow = forwardTool(
  'word_delete_table_row',
  '[Tables] Delete a row from a table.',
  {
    properties: {
      tableIndex: { type: 'number', description: 'Table index (0-based)' },
      rowIndex: { type: 'number', description: 'Row index (0-based)' },
    },
    required: ['tableIndex', 'rowIndex'],
  },
  'deleteTableRow',
);

export const mergeTableCells = forwardTool(
  'word_merge_table_cells',
  '[Tables] Merge a rectangular range of cells. All indices 0-based.',
  {
    properties: {
      tableIndex: { type: 'number' },
      topRow: { type: 'number' },
      firstCell: { type: 'number' },
      bottomRow: { type: 'number' },
      lastCell: { type: 'number' },
    },
    required: ['tableIndex', 'topRow', 'firstCell', 'bottomRow', 'lastCell'],
  },
  'mergeTableCells',
);

export const splitTableCell = forwardTool(
  'word_split_table_cell',
  '[Tables] Split a table cell into multiple rows/columns.',
  {
    properties: {
      tableIndex: { type: 'number' },
      row: { type: 'number' },
      col: { type: 'number' },
      rowCount: { type: 'number', description: 'Split into this many rows. Default: 1' },
      colCount: { type: 'number', description: 'Split into this many columns. Default: 2' },
    },
    required: ['tableIndex', 'row', 'col'],
  },
  'splitTableCell',
);

export const setTableStyle = forwardTool(
  'word_set_table_style',
  '[Tables] Apply a built-in table style.',
  {
    properties: {
      tableIndex: { type: 'number' },
      style: { type: 'string' },
    },
    required: ['tableIndex', 'style'],
  },
  'setTableStyle',
);

export const setTableCellShading = forwardTool(
  'word_set_table_cell_shading',
  '[Tables] Set background color on a table cell. Color must be hex.',
  {
    properties: {
      tableIndex: { type: 'number' },
      row: { type: 'number' },
      col: { type: 'number' },
      color: { type: 'string', description: 'Hex color e.g. #FFD700' },
    },
    required: ['tableIndex', 'row', 'col', 'color'],
  },
  'setTableCellShading',
);

export const tableTools: ToolDefinition[] = [
  insertTable,
  listTables,
  getTableData,
  setTableCell,
  addTableRow,
  deleteTableRow,
  mergeTableCells,
  splitTableCell,
  setTableStyle,
  setTableCellShading,
];
