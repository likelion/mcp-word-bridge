import type { ToolDefinition } from '../types';
import { ToolError } from '../types';
import { forwardTool } from './helpers';
import { checkNonNegative, checkHexColor } from '../validation';

export const insertTable = forwardTool(
  'word_insert_table',
  '[Tables] Insert a table with data. Provide rows (1-500), cols (1-63), and data as array of arrays.',
  {
    properties: {
      rows: { type: 'number' },
      cols: { type: 'number' },
      data: { type: 'array', items: { type: 'array', items: { type: 'string' } }, description: 'Cell values as array of row arrays' },
      location: { type: 'string', enum: ['Start', 'End'] },
      style: { type: 'string', description: 'Built-in Word table style name, e.g. "Grid Table 4 - Accent 1".' },
      headerRowCount: { type: 'number', description: 'Number of header rows (default: 0). Set to 1 to mark the first row as a repeating header.' },
      caption: { type: 'string', description: 'Optional caption text. Adds an auto-numbered "Table N: <caption>" caption paragraph above the table.' },
    },
    required: ['rows', 'cols'],
  },
  'insertTable',
  (args) => {
    const rows = args.rows as number;
    const cols = args.cols as number;
    if (typeof rows !== 'number' || !Number.isInteger(rows) || rows <= 0) throw new ToolError('rows must be a positive integer (minimum 1).');
    if (typeof cols !== 'number' || !Number.isInteger(cols) || cols <= 0) throw new ToolError('cols must be a positive integer (minimum 1).');
    if (cols > 63) throw new ToolError('cols must not exceed 63 (Word maximum column limit).');
    if (rows > 500) throw new ToolError('rows must not exceed 500 (practical limit for performance).');
    if (args.data !== undefined) {
      const data = args.data as unknown[];
      if (!Array.isArray(data)) throw new ToolError('data must be an array of arrays.');
      if (data.length !== rows) throw new ToolError(`Data rows (${data.length}) do not match specified rows (${rows}). Provide exactly ${rows} row(s) in the data array.`);
      for (let i = 0; i < data.length; i++) {
        if (!Array.isArray(data[i])) throw new ToolError(`data[${i}] must be an array.`);
        if ((data[i] as unknown[]).length !== cols) throw new ToolError(`Data row ${i} has ${(data[i] as unknown[]).length} columns but expected ${cols}.`);
      }
    }
  },
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
  (args) => {
    checkNonNegative(args.tableIndex, 'tableIndex');
    checkNonNegative(args.row, 'row');
    checkNonNegative(args.col, 'col');
  },
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
  (args) => {
    checkNonNegative(args.tableIndex, 'tableIndex');
    checkNonNegative(args.rowIndex, 'rowIndex');
  },
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
  (args) => {
    checkNonNegative(args.tableIndex, 'tableIndex');
    checkNonNegative(args.topRow, 'topRow');
    checkNonNegative(args.firstCell, 'firstCell');
    checkNonNegative(args.bottomRow, 'bottomRow');
    checkNonNegative(args.lastCell, 'lastCell');
    if ((args.topRow as number) > (args.bottomRow as number)) throw new ToolError(`topRow (${args.topRow}) must be less than or equal to bottomRow (${args.bottomRow}).`);
    if ((args.firstCell as number) > (args.lastCell as number)) throw new ToolError(`firstCell (${args.firstCell}) must be less than or equal to lastCell (${args.lastCell}).`);
    if (args.topRow === args.bottomRow && args.firstCell === args.lastCell) throw new ToolError('Cannot merge a single cell with itself. Provide a range spanning at least 2 cells.');
  },
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
  (args) => {
    checkNonNegative(args.tableIndex, 'tableIndex');
    checkNonNegative(args.row, 'row');
    checkNonNegative(args.col, 'col');
    if (args.rowCount !== undefined && (typeof args.rowCount !== 'number' || args.rowCount <= 0)) throw new ToolError('rowCount must be a positive integer.');
    if (args.colCount !== undefined && (typeof args.colCount !== 'number' || args.colCount <= 0)) throw new ToolError('colCount must be a positive integer.');
  },
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
  (args) => {
    checkNonNegative(args.tableIndex, 'tableIndex');
    checkNonNegative(args.row, 'row');
    checkNonNegative(args.col, 'col');
    checkHexColor(args.color as string, 'color');
  },
);

export const deleteTable = forwardTool(
  'word_delete_table',
  '[Tables] Delete an entire table by its 0-based index, including its caption paragraph if present.',
  {
    properties: {
      index: { type: 'number', description: 'Table index (0-based)' },
      deleteCaption: { type: 'boolean', description: 'Also delete the "Caption"-styled paragraph directly above the table. Default: true.' },
    },
    required: ['index'],
  },
  'deleteTable',
  (args) => {
    checkNonNegative(args.index, 'index');
  },
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
  deleteTable,
];
