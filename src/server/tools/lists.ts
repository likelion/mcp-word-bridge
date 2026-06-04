import type { ToolDefinition } from '../types';
import { forwardTool } from './helpers';

export const insertList = forwardTool(
  'word_insert_list',
  '[Lists] Insert a bulleted or numbered list from an array of item strings. Text is inserted literally (Word search codes like ^p or ^t are NOT interpreted).',
  {
    properties: {
      items: { type: 'array', items: { type: 'string' }, description: 'List item strings' },
      numbered: { type: 'boolean', description: 'true = numbered, false/omit = bulleted' },
      location: { type: 'string', enum: ['Start', 'End'] },
    },
    required: ['items'],
  },
  'insertList',
);

export const getListInfo = forwardTool(
  'word_get_list_info',
  '[Lists] Get list formatting details for a paragraph by index. Returns isListItem:false if not in a list.',
  {
    properties: {
      index: { type: 'number', description: 'Paragraph index (0-based)' },
    },
    required: ['index'],
  },
  'getListInfo',
);

export const setListLevel = forwardTool(
  'word_set_list_level',
  '[Lists] Set indent level of a list item (0=top level, up to 8).',
  {
    properties: {
      index: { type: 'number', description: 'Paragraph index (0-based)' },
      level: { type: 'number', description: 'Indent level: 0-8' },
    },
    required: ['index', 'level'],
  },
  'setListLevel',
);

export const listTools: ToolDefinition[] = [
  insertList,
  getListInfo,
  setListLevel,
];
