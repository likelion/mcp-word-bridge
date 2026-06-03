import type { ToolDefinition } from '../types';
import { ToolError } from '../types';
import { forwardTool, jsonResult } from './helpers';
import { checkNoSpecialCodes } from '../validation';

export const search = forwardTool(
  'word_search',
  '[Search] Find text in the document. Returns match count and up to 30 matches. Query must be ≤255 chars.',
  {
    properties: {
      query: { type: 'string' },
      matchCase: { type: 'boolean', description: 'Case-sensitive search. Default: false' },
      matchWholeWord: { type: 'boolean' },
    },
    required: ['query'],
  },
  'search',
);

export const searchAndReplace: ToolDefinition = {
  name: 'word_search_and_replace',
  description: '[Search] Find and replace ALL occurrences. For single-paragraph edits, prefer word_replace_paragraph_text.',
  schema: {
    properties: {
      find: { type: 'string' },
      replace: { type: 'string' },
      matchCase: { type: 'boolean', description: 'Default: false' },
      matchWholeWord: { type: 'boolean' },
    },
    required: ['find', 'replace'],
  },
  async handler(args, bridge) {
    const find = args.find as string;
    const replace = args.replace as string;

    if (!find || typeof find !== 'string' || find.trim() === '') {
      throw new ToolError('find string cannot be empty.');
    }

    // BUG-01: Reject Word special codes that can corrupt document structure
    checkNoSpecialCodes(find, 'find');
    checkNoSpecialCodes(replace, 'replace');

    const result = await bridge.send('searchAndReplace', args);
    return jsonResult(result);
  },
};

export const insertTextAtMatch = forwardTool(
  'word_insert_text_at_match',
  '[Search] Insert text before or after a search match. Provide "after" OR "before" as the anchor text. Use occurrence for Nth match.',
  {
    properties: {
      text: { type: 'string', description: 'Text to insert' },
      after: { type: 'string', description: 'Search for this text and insert AFTER it' },
      before: { type: 'string', description: 'Search for this text and insert BEFORE it' },
      occurrence: { type: 'number', description: '0=first, 1=second, etc. Default: 0' },
      matchCase: { type: 'boolean', description: 'Default: false' },
    },
    required: ['text'],
  },
  'insertText',
);

export const getSelectionInfo = forwardTool(
  'word_get_selection_info',
  '[Search] Get the current cursor selection text with font and style details.',
  { properties: {} },
  'getSelectionInfo',
);

export const insertTextAtSelection = forwardTool(
  'word_insert_text_at_selection',
  '[Search] Insert text at the current cursor position, or replace the selection.',
  {
    properties: {
      text: { type: 'string' },
      replace: { type: 'boolean', description: 'Replace current selection. Default: false' },
    },
    required: ['text'],
  },
  'insertTextAtSelection',
);

export const insertLineBreak = forwardTool(
  'word_insert_line_break',
  '[Search] Insert a soft line break (Shift+Enter) before or after a text match.',
  {
    properties: {
      anchorText: { type: 'string', description: 'Text to search for' },
      before: { type: 'boolean', description: 'Insert before the match (default: after)' },
      occurrence: { type: 'number', description: '0=first, 1=second, etc. Default: 0' },
      matchCase: { type: 'boolean', description: 'Default: false' },
    },
    required: ['anchorText'],
  },
  'insertLineBreak',
);

export const searchTools: ToolDefinition[] = [
  search,
  searchAndReplace,
  insertTextAtMatch,
  getSelectionInfo,
  insertTextAtSelection,
  insertLineBreak,
];
