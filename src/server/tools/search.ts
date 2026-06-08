import type { ToolDefinition } from '../types';
import { ToolError } from '../types';
import { forwardTool, jsonResult } from './helpers';
import { checkNonEmpty } from '../validation';

export const search = forwardTool(
  'word_search',
  '[Search] Find text in the document. Returns match count and up to 30 matches. Query must be ≤255 chars. Supports Word search codes (^p = paragraph mark, ^t = tab, etc.) and wildcard mode (?, *, [], {n,m}). To search for literal "^", use "^^".',
  {
    properties: {
      query: { type: 'string' },
      matchCase: { type: 'boolean', description: 'Case-sensitive search. Default: false' },
      matchWholeWord: { type: 'boolean', description: 'Match whole words only. Default: false' },
      matchWildcards: { type: 'boolean', description: 'Enable wildcard/regex search (?, *, [], {n,m}, @). Default: false' },
      matchPrefix: { type: 'boolean', description: 'Match words that begin with the search string. Default: false' },
      matchSuffix: { type: 'boolean', description: 'Match words that end with the search string. Default: false' },
      ignorePunct: { type: 'boolean', description: 'Ignore punctuation between words when matching. Default: false' },
      ignoreSpace: { type: 'boolean', description: 'Ignore whitespace between words when matching. Default: false' },
    },
    required: ['query'],
  },
  'search',
  (args) => {
    checkNonEmpty(args.query, 'query');
  },
);

export const searchAndReplace: ToolDefinition = {
  name: 'word_search_and_replace',
  description: '[Search] Find and replace ALL occurrences. Supports Word search codes in both find and replace: ^p (paragraph mark), ^l (line break), ^m (page break), ^n (column break), ^t (tab), ^s (non-breaking space), ^~ (non-breaking hyphen), ^- (optional hyphen), ^+ (em dash), ^= (en dash), ^^ (literal caret). Supports wildcard search in find string. For single-paragraph edits, prefer word_replace_paragraph_text.',
  schema: {
    properties: {
      find: { type: 'string' },
      replace: { type: 'string' },
      matchCase: { type: 'boolean', description: 'Default: false' },
      matchWholeWord: { type: 'boolean', description: 'Match whole words only. Default: false' },
      matchWildcards: { type: 'boolean', description: 'Enable wildcard/regex search in find string (?, *, [], {n,m}, @). Default: false' },
      matchPrefix: { type: 'boolean', description: 'Match words that begin with the find string. Default: false' },
      matchSuffix: { type: 'boolean', description: 'Match words that end with the find string. Default: false' },
      ignorePunct: { type: 'boolean', description: 'Ignore punctuation between words when matching. Default: false' },
      ignoreSpace: { type: 'boolean', description: 'Ignore whitespace between words when matching. Default: false' },
      preserveBookmarks: { type: 'boolean', description: 'Re-create bookmarks on replacement text after replace. Default: false' },
    },
    required: ['find', 'replace'],
  },
  async handler(args, bridge) {
    const find = args.find as string;

    if (!find || typeof find !== 'string' || find.trim() === '') {
      throw new ToolError('find string cannot be empty.');
    }

    const result = await bridge.send<{ replacements: number; bookmarksLost?: number; bookmarksRestored?: number; warning?: string }>('searchAndReplace', args);
    return jsonResult(result);
  },
};

export const insertTextAtMatch = forwardTool(
  'word_insert_text_at_match',
  '[Search] Insert text before or after a search match. Anchor text (after/before) must be ≤255 chars. Supports Word search codes in inserted text: ^p (paragraph mark), ^l (line break), ^t (tab), ^s (non-breaking space), ^m (page break), ^n (column break), ^~ (non-breaking hyphen), ^- (optional hyphen), ^+ (em dash), ^= (en dash), ^^ (literal caret). Provide "after" OR "before" as the anchor text. Use occurrence for Nth match. Note: inserting ^p after hyperlinked text may create a paragraph that inherits the hyperlink character style.',
  {
    properties: {
      text: { type: 'string', description: 'Text to insert' },
      after: { type: 'string', description: 'Search for this text and insert AFTER it' },
      before: { type: 'string', description: 'Search for this text and insert BEFORE it' },
      occurrence: { type: 'number', description: '0=first, 1=second, etc. Default: 0' },
      matchCase: { type: 'boolean', description: 'Default: false' },
      matchWildcards: { type: 'boolean', description: 'Enable wildcard/regex search for anchor text. Default: false' },
      matchPrefix: { type: 'boolean', description: 'Match words that begin with the anchor text. Default: false' },
      matchSuffix: { type: 'boolean', description: 'Match words that end with the anchor text. Default: false' },
      ignorePunct: { type: 'boolean', description: 'Ignore punctuation between words when matching. Default: false' },
      ignoreSpace: { type: 'boolean', description: 'Ignore whitespace between words when matching. Default: false' },
    },
    required: ['text'],
  },
  'insertText',
  (args) => {
    checkNonEmpty(args.text, 'text');
  },
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
  (args) => {
    checkNonEmpty(args.anchorText, 'anchorText');
  },
);

export const searchTools: ToolDefinition[] = [
  search,
  searchAndReplace,
  insertTextAtMatch,
  getSelectionInfo,
  insertTextAtSelection,
  insertLineBreak,
];
