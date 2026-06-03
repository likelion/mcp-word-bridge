import type { ToolDefinition } from '../types';
import { forwardTool } from './helpers';

export const getBookmarks = forwardTool(
  'word_get_bookmarks',
  '[Bookmarks] Get all bookmark names in the document.',
  { properties: {} },
  'getBookmarks',
);

export const insertBookmark = forwardTool(
  'word_insert_bookmark',
  '[Bookmarks] Create a named bookmark at a text match. If name already exists, the bookmark is moved.',
  {
    properties: {
      name: { type: 'string', description: 'Bookmark name (must be unique)' },
      anchorText: { type: 'string', description: 'Text to search for as bookmark location' },
      occurrence: { type: 'number', description: '0=first, 1=second, etc. Default: 0' },
      matchCase: { type: 'boolean', description: 'Default: false' },
    },
    required: ['name', 'anchorText'],
  },
  'insertBookmark',
);

export const deleteBookmark = forwardTool(
  'word_delete_bookmark',
  '[Bookmarks] Delete a bookmark by name (text remains, only the reference is removed).',
  {
    properties: {
      name: { type: 'string' },
    },
    required: ['name'],
  },
  'deleteBookmark',
);

export const goToBookmark = forwardTool(
  'word_go_to_bookmark',
  '[Bookmarks] Navigate to a bookmark and select its text range.',
  {
    properties: {
      name: { type: 'string' },
    },
    required: ['name'],
  },
  'goToBookmark',
);

export const getBookmarkText = forwardTool(
  'word_get_bookmark_text',
  '[Bookmarks] Get the text content within a named bookmark.',
  {
    properties: {
      name: { type: 'string' },
    },
    required: ['name'],
  },
  'getBookmarkText',
);

export const bookmarkTools: ToolDefinition[] = [
  getBookmarks,
  insertBookmark,
  deleteBookmark,
  goToBookmark,
  getBookmarkText,
];
