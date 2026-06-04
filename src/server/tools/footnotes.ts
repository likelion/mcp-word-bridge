import type { ToolDefinition } from '../types';
import { forwardTool } from './helpers';
import { checkNonEmpty } from '../validation';

export const insertFootnote = forwardTool(
  'word_insert_footnote',
  '[Footnotes] Insert a footnote anchored to a text match. Note: multiple footnotes on the same anchor appear in reverse insertion order (most recent first). Use word_insert_footnote_at_index for explicit placement.',
  {
    properties: {
      anchorText: { type: 'string', description: 'Text to search for as anchor point' },
      text: { type: 'string', description: 'Footnote content' },
      occurrence: { type: 'number', description: '0=first, 1=second, etc. Default: 0' },
      matchCase: { type: 'boolean', description: 'Default: false' },
    },
    required: ['anchorText', 'text'],
  },
  'insertFootnote',
  (args) => {
    checkNonEmpty(args.anchorText, 'anchorText');
    checkNonEmpty(args.text, 'text');
  },
);

export const insertFootnoteAtIndex = forwardTool(
  'word_insert_footnote_at_index',
  '[Footnotes] Insert a footnote at the end of a paragraph by index.',
  {
    properties: {
      paragraphIndex: { type: 'number', description: 'Paragraph index (0-based)' },
      text: { type: 'string', description: 'Footnote content' },
    },
    required: ['paragraphIndex', 'text'],
  },
  'insertFootnoteAtIndex',
  (args) => {
    checkNonEmpty(args.text, 'text');
  },
);

export const insertEndnote = forwardTool(
  'word_insert_endnote',
  '[Footnotes] Insert an endnote anchored to a text match.',
  {
    properties: {
      anchorText: { type: 'string', description: 'Text to search for as anchor point' },
      text: { type: 'string', description: 'Endnote content' },
      occurrence: { type: 'number', description: '0=first, 1=second, etc. Default: 0' },
      matchCase: { type: 'boolean', description: 'Default: false' },
    },
    required: ['anchorText', 'text'],
  },
  'insertEndnote',
  (args) => {
    checkNonEmpty(args.anchorText, 'anchorText');
    checkNonEmpty(args.text, 'text');
  },
);

export const getFootnotes = forwardTool(
  'word_get_footnotes',
  '[Footnotes] Get all footnotes with index and text content.',
  { properties: {} },
  'getFootnotes',
);

export const getEndnotes = forwardTool(
  'word_get_endnotes',
  '[Footnotes] Get all endnotes with index and text content.',
  { properties: {} },
  'getEndnotes',
);

export const deleteFootnote = forwardTool(
  'word_delete_footnote',
  '[Footnotes] Delete a footnote by its 0-based index.',
  {
    properties: {
      index: { type: 'number', description: 'Footnote index (0-based)' },
    },
    required: ['index'],
  },
  'deleteFootnote',
);

export const deleteEndnote = forwardTool(
  'word_delete_endnote',
  '[Footnotes] Delete an endnote by its 0-based index.',
  {
    properties: {
      index: { type: 'number', description: 'Endnote index (0-based)' },
    },
    required: ['index'],
  },
  'deleteEndnote',
);

export const footnoteTools: ToolDefinition[] = [
  insertFootnote,
  insertFootnoteAtIndex,
  insertEndnote,
  getFootnotes,
  getEndnotes,
  deleteFootnote,
  deleteEndnote,
];
