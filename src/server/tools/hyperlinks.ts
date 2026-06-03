import type { ToolDefinition } from '../types';
import { forwardTool } from './helpers';

export const insertHyperlink = forwardTool(
  'word_insert_hyperlink',
  '[Hyperlinks] Add a hyperlink URL to existing text (searches for anchorText, applies the link).',
  {
    properties: {
      anchorText: { type: 'string', description: 'Text to search for and link' },
      url: { type: 'string', description: 'URL (must start with http:// or https://)' },
      occurrence: { type: 'number', description: '0=first, 1=second, etc. Default: 0' },
      matchCase: { type: 'boolean', description: 'Default: false' },
    },
    required: ['anchorText', 'url'],
  },
  'insertHyperlink',
);

export const getHyperlinks = forwardTool(
  'word_get_hyperlinks',
  '[Hyperlinks] List all hyperlinks with URL, display text, and internal flag.',
  { properties: {} },
  'getHyperlinks',
);

export const removeHyperlink = forwardTool(
  'word_remove_hyperlink',
  '[Hyperlinks] Remove a hyperlink from text (keeps the text, removes only the link).',
  {
    properties: {
      anchorText: { type: 'string', description: 'Text of the hyperlink to remove' },
      occurrence: { type: 'number', description: '0=first, 1=second, etc. Default: 0' },
      matchCase: { type: 'boolean', description: 'Default: false' },
    },
    required: ['anchorText'],
  },
  'removeHyperlink',
);

export const hyperlinkTools: ToolDefinition[] = [
  insertHyperlink,
  getHyperlinks,
  removeHyperlink,
];
