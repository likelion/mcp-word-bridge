import type { ToolDefinition } from '../types';
import { forwardTool } from './helpers';

export const insertHtml = forwardTool(
  'word_insert_html',
  '[Advanced] Insert HTML content that Word converts to native formatting. Supports headings, bold, italic, links, tables, and lists.',
  {
    properties: {
      html: { type: 'string' },
      location: { type: 'string', enum: ['Start', 'End'], description: 'Default: End' },
    },
    required: ['html'],
  },
  'insertHtml',
);

export const insertOoxml = forwardTool(
  'word_insert_ooxml',
  '[Advanced] Insert raw Office Open XML for precise formatting control when HTML is insufficient.',
  {
    properties: {
      ooxml: { type: 'string' },
      location: { type: 'string', enum: ['Start', 'End'], description: 'Default: End' },
    },
    required: ['ooxml'],
  },
  'insertOoxml',
);

export const insertTableOfContents = forwardTool(
  'word_insert_table_of_contents',
  '[Advanced] Insert a table of contents based on heading styles.',
  {
    properties: {
      location: { type: 'string', enum: ['Start', 'End'], description: 'Default: Start' },
      switches: { type: 'string', description: 'TOC field switches. Default: \\o "1-3" \\h \\z \\u' },
    },
  },
  'insertTableOfContents',
);

export const advancedTools: ToolDefinition[] = [
  insertHtml,
  insertOoxml,
  insertTableOfContents,
];
