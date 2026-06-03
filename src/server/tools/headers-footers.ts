import type { ToolDefinition } from '../types';
import { forwardTool } from './helpers';

export const getHeaderFooter = forwardTool(
  'word_get_header_footer',
  '[Headers/Footers] Get header or footer text content.',
  {
    properties: {
      type: { type: 'string', enum: ['header', 'footer'] },
      sectionIndex: { type: 'number', description: 'Section index (0-based). Default: 0' },
      headerType: { type: 'string', enum: ['Primary', 'FirstPage', 'EvenPages'], description: 'Default: Primary' },
    },
    required: ['type'],
  },
  'getHeaderFooter',
);

export const setHeaderFooter = forwardTool(
  'word_set_header_footer',
  '[Headers/Footers] Set header or footer text (replaces existing content).',
  {
    properties: {
      type: { type: 'string', enum: ['header', 'footer'] },
      text: { type: 'string' },
      sectionIndex: { type: 'number', description: 'Section index (0-based). Default: 0' },
      headerType: { type: 'string', enum: ['Primary', 'FirstPage', 'EvenPages'], description: 'Default: Primary' },
    },
    required: ['type', 'text'],
  },
  'setHeaderFooter',
);

export const headerFooterTools: ToolDefinition[] = [
  getHeaderFooter,
  setHeaderFooter,
];
