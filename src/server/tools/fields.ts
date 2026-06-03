import type { ToolDefinition } from '../types';
import { forwardTool } from './helpers';

export const getFields = forwardTool(
  'word_get_fields',
  '[Advanced] Get all fields in the document (hyperlinks, TOC entries, page numbers, etc).',
  { properties: {} },
  'getFields',
);

export const fieldTools: ToolDefinition[] = [
  getFields,
];
