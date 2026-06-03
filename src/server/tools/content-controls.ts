import type { ToolDefinition } from '../types';
import { forwardTool } from './helpers';

export const getContentControls = forwardTool(
  'word_get_content_controls',
  '[Content Controls] Get all content controls with id, tag, title, type, and text.',
  { properties: {} },
  'getContentControls',
);

export const insertContentControl = forwardTool(
  'word_insert_content_control',
  '[Content Controls] Wrap a text match in a content control. CheckBox type REPLACES anchor text with a checkbox.',
  {
    properties: {
      anchorText: { type: 'string', description: 'Text to search for and wrap' },
      type: { type: 'string', enum: ['RichText', 'PlainText', 'CheckBox'], description: 'Default: RichText' },
      title: { type: 'string' },
      tag: { type: 'string', description: 'Tag for programmatic identification' },
      color: { type: 'string', description: 'Border color' },
      occurrence: { type: 'number', description: '0=first, 1=second, etc. Default: 0' },
      matchCase: { type: 'boolean', description: 'Default: false' },
    },
  },
  'insertContentControl',
);

export const setContentControlText = forwardTool(
  'word_set_content_control_text',
  '[Content Controls] Set text in a content control identified by ID or tag. Does NOT work on CheckBox controls.',
  {
    properties: {
      id: { type: 'number', description: 'Content control ID' },
      tag: { type: 'string', description: 'Content control tag (alternative to ID)' },
      text: { type: 'string' },
    },
    required: ['text'],
  },
  'setContentControlText',
);

export const contentControlTools: ToolDefinition[] = [
  getContentControls,
  insertContentControl,
  setContentControlText,
];
