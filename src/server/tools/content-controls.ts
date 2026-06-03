import type { ToolDefinition } from '../types';
import { ToolError } from '../types';
import { forwardTool, jsonResult } from './helpers';

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

export const setContentControlText: ToolDefinition = {
  name: 'word_set_content_control_text',
  description: '[Content Controls] Set text in a content control identified by ID or tag. Does NOT work on CheckBox controls.',
  schema: {
    properties: {
      id: { type: 'number', description: 'Content control ID' },
      tag: { type: 'string', description: 'Content control tag (alternative to ID)' },
      text: { type: 'string' },
    },
    required: ['text'],
  },
  async handler(args, bridge) {
    const tag = args.tag as string | undefined;
    const id = args.id as number | undefined;

    if (!tag && id === undefined) {
      throw new ToolError('Provide "id" or "tag" to identify the content control. Use word_get_content_controls to list available controls.');
    }

    // BUG-03: Check for duplicate tags before forwarding
    if (tag) {
      const ccResult = await bridge.send<{ count: number; controls: Array<{ id: number; tag: string }> }>('getContentControls', {});
      const matches = ccResult.controls.filter(c => c.tag === tag);
      if (matches.length > 1) {
        throw new ToolError(
          `Multiple content controls (${matches.length}) share tag "${tag}". ` +
          `Use "id" instead to target a specific control. ` +
          `Matching IDs: ${matches.map(m => m.id).join(', ')}.`,
        );
      }
    }

    const result = await bridge.send('setContentControlText', args);
    return jsonResult(result);
  },
};

export const contentControlTools: ToolDefinition[] = [
  getContentControls,
  insertContentControl,
  setContentControlText,
];
