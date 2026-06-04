import type { ToolDefinition } from '../types';
import { forwardTool, jsonResult } from './helpers';
import { checkNonEmpty, checkPropertyKeyLength } from '../validation';

export const getCustomProperties = forwardTool(
  'word_get_custom_properties',
  '[Properties] Get all custom document properties (key-value pairs with types).',
  { properties: {} },
  'getCustomProperties',
);

export const setCustomProperty: ToolDefinition = {
  name: 'word_set_custom_property',
  description: '[Properties] Set a custom document property. Creates or updates the key-value pair.',
  schema: {
    properties: {
      key: { type: 'string' },
      value: { type: 'string' },
    },
    required: ['key', 'value'],
  },
  async handler(args, bridge) {
    const key = args.key as string;
    checkNonEmpty(key, 'key');
    // Validate key length to prevent silent truncation
    checkPropertyKeyLength(key);

    const result = await bridge.send('setCustomProperty', args);
    return jsonResult(result);
  },
};

export const deleteCustomProperty = forwardTool(
  'word_delete_custom_property',
  '[Properties] Delete a custom document property by key.',
  {
    properties: {
      key: { type: 'string' },
    },
    required: ['key'],
  },
  'deleteCustomProperty',
  (args) => {
    checkNonEmpty(args.key, 'key');
  },
);

export const propertyTools: ToolDefinition[] = [
  getCustomProperties,
  setCustomProperty,
  deleteCustomProperty,
];
