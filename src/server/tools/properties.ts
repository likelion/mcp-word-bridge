import type { ToolDefinition } from '../types';
import { forwardTool } from './helpers';

export const getCustomProperties = forwardTool(
  'word_get_custom_properties',
  '[Properties] Get all custom document properties (key-value pairs with types).',
  { properties: {} },
  'getCustomProperties',
);

export const setCustomProperty = forwardTool(
  'word_set_custom_property',
  '[Properties] Set a custom document property. Creates or updates the key-value pair.',
  {
    properties: {
      key: { type: 'string' },
      value: { type: 'string' },
    },
    required: ['key', 'value'],
  },
  'setCustomProperty',
);

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
);

export const propertyTools: ToolDefinition[] = [
  getCustomProperties,
  setCustomProperty,
  deleteCustomProperty,
];
