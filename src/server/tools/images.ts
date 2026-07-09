import type { ToolDefinition } from '../types';
import { forwardTool } from './helpers';

export const insertImage = forwardTool(
  'word_insert_image',
  '[Images] Insert an inline image from base64-encoded data (PNG, JPEG, or GIF).',
  {
    properties: {
      base64: { type: 'string', description: 'Base64-encoded image data' },
      location: { type: 'string', enum: ['Start', 'End'] },
      width: { type: 'number', description: 'Width in points' },
      height: { type: 'number', description: 'Height in points' },
      altText: { type: 'string', description: 'Alt text for accessibility' },
      caption: { type: 'string', description: 'Optional caption text. Adds an auto-numbered "Figure N: <caption>" caption paragraph below the image.' },
    },
    required: ['base64'],
  },
  'insertImage',
);

export const getImages = forwardTool(
  'word_get_images',
  '[Images] List all inline images with index, dimensions, alt text, and hyperlinks.',
  { properties: {} },
  'getImages',
);

export const deleteImage = forwardTool(
  'word_delete_image',
  '[Images] Delete an inline image (figure) by its 0-based index, including its caption paragraph if present.',
  {
    properties: {
      index: { type: 'number', description: 'Image index (0-based)' },
      deleteCaption: { type: 'boolean', description: 'Also delete the "Caption"-styled paragraph directly below the image. Default: true.' },
    },
    required: ['index'],
  },
  'deleteImage',
);

export const imageTools: ToolDefinition[] = [
  insertImage,
  getImages,
  deleteImage,
];
