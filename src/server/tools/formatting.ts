import type { ToolDefinition } from '../types';
import { ToolError } from '../types';
import { forwardTool } from './helpers';
import { checkNonEmpty, checkHexColor } from '../validation';
import { HIGHLIGHT_COLORS } from '../../shared/constants';

function validateFormatText(args: Record<string, unknown>): void {
  checkNonEmpty(args.text, 'text');

  const hasFormatting =
    args.bold !== undefined ||
    args.italic !== undefined ||
    args.underline !== undefined ||
    args.strikeThrough !== undefined ||
    args.color !== undefined ||
    args.highlightColor !== undefined ||
    args.size !== undefined ||
    args.name !== undefined;

  if (!hasFormatting) {
    throw new ToolError(
      'At least one formatting property must be specified (bold, italic, underline, strikeThrough, color, highlightColor, size, or name).',
    );
  }

  if (args.size !== undefined) {
    const size = args.size as number;
    if (size <= 0) throw new ToolError('size must be positive (minimum 1 point).');
    if (size > 1638) throw new ToolError('size must not exceed 1638 points (Word maximum).');
    if (!Number.isFinite(size)) throw new ToolError('size must be a finite number.');
  }

  if (args.color !== undefined) {
    checkHexColor(args.color as string, 'color');
  }

  if (args.highlightColor !== undefined) {
    const hc = args.highlightColor as string;
    const isNamed = HIGHLIGHT_COLORS.some(c => c.toLowerCase() === hc.toLowerCase());
    if (!isNamed) {
      throw new ToolError(
        `Invalid highlightColor: "${hc}". Valid values: ${HIGHLIGHT_COLORS.join(', ')}.`,
      );
    }
  }
}

export const formatText = forwardTool(
  'word_format_text',
  '[Formatting] Apply formatting (bold, italic, color, size, font) to a text match. Color must be hex (#FF0000). Size: 1-1638pt.',
  {
    properties: {
      text: { type: 'string', description: 'Text to search for and format' },
      occurrence: { type: 'number', description: '0=first, 1=second, etc. Default: 0' },
      bold: { type: 'boolean' },
      italic: { type: 'boolean' },
      underline: { type: 'boolean' },
      strikeThrough: { type: 'boolean' },
      color: { type: 'string', description: 'Hex color e.g. #FF0000' },
      highlightColor: { type: 'string', description: 'Highlight color name (Yellow, Green, Cyan, Magenta, Blue, Red, DarkBlue, DarkCyan, DarkGreen, DarkMagenta, DarkRed, DarkYellow, Gray25, Gray50, Black, White, NoHighlight)' },
      size: { type: 'number', description: 'Font size in points (1-1638)' },
      name: { type: 'string', description: 'Font name' },
      matchCase: { type: 'boolean', description: 'Default: false' },
    },
    required: ['text'],
  },
  'formatRange',
  validateFormatText,
);

export const clearFormatting = forwardTool(
  'word_clear_formatting',
  '[Formatting] Clear direct formatting from a text match, reverting to paragraph style defaults.',
  {
    properties: {
      text: { type: 'string', description: 'Text to search for' },
      occurrence: { type: 'number', description: '0=first, 1=second, etc. Default: 0' },
      matchCase: { type: 'boolean', description: 'Default: false' },
    },
    required: ['text'],
  },
  'clearFormatting',
  (args) => {
    checkNonEmpty(args.text, 'text');
  },
);

export const getFontInfo = forwardTool(
  'word_get_font_info',
  '[Formatting] Inspect font properties (name, size, bold, italic, color) of a text match.',
  {
    properties: {
      text: { type: 'string', description: 'Text to search for' },
      occurrence: { type: 'number', description: '0=first, 1=second, etc. Default: 0' },
      matchCase: { type: 'boolean', description: 'Default: false' },
    },
    required: ['text'],
  },
  'getFontInfo',
  (args) => {
    checkNonEmpty(args.text, 'text');
  },
);

export const formattingTools: ToolDefinition[] = [
  formatText,
  clearFormatting,
  getFontInfo,
];
