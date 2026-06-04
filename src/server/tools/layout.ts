import type { ToolDefinition } from '../types';
import { ToolError } from '../types';
import { forwardTool } from './helpers';

export const getPageLayout = forwardTool(
  'word_get_page_layout',
  '[Layout] Get page layout (margins, orientation, paper size) for a section.',
  {
    properties: {
      sectionIndex: { type: 'number', description: 'Section index (0-based). Default: 0' },
    },
  },
  'getPageLayout',
);

export const setPageLayout = forwardTool(
  'word_set_page_layout',
  '[Layout] Set page layout (margins, orientation, paper size) for a section.',
  {
    properties: {
      sectionIndex: { type: 'number', description: 'Section index (0-based). Default: 0' },
      orientation: { type: 'string', description: 'Portrait or Landscape' },
      topMargin: { type: 'number', description: 'Top margin in points' },
      bottomMargin: { type: 'number', description: 'Bottom margin in points' },
      leftMargin: { type: 'number', description: 'Left margin in points' },
      rightMargin: { type: 'number', description: 'Right margin in points' },
      paperSize: { type: 'string', description: 'Letter, A4, etc.' },
    },
  },
  'setPageLayout',
  (args) => {
    const hasProperty =
      args.orientation !== undefined ||
      args.topMargin !== undefined ||
      args.bottomMargin !== undefined ||
      args.leftMargin !== undefined ||
      args.rightMargin !== undefined ||
      args.paperSize !== undefined;
    if (!hasProperty) {
      throw new ToolError(
        'At least one layout property must be provided (orientation, topMargin, bottomMargin, leftMargin, rightMargin, paperSize).',
      );
    }
  },
);

export const getSections = forwardTool(
  'word_get_sections',
  '[Layout] List all sections with their page setup.',
  { properties: {} },
  'getSections',
);

export const insertPageBreak = forwardTool(
  'word_insert_page_break',
  '[Layout] Insert a page break after a paragraph. Omit paragraphIndex to insert at end.',
  {
    properties: {
      paragraphIndex: { type: 'number', description: 'Paragraph index (0-based). Omit for end.' },
    },
  },
  'insertPageBreak',
);

export const insertSectionBreak = forwardTool(
  'word_insert_section_break',
  '[Layout] Insert a section break after a paragraph.',
  {
    properties: {
      paragraphIndex: { type: 'number', description: 'Paragraph index (0-based). Omit for end.' },
      breakType: { type: 'string', enum: ['SectionNext', 'SectionContinuous', 'SectionEven', 'SectionOdd'], description: 'Default: SectionNext' },
    },
  },
  'insertSectionBreak',
);

export const getPageInfo = forwardTool(
  'word_get_page_info',
  '[Layout] Get page count and per-page paragraph ranges (Desktop only).',
  { properties: {} },
  'getPageInfo',
);

export const layoutTools: ToolDefinition[] = [
  getPageLayout,
  setPageLayout,
  getSections,
  insertPageBreak,
  insertSectionBreak,
  getPageInfo,
];
