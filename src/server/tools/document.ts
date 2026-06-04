import type { ToolDefinition } from '../types';
import { ToolError } from '../types';
import { forwardTool, jsonResult } from './helpers';
import type { GetParagraphsResult } from '../../shared/protocol';

export const getText = forwardTool(
  'word_get_text',
  '[Document] Get full plain text of the active document.',
  { properties: {} },
  'getDocumentText',
);

export const getDocumentProperties = forwardTool(
  'word_get_document_properties',
  '[Document] Get all document metadata including title, author, path, changeTrackingMode, template, security, and timestamps.',
  { properties: {} },
  'getDocumentProperties',
);

export const setDocumentProperties = forwardTool(
  'word_set_document_properties',
  '[Document] Set document metadata (title, subject, author, keywords, comments, category, company, manager, format).',
  {
    properties: {
      title: { type: 'string' },
      subject: { type: 'string' },
      author: { type: 'string' },
      keywords: { type: 'string' },
      comments: { type: 'string' },
      category: { type: 'string' },
      company: { type: 'string' },
      manager: { type: 'string' },
      format: { type: 'string' },
    },
  },
  'setDocumentProperties',
);

export const save = forwardTool(
  'word_save',
  '[Document] Save the document to disk.',
  { properties: {} },
  'saveDocument',
);

export const clear = forwardTool(
  'word_clear',
  '[Document] Clear all document body content. Does not clear headers/footers or custom properties. In multi-section documents, section breaks are removed and the last section\'s layout (margins, orientation) is preserved.',
  { properties: {} },
  'clearDocument',
);

export const getWordCount = forwardTool(
  'word_get_word_count',
  '[Document] Get word, character, and paragraph counts.',
  { properties: {} },
  'getWordCount',
);

export const getStyles = forwardTool(
  'word_get_styles',
  '[Document] Get available document styles (returns up to 80 styles with name, type, and builtIn flag).',
  { properties: {} },
  'getStyles',
);

export const getCoauthors = forwardTool(
  'word_get_coauthors',
  '[Document] Get current co-authors and coauthoring status (Desktop only).',
  { properties: {} },
  'getCoauthors',
);

export const setChangeTracking = forwardTool(
  'word_set_change_tracking',
  '[Document] Set change tracking mode. Call with "TrackAll" BEFORE making edits to show them as tracked changes.',
  {
    properties: {
      mode: { type: 'string', enum: ['TrackAll', 'TrackMineOnly', 'Off'] },
    },
    required: ['mode'],
  },
  'setChangeTracking',
);

export const getDocumentOutline: ToolDefinition = {
  name: 'word_get_document_outline',
  description: '[Document] Get the document heading hierarchy as a structured outline tree. Returns headings with level, text, and paragraph index.',
  schema: {
    properties: {
      maxLevel: { type: 'number', description: 'Maximum heading level to include (1-9). Default: 3' },
    },
  },
  async handler(args, bridge) {
    const maxLevel = (args.maxLevel as number) ?? 3;
    if (typeof args.maxLevel === 'number' && (args.maxLevel < 1 || args.maxLevel > 9 || !Number.isInteger(args.maxLevel))) {
      throw new ToolError('maxLevel must be an integer between 1 and 9.');
    }
    const result = await bridge.send<GetParagraphsResult>('getParagraphs', {});
    const headings: Array<{ level: number; text: string; index: number }> = [];

    for (const para of result.paragraphs) {
      if (para.isTocEntry) continue;

      let level: number | null = null;
      const match = para.style?.match(/^Heading (\d)$/i);
      if (match) {
        level = parseInt(match[1]!, 10);
      } else if (para.outlineLevel >= 1 && para.outlineLevel <= 9) {
        level = para.outlineLevel;
      }

      if (level !== null && level <= maxLevel) {
        headings.push({ level, text: para.text, index: para.index });
      }
    }

    return jsonResult({ count: headings.length, outline: headings });
  },
};

export const documentTools: ToolDefinition[] = [
  getText,
  getDocumentProperties,
  setDocumentProperties,
  save,
  clear,
  getWordCount,
  getStyles,
  getCoauthors,
  setChangeTracking,
  getDocumentOutline,
];
