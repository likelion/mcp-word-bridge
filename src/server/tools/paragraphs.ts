import type { ToolDefinition } from '../types';
import { ToolError } from '../types';
import { forwardTool, jsonResult } from './helpers';
import { checkNonNegative, checkBounds, checkSpacingBounds, checkIndentBounds } from '../validation';
import type { GetParagraphsResult, OoxmlResult } from '../../shared/protocol';

export const getParagraphs = forwardTool(
  'word_get_paragraphs',
  '[Paragraphs] Get paragraphs with text, style, alignment, and metadata. Supports pagination via optional start/end index range (0-based).',
  {
    properties: {
      start: { type: 'number', description: 'First paragraph index to return (0-based, inclusive)' },
      end: { type: 'number', description: 'Last paragraph index (exclusive). Omit to get all from start.' },
    },
  },
  'getParagraphs',
);

export const getParagraphByIndex = forwardTool(
  'word_get_paragraph_by_index',
  '[Paragraphs] Get full details of a single paragraph including font, spacing, indentation, and outline level. Font properties return null when the paragraph has mixed formatting (e.g. partially bold).',
  {
    properties: {
      index: { type: 'number', description: 'Paragraph index (0-based)' },
    },
    required: ['index'],
  },
  'getParagraphByIndex',
);

export const insertParagraph = forwardTool(
  'word_insert_paragraph',
  '[Paragraphs] Append or prepend a styled paragraph to the document (Start or End only). For inserting at a specific position, use word_insert_paragraph_at_index.',
  {
    properties: {
      text: { type: 'string' },
      location: { type: 'string', enum: ['Start', 'End'], description: 'Default: End' },
      style: { type: 'string', description: 'Paragraph style (e.g. "Heading 1", "Normal"). Default: Normal' },
      alignment: { type: 'string', description: 'Left, Center, Right, or Justified' },
    },
    required: ['text'],
  },
  'insertParagraph',
);

export const insertParagraphAtIndex = forwardTool(
  'word_insert_paragraph_at_index',
  '[Paragraphs] Insert a new paragraph Before or After a specific paragraph index.',
  {
    properties: {
      index: { type: 'number', description: 'Reference paragraph index (0-based)' },
      text: { type: 'string', description: 'Text content for the new paragraph' },
      location: { type: 'string', enum: ['Before', 'After'], description: 'Default: After' },
      style: { type: 'string', description: 'Paragraph style' },
      alignment: { type: 'string', description: 'Left, Center, Right, or Justified' },
    },
    required: ['index', 'text'],
  },
  'insertParagraphAtIndex',
);

export const deleteParagraph = forwardTool(
  'word_delete_paragraph',
  '[Paragraphs] Delete a paragraph by its 0-based index. For table cells with multiple paragraphs, extra paragraphs can be removed (the last paragraph in a cell cannot be deleted).',
  {
    properties: {
      index: { type: 'number', description: 'Paragraph index (0-based)' },
    },
    required: ['index'],
  },
  'deleteParagraph',
);

export const replaceParagraphText = forwardTool(
  'word_replace_paragraph_text',
  '[Paragraphs] Replace the entire text of a paragraph by index. Preserves style/formatting. Preferred over word_search_and_replace in collaborative editing.',
  {
    properties: {
      index: { type: 'number', description: 'Paragraph index (0-based)' },
      text: { type: 'string', description: 'New text content' },
    },
    required: ['index', 'text'],
  },
  'replaceParagraphText',
);

export const setParagraphStyle = forwardTool(
  'word_set_paragraph_style',
  '[Paragraphs] Change the style or alignment of a paragraph by index.',
  {
    properties: {
      index: { type: 'number', description: 'Paragraph index (0-based)' },
      style: { type: 'string' },
      alignment: { type: 'string', description: 'Left, Center, Right, or Justified' },
    },
    required: ['index'],
  },
  'setParagraphStyle',
);

export const setParagraphSpacing: ToolDefinition = {
  name: 'word_set_paragraph_spacing',
  description: '[Paragraphs] Set line spacing, before/after spacing, and indentation on a paragraph by index.',
  schema: {
    properties: {
      index: { type: 'number', description: 'Paragraph index (0-based)' },
      lineSpacing: { type: 'number', description: 'Line spacing in points' },
      spaceBefore: { type: 'number', description: 'Space before in points' },
      spaceAfter: { type: 'number', description: 'Space after in points' },
      firstLineIndent: { type: 'number', description: 'First line indent in points' },
      leftIndent: { type: 'number', description: 'Left indent in points' },
      rightIndent: { type: 'number', description: 'Right indent in points' },
    },
    required: ['index'],
  },
  async handler(args, bridge) {
    const index = args.index as number;
    checkNonNegative(index, 'index');

    // Require at least one spacing/indent property
    const hasProperty =
      args.lineSpacing !== undefined ||
      args.spaceBefore !== undefined ||
      args.spaceAfter !== undefined ||
      args.firstLineIndent !== undefined ||
      args.leftIndent !== undefined ||
      args.rightIndent !== undefined;
    if (!hasProperty) {
      throw new ToolError(
        'At least one spacing or indent property must be provided (lineSpacing, spaceBefore, spaceAfter, firstLineIndent, leftIndent, rightIndent).',
      );
    }

    // Validate bounds on spacing and indent values
    const spacingFields: Array<[string, unknown]> = [
      ['lineSpacing', args.lineSpacing],
      ['spaceBefore', args.spaceBefore],
      ['spaceAfter', args.spaceAfter],
    ];
    for (const [name, value] of spacingFields) {
      if (value !== undefined && typeof value === 'number') {
        checkSpacingBounds(value, name);
      }
    }

    const indentFields: Array<[string, unknown]> = [
      ['firstLineIndent', args.firstLineIndent],
      ['leftIndent', args.leftIndent],
      ['rightIndent', args.rightIndent],
    ];
    for (const [name, value] of indentFields) {
      if (value !== undefined && typeof value === 'number') {
        checkIndentBounds(value, name);
      }
    }

    const result = await bridge.send('setParagraphSpacing', args);
    return jsonResult(result);
  },
};

export const moveParagraph: ToolDefinition = {
  name: 'word_move_paragraph',
  description: '[Paragraphs] Move paragraph(s) to another position. Preserves all rich content including footnotes, hyperlinks, formatting, images, and comments.',
  schema: {
    properties: {
      fromIndex: { type: 'number', description: 'Source paragraph index (0-based)' },
      toIndex: { type: 'number', description: 'Destination paragraph index (0-based)' },
      location: { type: 'string', enum: ['Before', 'After'], description: 'Default: After' },
      count: { type: 'number', description: 'Number of consecutive paragraphs to move (default: 1)' },
    },
    required: ['fromIndex', 'toIndex'],
  },
  async handler(args, bridge) {
    const fromIndex = args.fromIndex as number;
    const toIndex = args.toIndex as number;
    const count = (args.count as number) ?? 1;
    const location = (args.location as string) ?? 'After';

    checkNonNegative(fromIndex, 'fromIndex');
    checkNonNegative(toIndex, 'toIndex');
    if (!Number.isInteger(count)) throw new ToolError('count must be an integer.');
    if (count < 1) throw new ToolError('count must be at least 1');
    if (fromIndex === toIndex && count === 1) throw new ToolError('fromIndex and toIndex must be different');
    if (toIndex >= fromIndex && toIndex < fromIndex + count) {
      throw new ToolError(`toIndex (${toIndex}) is inside the source range [${fromIndex}, ${fromIndex + count - 1}]. Move to a position outside the range.`);
    }

    const paraCount = await bridge.send<GetParagraphsResult>('getParagraphs', {});
    const total = paraCount.total;
    checkBounds(fromIndex, total, 'fromIndex');
    if (fromIndex + count - 1 >= total) throw new ToolError(`fromIndex + count (${fromIndex + count}) exceeds paragraph count (${total}).`);
    checkBounds(toIndex, total, 'toIndex');

    // Reject moves involving table-internal paragraphs
    for (const para of paraCount.paragraphs) {
      if (para.index >= fromIndex && para.index < fromIndex + count && para.inTable) {
        throw new ToolError(`Paragraph ${para.index} is inside a table cell. Use table-specific tools to modify table content.`);
      }
    }

    // Reject moves targeting table-internal paragraphs (would corrupt table structure)
    const destPara = paraCount.paragraphs.find(p => p.index === toIndex);
    if (destPara?.inTable) {
      throw new ToolError(
        `Destination paragraph ${toIndex} is inside a table cell. ` +
        `Moving content here would corrupt table structure. ` +
        `Use table-specific tools or target a paragraph outside the table.`,
      );
    }

    // Reject move of the mandatory trailing paragraph (Word auto-recreates it, causing duplication)
    const lastIdx = total - 1;
    if (fromIndex + count - 1 === lastIdx) {
      const lastPara = paraCount.paragraphs.find(p => p.index === lastIdx);
      if (lastPara && lastPara.text === '') {
        throw new ToolError(
          `Cannot move the last paragraph (index ${lastIdx}) — Word requires at least one paragraph and will auto-create a replacement, resulting in duplication. Use word_copy_paragraph instead, or exclude the trailing paragraph from the range.`,
        );
      }
    }

    // Detect no-op: moving paragraphs immediately after themselves
    const adjustedTo = fromIndex < toIndex ? toIndex - count : toIndex;
    if (toIndex === fromIndex + count && location === 'After') {
      return jsonResult({ success: true, warning: 'No move performed — destination is equivalent to source position.', moved: null });
    }
    if (adjustedTo === fromIndex && location === 'After') {
      return jsonResult({ success: true, warning: 'No move performed — destination is equivalent to source position.', moved: null });
    }

    // Capture full content via OOXML
    const ooxmlResult = await bridge.send<OoxmlResult>('getParaOoxml', { index: fromIndex, count });
    const savedOoxml = ooxmlResult.ooxml;

    // Delete source paragraphs (last to first)
    for (let i = count - 1; i >= 0; i--) {
      await bridge.send('deleteParagraph', { index: fromIndex + i });
    }

    // Insert at destination with restore-on-failure
    try {
      await bridge.send('insertOoxmlAtIndex', { ooxml: savedOoxml, index: adjustedTo, location });
    } catch (insertErr) {
      try {
        await bridge.send('insertOoxmlAtIndex', { ooxml: savedOoxml, index: Math.min(fromIndex, adjustedTo), location: 'Before' });
      } catch {
        throw new ToolError(`Move failed AND restore failed: ${(insertErr as Error).message}. Use Ctrl+Z to recover.`);
      }
      throw new ToolError(`Move failed (content restored to original position): ${(insertErr as Error).message}`);
    }

    return jsonResult({ success: true, moved: { from: fromIndex, count, to: adjustedTo, toIndexRequested: toIndex, location } });
  },
};

export const copyParagraph: ToolDefinition = {
  name: 'word_copy_paragraph',
  description: '[Paragraphs] Copy paragraph(s) to another position. Preserves all rich content. Source remains unchanged.',
  schema: {
    properties: {
      fromIndex: { type: 'number', description: 'Source paragraph index (0-based)' },
      toIndex: { type: 'number', description: 'Destination paragraph index (0-based)' },
      location: { type: 'string', enum: ['Before', 'After'], description: 'Default: After' },
      count: { type: 'number', description: 'Number of consecutive paragraphs to copy (default: 1)' },
    },
    required: ['fromIndex', 'toIndex'],
  },
  async handler(args, bridge) {
    const fromIndex = args.fromIndex as number;
    const toIndex = args.toIndex as number;
    const count = (args.count as number) ?? 1;
    const location = (args.location as string) ?? 'After';

    checkNonNegative(fromIndex, 'fromIndex');
    checkNonNegative(toIndex, 'toIndex');
    if (!Number.isInteger(count)) throw new ToolError('count must be an integer.');
    if (count < 1) throw new ToolError('count must be at least 1');

    const paraCount = await bridge.send<GetParagraphsResult>('getParagraphs', {});
    const total = paraCount.total;
    checkBounds(fromIndex, total, 'fromIndex');
    if (fromIndex + count - 1 >= total) throw new ToolError(`fromIndex + count (${fromIndex + count}) exceeds paragraph count (${total}).`);
    checkBounds(toIndex, total, 'toIndex');

    // Reject copies from table-internal paragraphs
    for (const para of paraCount.paragraphs) {
      if (para.index >= fromIndex && para.index < fromIndex + count && para.inTable) {
        throw new ToolError(`Paragraph ${para.index} is inside a table cell. Use table-specific tools to modify table content.`);
      }
    }

    // Reject copies targeting table-internal paragraphs (would corrupt table structure)
    const destPara = paraCount.paragraphs.find(p => p.index === toIndex);
    if (destPara?.inTable) {
      throw new ToolError(
        `Destination paragraph ${toIndex} is inside a table cell. ` +
        `Copying content here would corrupt table structure. ` +
        `Use table-specific tools or target a paragraph outside the table.`,
      );
    }

    const ooxmlResult = await bridge.send<OoxmlResult>('getParaOoxml', { index: fromIndex, count });

    // When copying to the same position, offset toIndex past the source range
    // to avoid OOXML insertion inside the source paragraph
    const effectiveToIndex = (toIndex >= fromIndex && toIndex < fromIndex + count)
      ? fromIndex + count - 1
      : toIndex;
    const effectiveLocation = (toIndex >= fromIndex && toIndex < fromIndex + count)
      ? 'After'
      : location;

    await bridge.send('insertOoxmlAtIndex', { ooxml: ooxmlResult.ooxml, index: effectiveToIndex, location: effectiveLocation });

    return jsonResult({ success: true, copied: { from: fromIndex, count, to: toIndex, location } });
  },
};

export const paragraphTools: ToolDefinition[] = [
  getParagraphs,
  getParagraphByIndex,
  insertParagraph,
  insertParagraphAtIndex,
  deleteParagraph,
  replaceParagraphText,
  moveParagraph,
  copyParagraph,
  setParagraphStyle,
  setParagraphSpacing,
];
