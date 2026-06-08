import type { CommandHandler } from './index';
import { sanitizeText, checkIndex, checkAlignment, checkStyleExists, normalizeAlignmentOutput } from './document';

/** Detect if a paragraph is a TOC entry based on style and text pattern */
function isTocEntry(style: string, text: string): boolean {
  return !!(style && style.startsWith('TOC')) || (style === '' && /\t\d+$/.test(text));
}

/** Infer style from outlineLevel when style is empty (e.g. after TOC generation) */
function resolveStyle(style: string, outlineLevel: number): string {
  if (style === '' && outlineLevel >= 1 && outlineLevel <= 9) {
    return `Heading ${outlineLevel}`;
  }
  return style;
}

export const paragraphCommands: Record<string, CommandHandler> = {
  async getParagraphs(ctx, p) {
    if (p.start !== undefined) { if (typeof p.start !== 'number' || !Number.isInteger(p.start) || p.start < 0) throw new Error('start index must be a non-negative integer'); }
    if (p.end !== undefined) { if (typeof p.end !== 'number' || !Number.isInteger(p.end) || p.end < 0) throw new Error('end index must be a non-negative integer'); }
    const paragraphs = ctx.document.body.paragraphs;
    let hasTableInfo = true;
    try {
      paragraphs.load('text,style,alignment,firstLineIndent,leftIndent,lineSpacing,isListItem,outlineLevel,parentTableCellOrNullObject');
      await ctx.sync();
    } catch {
      hasTableInfo = false;
      paragraphs.load('text,style,alignment,firstLineIndent,leftIndent,lineSpacing,isListItem,outlineLevel');
      await ctx.sync();
    }
    const start = p.start ?? 0;
    const end = p.end ?? paragraphs.items.length;
    if (start > end) throw new Error(`start (${start}) must be less than or equal to end (${end})`);
    const items: any[] = [];
    for (let i = start; i < Math.min(end, paragraphs.items.length); i++) {
      const para = paragraphs.items[i];
      let inTable = false;
      if (hasTableInfo) {
        try { inTable = para.parentTableCellOrNullObject && !para.parentTableCellOrNullObject.isNullObject; } catch { inTable = false; }
      }
      const isToc = isTocEntry(para.style, para.text);
      const resolvedStyle = resolveStyle(para.style, para.outlineLevel);
      items.push({ index: i, text: sanitizeText(para.text), style: resolvedStyle, alignment: normalizeAlignmentOutput(para.alignment), isListItem: para.isListItem, inTable, isTocEntry: isToc, outlineLevel: para.outlineLevel });
    }
    const result: any = { total: paragraphs.items.length, count: items.length, paragraphs: items };
    if (p.start !== undefined && p.start >= paragraphs.items.length) {
      result.warning = `start index (${p.start}) is beyond the last paragraph. Document has ${paragraphs.items.length} paragraphs (valid indices: 0-${paragraphs.items.length - 1}).`;
    }
    return result;
  },

  async getParagraphByIndex(ctx, p) {
    checkIndex(p.index, 'index');
    const paragraphs = ctx.document.body.paragraphs;
    paragraphs.load('text');
    await ctx.sync();
    if (p.index >= paragraphs.items.length) throw new Error(`Paragraph index out of range. Valid indices: 0-${paragraphs.items.length - 1} (document has ${paragraphs.items.length} paragraphs).`);
    const para = paragraphs.items[p.index];
    para.load('text,style,alignment,firstLineIndent,leftIndent,rightIndent,lineSpacing,spaceBefore,spaceAfter,outlineLevel,isListItem,parentTableCellOrNullObject');
    para.font.load('name,size,bold,italic,color,underline');
    await ctx.sync();
    let inTable = false;
    try { inTable = para.parentTableCellOrNullObject && !para.parentTableCellOrNullObject.isNullObject; } catch { inTable = false; }
    return { text: sanitizeText(para.text), style: para.style, alignment: normalizeAlignmentOutput(para.alignment), firstLineIndent: para.firstLineIndent, leftIndent: para.leftIndent, rightIndent: para.rightIndent, lineSpacing: para.lineSpacing, spaceBefore: para.spaceBefore, spaceAfter: para.spaceAfter, outlineLevel: para.outlineLevel, isListItem: para.isListItem, inTable, font: { name: para.font.name || null, size: para.font.size ?? null, bold: para.font.bold, italic: para.font.italic, color: para.font.color || null, underline: para.font.underline || null } };
  },

  async insertParagraph(ctx, p) {
    if (p.location && p.location !== 'Start' && p.location !== 'End') throw new Error(`Invalid location: "${p.location}". Valid values: Start, End`);
    const alignment = checkAlignment(p.alignment);
    const styleName = p.style || 'Normal';
    if (p.style) await checkStyleExists(ctx, p.style);
    const para = ctx.document.body.insertParagraph(p.text, p.location || 'End');
    para.style = styleName;
    await ctx.sync();
    if (alignment) { para.alignment = alignment; await ctx.sync(); }
    if (!styleName.startsWith('Heading')) {
      let defaultSize = 12;
      try {
        const s2 = ctx.document.getStyles().getByNameOrNullObject(styleName);
        s2.load('font/size');
        await ctx.sync();
        if (!s2.isNullObject && s2.font.size) defaultSize = s2.font.size;
      } catch { /* fallback 12pt */ }
      para.font.size = defaultSize;
      await ctx.sync();
    }
    para.getRange('End').select();
    await ctx.sync();
    const result: any = { success: true };
    if (!p.text || p.text.trim() === '') result.warning = 'Empty paragraph inserted.';
    return result;
  },

  async insertParagraphAtIndex(ctx, p) {
    checkIndex(p.index, 'index');
    const alignment = checkAlignment(p.alignment);
    const paragraphs = ctx.document.body.paragraphs;
    paragraphs.load('text,parentTableCellOrNullObject');
    await ctx.sync();
    if (p.index >= paragraphs.items.length) throw new Error(`Paragraph index out of range. Valid indices: 0-${paragraphs.items.length - 1} (document has ${paragraphs.items.length} paragraphs).`);
    let inTable = false;
    try { inTable = paragraphs.items[p.index].parentTableCellOrNullObject && !paragraphs.items[p.index].parentTableCellOrNullObject.isNullObject; } catch { /* ignore */ }
    const styleName = p.style || 'Normal';
    if (p.style) await checkStyleExists(ctx, p.style);
    const ref = paragraphs.items[p.index];
    const location = p.location === 'Before' ? 'Before' : 'After';

    let newPara: any;
    try {
      newPara = ref.insertParagraph(p.text, location);
      newPara.style = styleName;
      await ctx.sync();
    } catch (e: any) {
      if (!e.message?.includes('GeneralException')) throw e;
      // Fallback: use range-based insertion (insertText with paragraph mark)
      try {
        const rangeLocation = location === 'After' ? 'End' : 'Start';
        const range = ref.getRange(rangeLocation);
        const inserted = range.insertText('\n' + p.text, location === 'After' ? 'After' : 'Before');
        void inserted; // insertText returns a Range proxy; we only need the sync side-effect
        await ctx.sync();
        // Re-load paragraphs to find and style the new one
        const freshParas = ctx.document.body.paragraphs;
        freshParas.load('text,style');
        await ctx.sync();
        const newIdx = location === 'After' ? p.index + 1 : p.index;
        if (newIdx < freshParas.items.length) {
          freshParas.items[newIdx].style = styleName;
          await ctx.sync();
          if (alignment) { freshParas.items[newIdx].alignment = alignment; await ctx.sync(); }
          freshParas.items[newIdx].getRange('End').select();
          await ctx.sync();
        }
        const result: any = { success: true };
        if (inTable) result.warning = 'Paragraph inserted inside a table cell. This creates a multi-paragraph cell which may not be intended.';
        return result;
      } catch (fallbackErr: any) {
        // If fallback also fails, throw a descriptive error
        throw new Error(`Cannot insert paragraph at index ${p.index} (${location}). Primary insert failed with GeneralException; fallback also failed: ${fallbackErr.message}. The document structure at this position may not support paragraph insertion.`);
      }
    }

    if (alignment) { newPara.alignment = alignment; await ctx.sync(); }
    newPara.getRange('End').select();
    await ctx.sync();
    const result: any = { success: true };
    if (inTable) result.warning = 'Paragraph inserted inside a table cell. This creates a multi-paragraph cell which may not be intended.';
    return result;
  },

  async deleteParagraph(ctx, p) {
    checkIndex(p.index, 'index');
    const paragraphs = ctx.document.body.paragraphs;
    paragraphs.load('text,isListItem,parentTableCellOrNullObject');
    await ctx.sync();
    if (p.index >= paragraphs.items.length) throw new Error(`Paragraph index out of range. Valid indices: 0-${paragraphs.items.length - 1} (document has ${paragraphs.items.length} paragraphs).`);
    // Guard: allow deletion of extra paragraphs in table cells, block last one
    let inTable = false;
    try { inTable = paragraphs.items[p.index].parentTableCellOrNullObject && !paragraphs.items[p.index].parentTableCellOrNullObject.isNullObject; } catch { /* ignore */ }
    if (inTable) {
      // Check if this is the only paragraph in the cell — if so, block deletion
      try {
        const cell = paragraphs.items[p.index].parentTableCellOrNullObject;
        const cellParas = cell.body.paragraphs;
        cellParas.load('text');
        await ctx.sync();
        if (cellParas.items.length <= 1) {
          throw new Error('Cannot delete the only paragraph in a table cell. Use word_set_table_cell or word_replace_paragraph_text to modify its content.');
        }
      } catch (e: any) {
        // If the check itself fails (older API), fall back to blocking
        if (e.message?.includes('Cannot delete the only paragraph')) throw e;
        throw new Error('Cannot delete a paragraph inside a table cell. Use word_set_table_cell or word_replace_paragraph_text to modify table content.');
      }
    }
    const countBefore = paragraphs.items.length;
    const para = paragraphs.items[p.index];
    const isListItem = para.isListItem;

    // Strategy 1: Try direct paragraph.delete()
    let deleted = false;
    try {
      para.delete();
      await ctx.sync();
      deleted = true;
    } catch (e: any) {
      // If direct delete fails, try fallback strategies
      if (!e.message?.includes('GeneralException')) throw e;
    }

    // Strategy 2: If list item, clear list formatting then retry delete
    if (!deleted && isListItem) {
      try {
        const freshParas = ctx.document.body.paragraphs;
        freshParas.load('text,isListItem');
        await ctx.sync();
        const freshPara = freshParas.items[p.index];
        // Detach from list by setting style to Normal (removes list formatting)
        freshPara.style = 'Normal';
        await ctx.sync();
        freshPara.delete();
        await ctx.sync();
        deleted = true;
      } catch (e: any) {
        if (!e.message?.includes('GeneralException')) throw e;
      }
    }

    // Strategy 3: Use range.delete() which handles more edge cases
    if (!deleted) {
      try {
        const freshParas = ctx.document.body.paragraphs;
        freshParas.load('text');
        await ctx.sync();
        const freshPara = freshParas.items[p.index];
        const range = freshPara.getRange('Whole');
        range.delete();
        await ctx.sync();
        deleted = true;
      } catch (e: any) {
        if (e.message?.includes('GeneralException')) {
          throw new Error(`Cannot delete paragraph ${p.index}. Word refused the operation — the paragraph may be a TOC field entry, inside a protected region, or part of a list structure that cannot be broken. Try word_replace_paragraph_text to clear its content instead.`);
        }
        throw e;
      }
    }

    const parasAfter = ctx.document.body.paragraphs;
    parasAfter.load('text');
    await ctx.sync();
    if (parasAfter.items.length > 0) {
      const cursorIdx = Math.min(p.index, parasAfter.items.length - 1);
      parasAfter.items[cursorIdx].getRange('Start').select();
      await ctx.sync();
    }
    const result: any = { success: true };
    if (parasAfter.items.length >= countBefore) result.warning = 'Paragraph was cleared but not removed (Word requires at least one paragraph).';
    return result;
  },

  async replaceParagraphText(ctx, p) {
    checkIndex(p.index, 'index');
    const paragraphs = ctx.document.body.paragraphs;
    paragraphs.load('text');
    await ctx.sync();
    if (p.index >= paragraphs.items.length) throw new Error(`Paragraph index out of range. Valid indices: 0-${paragraphs.items.length - 1} (document has ${paragraphs.items.length} paragraphs).`);
    const inserted = paragraphs.items[p.index].insertText(p.text, 'Replace');
    inserted.getRange('End').select();
    try { await ctx.sync(); } catch (e: any) {
      if (e.message?.includes('GeneralException')) throw new Error(`Cannot replace text of paragraph ${p.index}. It may be a TOC field entry, inside a protected region, or part of an uneditable structure.`);
      throw e;
    }
    return { success: true };
  },

  async setParagraphStyle(ctx, p) {
    checkIndex(p.index, 'index');
    const alignment = checkAlignment(p.alignment);
    if (!p.style && !p.alignment) throw new Error('At least one of "style" or "alignment" must be provided.');
    const paragraphs = ctx.document.body.paragraphs;
    paragraphs.load('text,style,outlineLevel');
    await ctx.sync();
    if (p.index >= paragraphs.items.length) throw new Error(`Paragraph index out of range. Valid indices: 0-${paragraphs.items.length - 1} (document has ${paragraphs.items.length} paragraphs).`);
    const para = paragraphs.items[p.index];
    // Guard: reject style changes on TOC entry paragraphs
    if (p.style && isTocEntry(para.style, para.text)) {
      throw new Error('Cannot change the style of a TOC entry paragraph. Modify the source headings and update the TOC instead.');
    }
    if (p.style) {
      await checkStyleExists(ctx, p.style);
      para.style = p.style;
      await ctx.sync();
    }
    if (alignment) { para.alignment = alignment; await ctx.sync(); }
    para.getRange('End').select();
    await ctx.sync();
    return { success: true };
  },

  async setParagraphSpacing(ctx, p) {
    checkIndex(p.index, 'index');
    if (p.lineSpacing !== undefined && p.lineSpacing <= 0) throw new Error('lineSpacing must be positive');
    if (p.lineSpacing !== undefined && p.lineSpacing < 1) throw new Error(`lineSpacing value ${p.lineSpacing} is below the minimum (1 point). Use a value between 1 and 1584.`);
    if (p.spaceBefore !== undefined && p.spaceBefore < 0) throw new Error('spaceBefore must be non-negative');
    if (p.spaceAfter !== undefined && p.spaceAfter < 0) throw new Error('spaceAfter must be non-negative');
    // Enforce upper bounds on spacing/indent values (max 1584pt = 22 inches)
    const MAX_SPACING = 1584;
    if (p.lineSpacing !== undefined && p.lineSpacing > MAX_SPACING) throw new Error(`lineSpacing value ${p.lineSpacing} exceeds maximum (${MAX_SPACING} points = 22 inches).`);
    if (p.spaceBefore !== undefined && p.spaceBefore > MAX_SPACING) throw new Error(`spaceBefore value ${p.spaceBefore} exceeds maximum (${MAX_SPACING} points = 22 inches).`);
    if (p.spaceAfter !== undefined && p.spaceAfter > MAX_SPACING) throw new Error(`spaceAfter value ${p.spaceAfter} exceeds maximum (${MAX_SPACING} points = 22 inches).`);
    if (p.firstLineIndent !== undefined && p.firstLineIndent > MAX_SPACING) throw new Error(`firstLineIndent value ${p.firstLineIndent} exceeds maximum (${MAX_SPACING} points = 22 inches).`);
    if (p.leftIndent !== undefined && p.leftIndent > MAX_SPACING) throw new Error(`leftIndent value ${p.leftIndent} exceeds maximum (${MAX_SPACING} points = 22 inches).`);
    if (p.rightIndent !== undefined && p.rightIndent > MAX_SPACING) throw new Error(`rightIndent value ${p.rightIndent} exceeds maximum (${MAX_SPACING} points = 22 inches).`);
    const paragraphs = ctx.document.body.paragraphs;
    paragraphs.load('text');
    await ctx.sync();
    if (p.index >= paragraphs.items.length) throw new Error(`Paragraph index out of range. Valid indices: 0-${paragraphs.items.length - 1} (document has ${paragraphs.items.length} paragraphs).`);
    const para = paragraphs.items[p.index];
    if (p.lineSpacing !== undefined) para.lineSpacing = p.lineSpacing;
    if (p.spaceBefore !== undefined) para.spaceBefore = p.spaceBefore;
    if (p.spaceAfter !== undefined) para.spaceAfter = p.spaceAfter;
    if (p.firstLineIndent !== undefined) para.firstLineIndent = p.firstLineIndent;
    if (p.leftIndent !== undefined) para.leftIndent = p.leftIndent;
    if (p.rightIndent !== undefined) para.rightIndent = p.rightIndent;
    para.getRange('End').select();
    await ctx.sync();
    return { success: true };
  },
};
