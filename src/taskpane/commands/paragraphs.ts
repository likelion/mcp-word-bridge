import type { CommandHandler } from './index';

const ALIGNMENT_MAP: Record<string, string> = { Left: 'Left', Center: 'Centered', Centered: 'Centered', Right: 'Right', Justify: 'Justified', Justified: 'Justified' };

export const paragraphCommands: Record<string, CommandHandler> = {
  async getParagraphs(ctx, p) {
    if (p.start !== undefined && p.start < 0) throw new Error('start index must be non-negative');
    if (p.end !== undefined && p.end < 0) throw new Error('end index must be non-negative');
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
      const isTocEntry = !!(para.style && para.style.startsWith('TOC')) || (para.style === '' && /\t\d+$/.test(para.text));
      items.push({ index: i, text: para.text, style: para.style, alignment: para.alignment, isListItem: para.isListItem, inTable, isTocEntry, outlineLevel: para.outlineLevel });
    }
    const result: any = { count: paragraphs.items.length, paragraphs: items };
    if (p.start !== undefined && p.start >= paragraphs.items.length) {
      result.warning = `start index (${p.start}) is beyond the last paragraph. Document has ${paragraphs.items.length} paragraphs (valid indices: 0-${paragraphs.items.length - 1}).`;
    }
    return result;
  },

  async getParagraphByIndex(ctx, p) {
    if (p.index < 0) throw new Error('Index must be non-negative');
    const paragraphs = ctx.document.body.paragraphs;
    paragraphs.load('text');
    await ctx.sync();
    if (p.index >= paragraphs.items.length) throw new Error(`Paragraph index out of range. Valid indices: 0-${paragraphs.items.length - 1} (document has ${paragraphs.items.length} paragraphs).`);
    const para = paragraphs.items[p.index];
    para.load('text,style,alignment,firstLineIndent,leftIndent,rightIndent,lineSpacing,spaceBefore,spaceAfter,outlineLevel,isListItem');
    para.font.load('name,size,bold,italic,color,underline');
    await ctx.sync();
    return { text: para.text, style: para.style, alignment: para.alignment, firstLineIndent: para.firstLineIndent, leftIndent: para.leftIndent, rightIndent: para.rightIndent, lineSpacing: para.lineSpacing, spaceBefore: para.spaceBefore, spaceAfter: para.spaceAfter, outlineLevel: para.outlineLevel, isListItem: para.isListItem, font: { name: para.font.name, size: para.font.size, bold: para.font.bold, italic: para.font.italic, color: para.font.color, underline: para.font.underline } };
  },

  async insertParagraph(ctx, p) {
    if (p.location && p.location !== 'Start' && p.location !== 'End') throw new Error(`Invalid location: "${p.location}". Valid values: Start, End`);
    let alignment: string | null = null;
    if (p.alignment) {
      alignment = ALIGNMENT_MAP[p.alignment] ?? null;
      if (!alignment) throw new Error(`Invalid alignment: "${p.alignment}". Valid values: Left, Center, Right, Justified`);
    }
    const styleName = p.style || 'Normal';
    if (p.style) {
      const styleObj = ctx.document.getStyles().getByNameOrNullObject(p.style);
      styleObj.load('nameLocal');
      await ctx.sync();
      if (styleObj.isNullObject) throw new Error(`Style not found: "${p.style}". Use word_get_styles to see available styles.`);
    }
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
    if (p.index < 0) throw new Error('Index must be non-negative');
    let alignment: string | null = null;
    if (p.alignment) {
      alignment = ALIGNMENT_MAP[p.alignment] ?? null;
      if (!alignment) throw new Error(`Invalid alignment: "${p.alignment}". Valid values: Left, Center, Right, Justified`);
    }
    const paragraphs = ctx.document.body.paragraphs;
    paragraphs.load('text,parentTableCellOrNullObject');
    await ctx.sync();
    if (p.index >= paragraphs.items.length) throw new Error(`Paragraph index out of range. Valid indices: 0-${paragraphs.items.length - 1} (document has ${paragraphs.items.length} paragraphs).`);
    let inTable = false;
    try { inTable = paragraphs.items[p.index].parentTableCellOrNullObject && !paragraphs.items[p.index].parentTableCellOrNullObject.isNullObject; } catch { /* ignore */ }
    const styleName = p.style || 'Normal';
    if (p.style) {
      const styleObj = ctx.document.getStyles().getByNameOrNullObject(p.style);
      styleObj.load('nameLocal');
      await ctx.sync();
      if (styleObj.isNullObject) throw new Error(`Style not found: "${p.style}". Use word_get_styles to see available styles.`);
    }
    const ref = paragraphs.items[p.index];
    const location = p.location === 'Before' ? 'Before' : 'After';
    const newPara = ref.insertParagraph(p.text, location);
    newPara.style = styleName;
    await ctx.sync();
    if (alignment) { newPara.alignment = alignment; await ctx.sync(); }
    newPara.getRange('End').select();
    await ctx.sync();
    const result: any = { success: true };
    if (inTable) result.warning = 'Paragraph inserted inside a table cell. This creates a multi-paragraph cell which may not be intended.';
    return result;
  },

  async deleteParagraph(ctx, p) {
    if (p.index < 0) throw new Error('Index must be non-negative');
    const paragraphs = ctx.document.body.paragraphs;
    paragraphs.load('text');
    await ctx.sync();
    if (p.index >= paragraphs.items.length) throw new Error(`Paragraph index out of range. Valid indices: 0-${paragraphs.items.length - 1} (document has ${paragraphs.items.length} paragraphs).`);
    const countBefore = paragraphs.items.length;
    paragraphs.items[p.index].delete();
    try { await ctx.sync(); } catch (e: any) {
      if (e.message?.includes('GeneralException')) throw new Error(`Cannot delete paragraph ${p.index}. It may be a TOC field entry or inside a protected region.`);
      throw e;
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
    if (p.index < 0) throw new Error('Index must be non-negative');
    const paragraphs = ctx.document.body.paragraphs;
    paragraphs.load('text');
    await ctx.sync();
    if (p.index >= paragraphs.items.length) throw new Error(`Paragraph index out of range. Valid indices: 0-${paragraphs.items.length - 1} (document has ${paragraphs.items.length} paragraphs).`);
    const inserted = paragraphs.items[p.index].insertText(p.text, 'Replace');
    inserted.getRange('End').select();
    try { await ctx.sync(); } catch (e: any) {
      if (e.message?.includes('GeneralException')) throw new Error(`Cannot replace text of paragraph ${p.index}. It may be a TOC field entry or inside a protected region.`);
      throw e;
    }
    return { success: true };
  },

  async setParagraphStyle(ctx, p) {
    if (p.index < 0) throw new Error('Index must be non-negative');
    let alignment: string | null = null;
    if (p.alignment) {
      alignment = ALIGNMENT_MAP[p.alignment] ?? null;
      if (!alignment) throw new Error(`Invalid alignment: "${p.alignment}". Valid values: Left, Center, Right, Justified`);
    }
    if (!p.style && !p.alignment) throw new Error('At least one of "style" or "alignment" must be provided.');
    const paragraphs = ctx.document.body.paragraphs;
    paragraphs.load('text');
    await ctx.sync();
    if (p.index >= paragraphs.items.length) throw new Error(`Paragraph index out of range. Valid indices: 0-${paragraphs.items.length - 1} (document has ${paragraphs.items.length} paragraphs).`);
    const para = paragraphs.items[p.index];
    if (p.style) {
      const styleObj = ctx.document.getStyles().getByNameOrNullObject(p.style);
      styleObj.load('nameLocal');
      await ctx.sync();
      if (styleObj.isNullObject) throw new Error(`Style not found: "${p.style}". Use word_get_styles to see available styles.`);
      para.style = p.style;
      await ctx.sync();
    }
    if (alignment) { para.alignment = alignment; await ctx.sync(); }
    para.getRange('End').select();
    await ctx.sync();
    return { success: true };
  },

  async setParagraphSpacing(ctx, p) {
    if (p.index < 0) throw new Error('Index must be non-negative');
    if (p.lineSpacing !== undefined && p.lineSpacing <= 0) throw new Error('lineSpacing must be positive');
    if (p.spaceBefore !== undefined && p.spaceBefore < 0) throw new Error('spaceBefore must be non-negative');
    if (p.spaceAfter !== undefined && p.spaceAfter < 0) throw new Error('spaceAfter must be non-negative');
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
