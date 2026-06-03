import type { CommandHandler } from './index';
import { checkSearchLength } from './document';

export const footnoteCommands: Record<string, CommandHandler> = {
  async insertFootnote(ctx, p) {
    if (!p.text || typeof p.text !== 'string' || p.text.trim() === '')
      throw new Error('Footnote text must be a non-empty string');
    if (!p.anchorText || typeof p.anchorText !== 'string' || p.anchorText.trim() === '')
      throw new Error('anchorText cannot be empty. Provide a non-empty search string.');
    checkSearchLength(p.anchorText);
    const results = ctx.document.body.search(p.anchorText, { matchCase: p.matchCase || false });
    results.load('text');
    await ctx.sync();
    if (results.items.length === 0) throw new Error('Anchor not found: ' + p.anchorText);
    const idx = p.occurrence || 0;
    if (idx >= results.items.length)
      throw new Error(`Occurrence ${idx} not found (only ${results.items.length} match${results.items.length === 1 ? '' : 'es'})`);
    const target = results.items[idx];
    target.insertFootnote(p.text);
    target.getRange('End').select();
    await ctx.sync();
    return { success: true };
  },

  async insertEndnote(ctx, p) {
    if (!p.text || typeof p.text !== 'string' || p.text.trim() === '')
      throw new Error('Endnote text must be a non-empty string');
    if (!p.anchorText || typeof p.anchorText !== 'string' || p.anchorText.trim() === '')
      throw new Error('anchorText cannot be empty. Provide a non-empty search string.');
    checkSearchLength(p.anchorText);
    const results = ctx.document.body.search(p.anchorText, { matchCase: p.matchCase || false });
    results.load('text');
    await ctx.sync();
    if (results.items.length === 0) throw new Error('Anchor not found: ' + p.anchorText);
    const idx = p.occurrence || 0;
    if (idx >= results.items.length)
      throw new Error(`Occurrence ${idx} not found (only ${results.items.length} match${results.items.length === 1 ? '' : 'es'})`);
    const target = results.items[idx];
    target.insertEndnote(p.text);
    target.getRange('End').select();
    await ctx.sync();
    return { success: true };
  },

  async getFootnotes(ctx) {
    const footnotes = ctx.document.body.footnotes;
    footnotes.load('items');
    await ctx.sync();
    footnotes.load('items');
    await ctx.sync();
    const items: any[] = [];
    for (const fn of footnotes.items) {
      fn.body.load('text');
      await ctx.sync();
      items.push({ index: items.length, text: fn.body.text.replace(/^\u0002\s*/, '') });
    }
    return { count: items.length, footnotes: items };
  },

  async getEndnotes(ctx) {
    const endnotes = ctx.document.body.endnotes;
    endnotes.load('items');
    await ctx.sync();
    endnotes.load('items');
    await ctx.sync();
    const items: any[] = [];
    for (const en of endnotes.items) {
      en.body.load('text');
      await ctx.sync();
      items.push({ index: items.length, text: en.body.text.replace(/^\u0002\s*/, '') });
    }
    return { count: items.length, endnotes: items };
  },

  async deleteFootnote(ctx, p) {
    if (p.index < 0) throw new Error('Index must be non-negative');
    const footnotes = ctx.document.body.footnotes;
    footnotes.load('items');
    await ctx.sync();
    if (p.index >= footnotes.items.length)
      throw new Error(`Footnote index out of range. Document has ${footnotes.items.length} footnote(s) (0-indexed).`);
    footnotes.items[p.index].delete();
    await ctx.sync();
    return { success: true };
  },

  async deleteEndnote(ctx, p) {
    if (p.index < 0) throw new Error('Index must be non-negative');
    const endnotes = ctx.document.body.endnotes;
    endnotes.load('items');
    await ctx.sync();
    if (p.index >= endnotes.items.length)
      throw new Error(`Endnote index out of range. Document has ${endnotes.items.length} endnote(s) (0-indexed).`);
    endnotes.items[p.index].delete();
    await ctx.sync();
    return { success: true };
  },

  async insertFootnoteAtIndex(ctx, p) {
    if (p.paragraphIndex < 0) throw new Error('Index must be non-negative');
    const paragraphs = ctx.document.body.paragraphs;
    paragraphs.load('text');
    await ctx.sync();
    if (p.paragraphIndex >= paragraphs.items.length)
      throw new Error(`Paragraph index out of range. Valid indices: 0-${paragraphs.items.length - 1} (document has ${paragraphs.items.length} paragraphs).`);
    const para = paragraphs.items[p.paragraphIndex];
    const range = para.getRange('End');
    range.insertFootnote(p.text);
    range.select();
    await ctx.sync();
    return { success: true };
  },
};
