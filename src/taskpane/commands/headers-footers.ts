import type { CommandHandler } from './index';

declare const Word: any;

export const headerFooterCommands: Record<string, CommandHandler> = {
  async getHeaderFooter(ctx, p) {
    const sections = ctx.document.sections;
    sections.load('items');
    await ctx.sync();
    const idx = p.sectionIndex || 0;
    if (idx < 0 || idx >= sections.items.length)
      throw new Error(`Section index out of range. Document has ${sections.items.length} section(s) (0-indexed).`);
    const section = sections.items[idx];
    const target = p.type === 'footer'
      ? section.getFooter(p.headerType || 'Primary')
      : section.getHeader(p.headerType || 'Primary');
    target.load('text');
    await ctx.sync();
    return { text: target.text };
  },

  async setHeaderFooter(ctx, p) {
    const sections = ctx.document.sections;
    sections.load('items');
    await ctx.sync();
    const idx = p.sectionIndex || 0;
    if (idx < 0 || idx >= sections.items.length)
      throw new Error(`Section index out of range. Document has ${sections.items.length} section(s) (0-indexed).`);
    const section = sections.items[idx];
    const target = p.type === 'footer'
      ? section.getFooter(p.headerType || 'Primary')
      : section.getHeader(p.headerType || 'Primary');
    target.clear();
    target.insertText(p.text, Word.InsertLocation.start);
    await ctx.sync();
    return { success: true };
  },
};
