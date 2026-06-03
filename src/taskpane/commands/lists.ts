import type { CommandHandler } from './index';

declare const Word: any;

export const listCommands: Record<string, CommandHandler> = {
  async insertList(ctx, p) {
    if (!p.items || p.items.length === 0) throw new Error('items array must not be empty');
    const body = ctx.document.body;
    const location = p.location || 'End';
    if (location === 'End') {
      const lastPara = body.paragraphs.getLast();
      lastPara.load('isListItem');
      await ctx.sync();
      if (lastPara.isListItem) {
        const sep = body.insertParagraph('', 'End');
        sep.style = 'Normal';
        sep.detachFromList();
        await ctx.sync();
      }
    } else if (location === 'Start') {
      const firstPara = body.paragraphs.getFirst();
      firstPara.load('isListItem');
      await ctx.sync();
      if (firstPara.isListItem) {
        const sep = body.insertParagraph('', 'Start');
        sep.style = 'Normal';
        sep.detachFromList();
        await ctx.sync();
      }
    }
    const para = body.insertParagraph(p.items[0], location);
    para.style = 'Normal';
    await ctx.sync();
    const list = para.startNewList();
    for (let i = 1; i < p.items.length; i++) {
      list.insertParagraph(p.items[i], Word.InsertLocation.end);
    }
    if (p.numbered) {
      list.setLevelNumbering(0, Word.ListNumbering.arabic);
    } else {
      list.setLevelBullet(0, Word.ListBullet.solid);
    }
    await ctx.sync();
    const endPara = body.paragraphs.getLast();
    endPara.getRange('End').select();
    await ctx.sync();
    return { success: true };
  },

  async getListInfo(ctx, p) {
    if (p.index < 0) throw new Error('Index must be non-negative');
    const paragraphs = ctx.document.body.paragraphs;
    paragraphs.load('text,isListItem');
    await ctx.sync();
    if (p.index >= paragraphs.items.length)
      throw new Error(`Paragraph index out of range. Valid indices: 0-${paragraphs.items.length - 1} (document has ${paragraphs.items.length} paragraphs).`);
    const para = paragraphs.items[p.index];
    if (!para.isListItem) return { isListItem: false };
    para.load('listItem');
    await ctx.sync();
    const li = para.listItem;
    li.load('level,listString,siblingIndex');
    await ctx.sync();
    return { isListItem: true, level: li.level, listString: li.listString, siblingIndex: li.siblingIndex };
  },

  async setListLevel(ctx, p) {
    if (p.index < 0) throw new Error('Index must be non-negative');
    if (p.level < 0 || p.level > 8) throw new Error('level must be between 0 and 8 (inclusive)');
    const paragraphs = ctx.document.body.paragraphs;
    paragraphs.load('text,isListItem');
    await ctx.sync();
    if (p.index >= paragraphs.items.length)
      throw new Error(`Paragraph index out of range. Valid indices: 0-${paragraphs.items.length - 1} (document has ${paragraphs.items.length} paragraphs).`);
    const para = paragraphs.items[p.index];
    if (!para.isListItem) throw new Error('Paragraph is not a list item');
    para.load('listItem');
    await ctx.sync();
    para.listItem.level = p.level;
    para.getRange('End').select();
    await ctx.sync();
    return { success: true };
  },
};
