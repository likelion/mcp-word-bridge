import type { CommandHandler } from './index';
import { sanitizeText, checkOccurrence, checkAnchorText } from './document';

export const bookmarkCommands: Record<string, CommandHandler> = {
  async getBookmarks(ctx) {
    const range = ctx.document.body.getRange();
    try {
      const bookmarks = range.getBookmarks(true, true);
      await ctx.sync();
      return { bookmarks: bookmarks.value };
    } catch {
      try {
        const bookmarks = ctx.document.body.getBookmarks(true);
        await ctx.sync();
        return { bookmarks: bookmarks.value };
      } catch {
        throw new Error('getBookmarks is not supported on this Word platform version. Bookmarks can still be inserted and deleted by name.');
      }
    }
  },

  async insertBookmark(ctx, p) {
    checkOccurrence(p.occurrence);
    checkAnchorText(p.anchorText);
    if (!p.name || typeof p.name !== 'string' || p.name.trim() === '')
      throw new Error('Bookmark name must be a non-empty string.');
    if (!/^[A-Za-z_]\w*$/.test(p.name))
      throw new Error(`Invalid bookmark name: "${p.name}". Names must start with a letter or underscore and contain only letters, numbers, and underscores (no spaces).`);
    if (p.name.length > 40)
      throw new Error(`Bookmark name "${p.name}" exceeds Word's 40-character maximum (got ${p.name.length}).`);
    const results = ctx.document.body.search(p.anchorText, { matchCase: p.matchCase || false });
    results.load('text');
    const range = ctx.document.body.getRange();
    const existingBookmarks = range.getBookmarks(true, true);
    await ctx.sync();
    const isDuplicate = existingBookmarks.value && existingBookmarks.value.includes(p.name);
    if (results.items.length === 0) throw new Error('Anchor not found: ' + p.anchorText);
    const idx = p.occurrence || 0;
    if (idx >= results.items.length)
      throw new Error(`Occurrence ${idx} not found (only ${results.items.length} match${results.items.length === 1 ? '' : 'es'})`);
    const target = results.items[idx];
    target.insertBookmark(p.name);
    target.getRange('End').select();
    await ctx.sync();
    const result: any = { success: true };
    if (isDuplicate) result.warning = `Bookmark "${p.name}" already existed and was moved to the new location.`;
    return result;
  },

  async deleteBookmark(ctx, p) {
    const range = ctx.document.body.getRange();
    const bookmarks = range.getBookmarks(true, true);
    await ctx.sync();
    if (!bookmarks.value || !bookmarks.value.includes(p.name))
      throw new Error('Bookmark not found: ' + p.name);
    ctx.document.deleteBookmark(p.name);
    await ctx.sync();
    return { success: true };
  },

  async goToBookmark(ctx, p) {
    let range: any;
    try {
      range = ctx.document.getBookmarkRange(p.name);
      range.load('text');
      range.select();
      await ctx.sync();
    } catch (e: any) {
      if (e.message?.includes('ItemNotFound'))
        throw new Error('Bookmark not found: ' + p.name);
      throw e;
    }
    return { success: true, text: sanitizeText(range.text) };
  },

  async getBookmarkText(ctx, p) {
    let range: any;
    try {
      range = ctx.document.getBookmarkRange(p.name);
      range.load('text');
      await ctx.sync();
    } catch (e: any) {
      if (e.message?.includes('ItemNotFound'))
        throw new Error('Bookmark not found: ' + p.name);
      throw e;
    }
    return { text: sanitizeText(range.text) };
  },
};
