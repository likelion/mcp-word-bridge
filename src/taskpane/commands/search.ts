import type { CommandHandler } from './index';
import { checkSearchLength } from './document';

declare const Word: any;

export const searchCommands: Record<string, CommandHandler> = {
  async search(ctx, p) {
    if (!p.query || typeof p.query !== 'string' || p.query.trim() === '')
      throw new Error('Search query cannot be empty. Provide a non-empty search string.');
    const results = ctx.document.body.search(p.query, { matchCase: p.matchCase || false, matchWholeWord: p.matchWholeWord || false });
    results.load('text');
    try {
      await ctx.sync();
    } catch (e: any) {
      if (e.message?.includes('SearchStringInvalidOrTooLong'))
        throw new Error('Search query is too long (max ~255 characters). Shorten the query text.');
      throw e;
    }
    return { count: results.items.length, matches: results.items.slice(0, 30).map((r: any, i: number) => ({ index: i, text: r.text })) };
  },

  async searchAndReplace(ctx, p) {
    if (!p.find || typeof p.find !== 'string' || p.find.trim() === '')
      throw new Error('find string cannot be empty. Provide a non-empty search string.');
    checkSearchLength(p.find);
    // BUG-01: Reject Word special codes that can corrupt document structure
    const specialCodePattern = /\^(p|w|t|l|m|b|n|s|d|a|e|f|g|v|~|\^|\-|13|11|14|12|07|09)/;
    const findMatch = (p.find as string).match(specialCodePattern);
    if (findMatch) throw new Error(`find contains Word special code "${findMatch[0]}" which can corrupt document structure. Use literal text only. Common special codes: ^p (paragraph mark), ^t (tab), ^w (whitespace), ^13 (paragraph mark).`);
    const replaceMatch = (p.replace as string).match(specialCodePattern);
    if (replaceMatch) throw new Error(`replace contains Word special code "${replaceMatch[0]}" which can corrupt document structure. Use literal text only.`);
    const results = ctx.document.body.search(p.find, { matchCase: p.matchCase || false, matchWholeWord: p.matchWholeWord || false });
    results.load('text');
    await ctx.sync();
    const count = results.items.length;
    let lastRange: any = null;
    for (let i = 0; i < count; i++) {
      lastRange = results.items[i].insertText(p.replace, Word.InsertLocation.replace);
    }
    if (lastRange) lastRange.getRange('End').select();
    await ctx.sync();
    return { replacements: count };
  },

  async insertText(ctx, p) {
    if (p.after && p.before) throw new Error('Provide only one of "after" or "before", not both.');
    if (p.occurrence !== undefined && p.occurrence < 0) throw new Error('occurrence must be non-negative (0-indexed)');
    const anchor = p.after || p.before;
    if (!anchor) throw new Error('Either "after" or "before" anchor text must be provided');
    checkSearchLength(anchor);
    const results = ctx.document.body.search(anchor, { matchCase: p.matchCase || false });
    results.load('text');
    await ctx.sync();
    if (results.items.length === 0) throw new Error('Anchor not found: ' + anchor);
    const idx = p.occurrence || 0;
    if (idx >= results.items.length)
      throw new Error(`Occurrence index ${idx} is out of range (valid: 0 to ${results.items.length - 1} for ${results.items.length} match${results.items.length === 1 ? '' : 'es'})`);
    const target = results.items[idx];
    const loc = p.after ? Word.InsertLocation.after : Word.InsertLocation.before;
    const inserted = target.insertText(p.text, loc);
    inserted.getRange('End').select();
    await ctx.sync();
    return { success: true };
  },

  async getSelectionInfo(ctx) {
    const sel = ctx.document.getSelection();
    sel.load('text,style,isEmpty');
    sel.font.load('name,size,bold,italic,color,underline,strikeThrough,highlightColor');
    await ctx.sync();
    return {
      text: sel.text, style: sel.style, isEmpty: sel.isEmpty,
      font: { name: sel.font.name, size: sel.font.size, bold: sel.font.bold, italic: sel.font.italic, color: sel.font.color, underline: sel.font.underline, strikeThrough: sel.font.strikeThrough, highlightColor: sel.font.highlightColor },
    };
  },

  async insertTextAtSelection(ctx, p) {
    const sel = ctx.document.getSelection();
    const loc = p.replace ? Word.InsertLocation.replace : Word.InsertLocation.end;
    const inserted = sel.insertText(p.text, loc);
    inserted.getRange('End').select();
    await ctx.sync();
    return { success: true };
  },

  async insertLineBreak(ctx, p) {
    if (p.occurrence !== undefined && p.occurrence < 0) throw new Error('occurrence must be non-negative (0-indexed)');
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
    const loc = p.before ? Word.InsertLocation.before : Word.InsertLocation.after;
    target.insertBreak('Line', loc);
    target.getRange('End').select();
    await ctx.sync();
    return { success: true };
  },
};
