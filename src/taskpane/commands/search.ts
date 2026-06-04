import type { CommandHandler } from './index';
import { checkSearchLength, sanitizeText, checkOccurrence, checkAnchorText } from './document';
import { insertWithCodes } from './special-codes';

declare const Word: any;

export const searchCommands: Record<string, CommandHandler> = {
  async search(ctx, p) {
    if (!p.query || typeof p.query !== 'string' || p.query.trim() === '')
      throw new Error('Search query cannot be empty. Provide a non-empty search string.');
    const results = ctx.document.body.search(p.query, {
      matchCase: p.matchCase || false,
      matchWholeWord: p.matchWholeWord || false,
      matchWildcards: p.matchWildcards || false,
      matchPrefix: p.matchPrefix || false,
      matchSuffix: p.matchSuffix || false,
      ignorePunct: p.ignorePunct || false,
      ignoreSpace: p.ignoreSpace || false,
    });
    results.load('text');
    try {
      await ctx.sync();
    } catch (e: any) {
      if (e.message?.includes('SearchStringInvalidOrTooLong'))
        throw new Error('Search query is too long (max ~255 characters). Shorten the query text.');
      throw e;
    }
    return { count: results.items.length, matches: results.items.slice(0, 30).map((r: any, i: number) => ({ index: i, text: sanitizeText(r.text) })) };
  },

  async searchAndReplace(ctx, p) {
    if (!p.find || typeof p.find !== 'string' || p.find.trim() === '')
      throw new Error('find string cannot be empty. Provide a non-empty search string.');
    checkSearchLength(p.find);

    // Snapshot bookmarks before replacement
    let bookmarksBefore: string[] = [];
    try {
      const bodyRange = ctx.document.body.getRange();
      const bm = bodyRange.getBookmarks(true, true);
      await ctx.sync();
      bookmarksBefore = bm.value || [];
    } catch { /* older API — skip bookmark tracking */ }

    // If preserveBookmarks, collect which bookmarks are on each match
    const preserveBookmarks = !!p.preserveBookmarks;
    const matchBookmarks: string[][] = [];

    const results = ctx.document.body.search(p.find, {
      matchCase: p.matchCase || false,
      matchWholeWord: p.matchWholeWord || false,
      matchWildcards: p.matchWildcards || false,
      matchPrefix: p.matchPrefix || false,
      matchSuffix: p.matchSuffix || false,
      ignorePunct: p.ignorePunct || false,
      ignoreSpace: p.ignoreSpace || false,
    });
    results.load('text');
    await ctx.sync();
    const count = results.items.length;

    if (preserveBookmarks && count > 0 && bookmarksBefore.length > 0) {
      // For each match range, find bookmarks that intersect it
      for (let i = 0; i < count; i++) {
        try {
          const rangeBm = results.items[i].getBookmarks(true, true);
          await ctx.sync();
          matchBookmarks.push(rangeBm.value || []);
        } catch {
          matchBookmarks.push([]);
        }
      }
    }

    // Perform the replacements — skip matches where text already equals replace
    // string exactly (prevents case-insensitive replace from altering casing)
    const insertedRanges: any[] = [];
    let actualReplacements = 0;
    for (let i = 0; i < count; i++) {
      if (results.items[i].text === p.replace) {
        // Already identical — no replacement needed, push original range for bookmark tracking
        insertedRanges.push(results.items[i]);
        continue;
      }
      const inserted = await insertWithCodes(results.items[i], p.replace, Word.InsertLocation.replace, ctx);
      insertedRanges.push(inserted);
      actualReplacements++;
    }
    if (insertedRanges.length > 0) insertedRanges[insertedRanges.length - 1].getRange('End').select();
    if (actualReplacements > 0) await ctx.sync();

    // Re-create bookmarks on replacement text if requested
    let bookmarksRestored = 0;
    if (preserveBookmarks && matchBookmarks.length > 0) {
      for (let i = 0; i < matchBookmarks.length; i++) {
        const bmNames = matchBookmarks[i] || [];
        for (const bmName of bmNames) {
          try {
            insertedRanges[i].insertBookmark(bmName);
            bookmarksRestored++;
          } catch { /* bookmark restore failed — skip */ }
        }
      }
      if (bookmarksRestored > 0) await ctx.sync();
    }

    // Check how many bookmarks were lost
    let bookmarksAfter: string[] = [];
    try {
      const bodyRange2 = ctx.document.body.getRange();
      const bm2 = bodyRange2.getBookmarks(true, true);
      await ctx.sync();
      bookmarksAfter = bm2.value || [];
    } catch { /* skip */ }

    const result: any = { replacements: actualReplacements };
    if (count > actualReplacements) {
      result.skipped = count - actualReplacements;
    }
    const lost = bookmarksBefore.filter(b => !bookmarksAfter.includes(b));
    if (lost.length > 0 && !preserveBookmarks) {
      result.warning = `${lost.length} bookmark(s) destroyed by this replacement: ${lost.join(', ')}. Use preserveBookmarks: true to re-create them on the replacement text.`;
      result.bookmarksLost = lost.length;
    } else if (bookmarksRestored > 0) {
      result.bookmarksRestored = bookmarksRestored;
    }

    return result;
  },

  async insertText(ctx, p) {
    if (p.after && p.before) throw new Error('Provide only one of "after" or "before", not both.');
    checkOccurrence(p.occurrence);
    const anchor = p.after || p.before;
    if (!anchor) throw new Error('Either "after" or "before" anchor text must be provided');
    checkSearchLength(anchor);
    const results = ctx.document.body.search(anchor, {
      matchCase: p.matchCase || false,
      matchWildcards: p.matchWildcards || false,
      matchPrefix: p.matchPrefix || false,
      matchSuffix: p.matchSuffix || false,
      ignorePunct: p.ignorePunct || false,
      ignoreSpace: p.ignoreSpace || false,
    });
    results.load('text');
    await ctx.sync();
    if (results.items.length === 0) throw new Error('Anchor not found: ' + anchor);
    const idx = p.occurrence || 0;
    if (idx >= results.items.length)
      throw new Error(`Occurrence index ${idx} is out of range (valid: 0 to ${results.items.length - 1} for ${results.items.length} match${results.items.length === 1 ? '' : 'es'})`);
    const target = results.items[idx];
    const loc = p.after ? Word.InsertLocation.after : Word.InsertLocation.before;
    const inserted = await insertWithCodes(target, p.text, loc, ctx);
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
      text: sanitizeText(sel.text), style: sel.style, isEmpty: sel.isEmpty,
      font: { name: sel.font.name, size: sel.font.size, bold: sel.font.bold, italic: sel.font.italic, color: sel.font.color, underline: sel.font.underline, strikeThrough: sel.font.strikeThrough, highlightColor: sel.font.highlightColor },
    };
  },

  async insertTextAtSelection(ctx, p) {
    if (!p.text && p.text !== 0) {
      return { success: true, warning: 'Empty text — no content inserted.' };
    }
    const textStr = String(p.text);
    if (textStr === '') {
      return { success: true, warning: 'Empty text — no content inserted.' };
    }
    const sel = ctx.document.getSelection();
    const loc = p.replace ? Word.InsertLocation.replace : Word.InsertLocation.end;
    const inserted = sel.insertText(textStr, loc);
    inserted.getRange('End').select();
    await ctx.sync();
    return { success: true };
  },

  async insertLineBreak(ctx, p) {
    checkOccurrence(p.occurrence);
    checkAnchorText(p.anchorText);
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
