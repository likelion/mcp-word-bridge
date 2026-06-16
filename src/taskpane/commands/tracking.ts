import type { CommandHandler } from './index';
import { sanitizeText } from './document';

export const trackingCommands: Record<string, CommandHandler> = {
  async getTrackedChanges(ctx) {
    // TrackedChangeCollection has no 'count' property — must load items.
    // The 'type' property throws GeneralException on some tracked change types
    // (formatting-only changes, complex content). We load items, then batch-load
    // only safe properties (author, date, text) and attempt type with fallback.
    const changes = ctx.document.body.getTrackedChanges();
    try {
      changes.load('items');
      await ctx.sync();
    } catch {
      // If load('items') itself fails, fall back to acceptAll-friendly approach
      return { count: -1, changes: [], error: 'Unable to enumerate tracked changes. Use accept_all or accept_in_range.' };
    }

    const count = changes.items.length;
    if (count === 0) return { count: 0, changes: [] };

    // Batch-load safe properties
    for (let i = 0; i < count; i++) {
      changes.items[i].load('author,date,text');
    }
    await ctx.sync();

    const results = [];
    for (let i = 0; i < count; i++) {
      const c = changes.items[i];
      const entry: any = {
        index: i,
        type: 'Unknown',
        author: c.author ?? '',
        date: c.date ?? '',
        text: sanitizeText(c.text ?? ''),
      };
      results.push(entry);
    }

    // Attempt to load type for each item individually (this is what fails on some docs)
    for (let i = 0; i < count; i++) {
      try {
        changes.items[i].load('type');
      } catch { /* skip */ }
    }
    try {
      await ctx.sync();
      for (let i = 0; i < count; i++) {
        try {
          results[i].type = changes.items[i].type ?? 'Unknown';
        } catch { /* keep Unknown */ }
      }
    } catch {
      // type batch failed — keep all as Unknown
    }

    return { count: results.length, changes: results };
  },

  async acceptTrackedChange(ctx, p) {
    if (p.index < 0) throw new Error('Index must be non-negative');
    const changes = ctx.document.body.getTrackedChanges();
    changes.load('items');
    await ctx.sync();
    if (p.index >= changes.items.length)
      throw new Error(`Change index out of range. Document has ${changes.items.length} tracked change(s).`);
    changes.items[p.index].accept();
    await ctx.sync();
    return { success: true };
  },

  async rejectTrackedChange(ctx, p) {
    if (p.index < 0) throw new Error('Index must be non-negative');
    const changes = ctx.document.body.getTrackedChanges();
    changes.load('items');
    await ctx.sync();
    if (p.index >= changes.items.length)
      throw new Error(`Change index out of range. Document has ${changes.items.length} tracked change(s).`);
    changes.items[p.index].reject();
    await ctx.sync();
    return { success: true };
  },

  async acceptAllTrackedChanges(ctx) {
    const changes = ctx.document.body.getTrackedChanges();
    // Do NOT load items first — acceptAll works on the collection directly
    // and avoids triggering property serialization that can fail on complex changes
    changes.acceptAll();
    await ctx.sync();
    return { success: true, count: -1 };
  },

  async rejectAllTrackedChanges(ctx) {
    const changes = ctx.document.body.getTrackedChanges();
    // Do NOT load items first — rejectAll works on the collection directly
    changes.rejectAll();
    await ctx.sync();
    return { success: true, count: -1 };
  },

  async getChangeTrackingMode(ctx) {
    const doc = ctx.document;
    doc.load('changeTrackingMode');
    await ctx.sync();
    return { mode: doc.changeTrackingMode };
  },

  /**
   * Accept tracked changes within a paragraph index range.
   * Gets the range spanning [startIndex, endIndex) paragraphs, retrieves
   * tracked changes scoped to that range, and accepts them all.
   */
  async acceptTrackedChangesInRange(ctx, p) {
    const startIndex = p.startIndex ?? 0;
    const endIndex = p.endIndex ?? Number.MAX_SAFE_INTEGER;

    if (startIndex < 0) throw new Error('startIndex must be non-negative');
    if (endIndex <= startIndex) throw new Error('endIndex must be greater than startIndex');

    // Get paragraphs to build target range
    const paragraphs = ctx.document.body.paragraphs;
    paragraphs.load('items');
    await ctx.sync();

    const totalParas = paragraphs.items.length;
    const effectiveStart = Math.min(startIndex, totalParas - 1);
    const effectiveEnd = Math.min(endIndex - 1, totalParas - 1);

    const startRange = paragraphs.items[effectiveStart].getRange('Start');
    const endRange = paragraphs.items[effectiveEnd].getRange('End');
    const targetRange = startRange.expandTo(endRange);

    // Get tracked changes scoped to this range and accept them all
    const scopedChanges = targetRange.getTrackedChanges();
    scopedChanges.acceptAll();
    await ctx.sync();

    return { success: true };
  },
};
