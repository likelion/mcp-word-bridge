import type { CommandHandler } from './index';

export const trackingCommands: Record<string, CommandHandler> = {
  async getTrackedChanges(ctx) {
    const changes = ctx.document.body.getTrackedChanges();
    changes.load('type,author,date,text');
    await ctx.sync();
    const items = changes.items.map((c: any, i: number) => ({ index: i, type: c.type, author: c.author, date: c.date, text: c.text }));
    return { count: items.length, changes: items };
  },

  async acceptTrackedChange(ctx, p) {
    if (p.index < 0) throw new Error('Index must be non-negative');
    const changes = ctx.document.body.getTrackedChanges();
    changes.load('items');
    await ctx.sync();
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
    changes.load('items');
    await ctx.sync();
    changes.acceptAll();
    await ctx.sync();
    return { success: true };
  },

  async rejectAllTrackedChanges(ctx) {
    const changes = ctx.document.body.getTrackedChanges();
    changes.load('items');
    await ctx.sync();
    changes.rejectAll();
    await ctx.sync();
    return { success: true };
  },

  async getChangeTrackingMode(ctx) {
    const doc = ctx.document;
    doc.load('changeTrackingMode');
    await ctx.sync();
    return { mode: doc.changeTrackingMode };
  },
};
