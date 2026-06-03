import type { CommandHandler } from './index';

export const fieldCommands: Record<string, CommandHandler> = {
  async getFields(ctx) {
    const fields = ctx.document.body.fields;
    fields.load('code,type');
    await ctx.sync();
    for (let i = 0; i < fields.items.length; i++) {
      fields.items[i].result.load('text');
    }
    await ctx.sync();
    const items: any[] = [];
    for (let i = 0; i < fields.items.length; i++) {
      const f = fields.items[i];
      const code = (f.code || '').replace(/[\u0001\u0013\u0014\u0015]/g, '').trim();
      const resultText = (f.result.text || '').replace(/[\u0001\u0002\u0013\u0014\u0015]/g, '').trim();
      items.push({ index: i, code, result: resultText, type: f.type });
    }
    return { count: items.length, fields: items };
  },
};
