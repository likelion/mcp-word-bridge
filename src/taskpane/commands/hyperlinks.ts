import type { CommandHandler } from './index';
import { checkSearchLength } from './document';

declare const Word: any;

export const hyperlinkCommands: Record<string, CommandHandler> = {
  async insertHyperlink(ctx, p) {
    if (!p.url || !/^https?:\/\/.+/i.test(p.url))
      throw new Error('URL must be a valid HTTP or HTTPS URL (e.g. https://example.com)');
    if (/[<>"{}|\\^`]/.test(p.url))
      throw new Error(`Malformed URL: "${p.url}". URL contains invalid characters that must be percent-encoded.`);
    if (!p.anchorText || typeof p.anchorText !== 'string' || p.anchorText.trim() === '')
      throw new Error('anchorText cannot be empty. Provide a non-empty search string.');
    checkSearchLength(p.anchorText);
    const results = ctx.document.body.search(p.anchorText, { matchCase: p.matchCase || false });
    results.load('text,hyperlink');
    await ctx.sync();
    if (results.items.length === 0) throw new Error('Anchor not found: ' + p.anchorText);
    const idx = p.occurrence || 0;
    if (idx >= results.items.length)
      throw new Error(`Occurrence ${idx} not found (only ${results.items.length} match${results.items.length === 1 ? '' : 'es'})`);
    const target = results.items[idx];
    // BUG-14: Warn if the text already has a hyperlink
    const existingHyperlink = target.hyperlink;
    target.hyperlink = p.url;
    target.getRange('End').select();
    await ctx.sync();
    const result: any = { success: true };
    if (existingHyperlink) {
      result.warning = `Replaced existing hyperlink "${existingHyperlink}" with "${p.url}".`;
    }
    return result;
  },

  async getHyperlinks(ctx) {
    const fields = ctx.document.body.fields;
    fields.load('code,type');
    await ctx.sync();
    const links: any[] = [];
    for (let i = 0; i < fields.items.length; i++) {
      if (fields.items[i].type === 'Hyperlink') {
        const code = (fields.items[i].code || '').replace(/[\u0001\u0013\u0014\u0015]/g, '').trim();
        const match = code.match(/HYPERLINK\s+"([^"]+)"/);
        const url = match ? match[1] : '';
        const isInternal = /HYPERLINK\s+\\l\s+/.test(code);
        fields.items[i].result.load('text');
        await ctx.sync();
        const rawText = fields.items[i].result.text || '';
        const cleanText = rawText.replace(/[\u0001\u0002\u0013\u0014\u0015]/g, '').replace(/\s{2,}/g, ' ').trim();
        links.push({ index: i, url, text: cleanText, internal: isInternal });
      }
    }
    return { count: links.length, hyperlinks: links };
  },

  async removeHyperlink(ctx, p) {
    if (!p.anchorText || typeof p.anchorText !== 'string' || p.anchorText.trim() === '')
      throw new Error('anchorText cannot be empty. Provide a non-empty search string.');
    checkSearchLength(p.anchorText);
    const results = ctx.document.body.search(p.anchorText, { matchCase: p.matchCase || false });
    results.load('hyperlink');
    await ctx.sync();
    if (results.items.length === 0) throw new Error('Text not found: ' + p.anchorText);
    const idx = p.occurrence || 0;
    if (idx >= results.items.length)
      throw new Error(`Occurrence ${idx} not found (only ${results.items.length} match${results.items.length === 1 ? '' : 'es'})`);
    const target = results.items[idx];
    target.hyperlink = '';
    target.getRange('End').select();
    await ctx.sync();
    return { success: true };
  },
};
