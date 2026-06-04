import type { CommandHandler } from './index';
import { checkAnchorText } from './document';

declare const Word: any;

export const hyperlinkCommands: Record<string, CommandHandler> = {
  async insertHyperlink(ctx, p) {
    if (!p.url || !/^https?:\/\/.+/i.test(p.url))
      throw new Error('URL must be a valid HTTP or HTTPS URL (e.g. https://example.com)');
    if (/[<>"{}|\\^`]/.test(p.url))
      throw new Error(`Malformed URL: "${p.url}". URL contains invalid characters that must be percent-encoded.`);
    checkAnchorText(p.anchorText);
    const results = ctx.document.body.search(p.anchorText, { matchCase: p.matchCase || false });
    results.load('text,hyperlink');
    await ctx.sync();
    if (results.items.length === 0) throw new Error('Anchor not found: ' + p.anchorText);
    const idx = p.occurrence || 0;
    if (idx >= results.items.length)
      throw new Error(`Occurrence ${idx} not found (only ${results.items.length} match${results.items.length === 1 ? '' : 'es'})`);
    const target = results.items[idx];
    const existingHyperlink = target.hyperlink;
    target.hyperlink = p.url;
    // Load the resulting range text to detect if Word expanded the hyperlink range
    target.load('text');
    target.getRange('End').select();
    await ctx.sync();
    const result: any = { success: true };
    const warnings: string[] = [];
    if (existingHyperlink) {
      warnings.push(`Replaced existing hyperlink "${existingHyperlink}" with "${p.url}".`);
    }
    if (target.text && target.text !== p.anchorText) {
      warnings.push(`Hyperlink applied to "${target.text}" (expanded from requested "${p.anchorText}" due to existing hyperlink range).`);
    }
    if (warnings.length > 0) result.warning = warnings.join(' ');
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
        // Strip control chars, nested PAGEREF field codes, and excess whitespace
        const cleanText = rawText
          .replace(/[\u0001\u0002\u0013\u0014\u0015]/g, '')
          .replace(/\s*PAGEREF\s+\S+\s*(\\[a-z]\s*)*/gi, '')
          .replace(/\s{2,}/g, ' ')
          .trim();
        links.push({ index: i, url, text: cleanText, internal: isInternal });
      }
    }
    return { count: links.length, hyperlinks: links };
  },

  async removeHyperlink(ctx, p) {
    checkAnchorText(p.anchorText);
    const results = ctx.document.body.search(p.anchorText, { matchCase: p.matchCase || false });
    results.load('hyperlink');
    await ctx.sync();
    if (results.items.length === 0) throw new Error('Text not found: ' + p.anchorText);
    const idx = p.occurrence || 0;
    if (idx >= results.items.length)
      throw new Error(`Occurrence ${idx} not found (only ${results.items.length} match${results.items.length === 1 ? '' : 'es'})`);
    const target = results.items[idx];
    if (!target.hyperlink || target.hyperlink === '') {
      throw new Error(`Text "${p.anchorText}" does not have a hyperlink.`);
    }
    target.hyperlink = '';
    target.getRange('End').select();
    await ctx.sync();
    return { success: true };
  },
};
