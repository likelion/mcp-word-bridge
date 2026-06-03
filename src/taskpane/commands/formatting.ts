import type { CommandHandler } from './index';
import { checkSearchLength } from './document';

export const formattingCommands: Record<string, CommandHandler> = {
  async formatRange(ctx, p) {
    if (p.occurrence !== undefined && p.occurrence < 0) throw new Error('occurrence must be non-negative (0-indexed)');
    if (!p.text || typeof p.text !== 'string' || p.text.trim() === '')
      throw new Error('text to format cannot be empty. Provide a non-empty search string.');
    if (p.size !== undefined && p.size <= 0) throw new Error('size must be positive');
    if (p.size !== undefined && p.size > 1638) throw new Error('size must not exceed 1638 points (Word maximum)');
    if (p.color && !/^#[0-9A-Fa-f]{6}$/.test(p.color)) throw new Error('color must be a valid hex color (e.g. #FF0000)');
    checkSearchLength(p.text);
    const results = ctx.document.body.search(p.text, { matchCase: p.matchCase || false });
    results.load('font');
    await ctx.sync();
    if (results.items.length === 0) throw new Error('Text not found: ' + p.text);
    const idx = p.occurrence || 0;
    if (idx >= results.items.length)
      throw new Error(`Occurrence ${idx} not found (only ${results.items.length} match${results.items.length === 1 ? '' : 'es'})`);
    const target = results.items[idx];
    const font = target.font;
    if (p.bold !== undefined) font.bold = p.bold;
    if (p.italic !== undefined) font.italic = p.italic;
    if (p.underline !== undefined) font.underline = p.underline ? 'Single' : 'None';
    if (p.strikeThrough !== undefined) font.strikeThrough = p.strikeThrough;
    if (p.color) font.color = p.color;
    if (p.highlightColor) font.highlightColor = p.highlightColor;
    if (p.size) font.size = p.size;
    if (p.name) font.name = p.name;
    target.getRange('End').select();
    await ctx.sync();
    return { success: true };
  },

  async clearFormatting(ctx, p) {
    if (p.occurrence !== undefined && p.occurrence < 0) throw new Error('occurrence must be non-negative (0-indexed)');
    if (!p.text || typeof p.text !== 'string' || p.text.trim() === '')
      throw new Error('text cannot be empty. Provide a non-empty search string.');
    checkSearchLength(p.text);
    const results = ctx.document.body.search(p.text, { matchCase: p.matchCase || false });
    results.load('font,style');
    await ctx.sync();
    if (results.items.length === 0) throw new Error('Text not found: ' + p.text);
    const idx = p.occurrence || 0;
    if (idx >= results.items.length)
      throw new Error(`Occurrence ${idx} not found (only ${results.items.length} match${results.items.length === 1 ? '' : 'es'})`);
    const target = results.items[idx];
    const styleName = target.style || 'Normal';
    const styleObj = ctx.document.getStyles().getByNameOrNullObject(styleName);
    styleObj.load('font/size,font/name');
    await ctx.sync();
    const font = target.font;
    font.bold = false;
    font.italic = false;
    font.underline = 'None';
    font.strikeThrough = false;
    font.color = '#000000';
    if (!styleObj.isNullObject) {
      if (styleObj.font.size) font.size = styleObj.font.size;
      if (styleObj.font.name) font.name = styleObj.font.name;
    }
    target.getRange('End').select();
    await ctx.sync();
    return { success: true };
  },

  async getFontInfo(ctx, p) {
    if (p.occurrence !== undefined && p.occurrence < 0) throw new Error('occurrence must be non-negative (0-indexed)');
    if (!p.text || typeof p.text !== 'string' || p.text.trim() === '')
      throw new Error('text cannot be empty. Provide a non-empty search string.');
    checkSearchLength(p.text);
    const results = ctx.document.body.search(p.text, { matchCase: p.matchCase || false });
    results.load('font');
    await ctx.sync();
    if (results.items.length === 0) throw new Error('Text not found: ' + p.text);
    const idx = p.occurrence || 0;
    if (idx >= results.items.length)
      throw new Error(`Occurrence ${idx} not found (only ${results.items.length} match${results.items.length === 1 ? '' : 'es'})`);
    const target = results.items[idx];
    const font = target.font;
    font.load('name,size,bold,italic,underline,strikeThrough,color,highlightColor');
    await ctx.sync();
    return { name: font.name, size: font.size, bold: font.bold, italic: font.italic, underline: font.underline, strikeThrough: font.strikeThrough, color: font.color, highlightColor: font.highlightColor };
  },
};
