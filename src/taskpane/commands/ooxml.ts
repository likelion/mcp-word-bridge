import type { CommandHandler } from './index';

declare const Word: any;

export const ooxmlCommands: Record<string, CommandHandler> = {
  async insertHtml(ctx, p) {
    if (!p.html || typeof p.html !== 'string' || p.html.trim() === '')
      throw new Error('html parameter must be a non-empty string');
    const body = ctx.document.body;
    const loc = p.location || 'End';
    const range = body.insertHtml(p.html, loc);
    range.getRange('End').select();
    await ctx.sync();
    return { success: true };
  },

  async insertOoxml(ctx, p) {
    if (!p.ooxml || typeof p.ooxml !== 'string' || p.ooxml.trim() === '')
      throw new Error('ooxml parameter must be a non-empty string');
    if (!p.ooxml.includes('pkg:package') && !p.ooxml.includes('pkg:part'))
      throw new Error('Invalid OOXML: missing pkg:package structure. The XML must be a valid Office Open XML flat package (containing pkg:package and pkg:part elements with a word/document.xml part).');
    const body = ctx.document.body;
    const range = body.insertOoxml(p.ooxml, p.location || 'End');
    range.getRange('End').select();
    try {
      await ctx.sync();
    } catch (e: any) {
      if (e.message?.includes('GeneralException'))
        throw new Error('Invalid OOXML. Ensure the XML follows the Office Open XML package structure (pkg:package with word/document.xml part).');
      throw e;
    }
    return { success: true };
  },

  async insertOoxmlAtSelection(ctx, p) {
    const sel = ctx.document.getSelection();
    const range = sel.insertOoxml(p.ooxml, Word.InsertLocation.replace);
    range.getRange('End').select();
    try {
      await ctx.sync();
    } catch (e: any) {
      if (e.message?.includes('GeneralException'))
        throw new Error('Invalid OOXML. Ensure the XML follows the Office Open XML package structure (pkg:package with word/document.xml part).');
      throw e;
    }
    return { success: true };
  },

  async getParaOoxml(ctx, p) {
    if (p.index < 0) throw new Error('Index must be non-negative');
    const count = p.count || 1;
    if (count < 1) throw new Error('count must be at least 1');
    const paragraphs = ctx.document.body.paragraphs;
    paragraphs.load('text');
    await ctx.sync();
    if (p.index >= paragraphs.items.length)
      throw new Error(`Paragraph index out of range. Valid indices: 0-${paragraphs.items.length - 1} (document has ${paragraphs.items.length} paragraphs).`);
    if (p.index + count - 1 >= paragraphs.items.length)
      throw new Error(`index + count (${p.index + count}) exceeds paragraph count (${paragraphs.items.length}).`);
    let range: any;
    if (count === 1) {
      range = paragraphs.items[p.index].getRange('Whole');
    } else {
      const firstRange = paragraphs.items[p.index].getRange('Whole');
      const lastRange = paragraphs.items[p.index + count - 1].getRange('Whole');
      range = firstRange.expandTo(lastRange);
    }
    const ooxml = range.getOoxml();
    await ctx.sync();
    return { ooxml: ooxml.value };
  },

  async insertOoxmlAtIndex(ctx, p) {
    if (p.index < 0) throw new Error('Index must be non-negative');
    if (!p.ooxml || typeof p.ooxml !== 'string' || p.ooxml.trim() === '')
      throw new Error('ooxml parameter must be a non-empty string');
    const paragraphs = ctx.document.body.paragraphs;
    paragraphs.load('text');
    await ctx.sync();
    if (p.index >= paragraphs.items.length)
      throw new Error(`Paragraph index out of range. Valid indices: 0-${paragraphs.items.length - 1} (document has ${paragraphs.items.length} paragraphs).`);

    const ref = paragraphs.items[p.index];
    const location = p.location === 'Before' ? Word.InsertLocation.before : Word.InsertLocation.after;
    const range = ref.getRange('Whole');
    range.insertOoxml(p.ooxml, location);
    await ctx.sync();
    return { success: true };
  },

  async insertOoxmlAfterMatch(ctx, p) {
    if (!p.ooxml || typeof p.ooxml !== 'string' || p.ooxml.trim() === '')
      throw new Error('ooxml parameter must be a non-empty string');
    if (!p.anchorText || typeof p.anchorText !== 'string' || p.anchorText.trim() === '')
      throw new Error('anchorText must be a non-empty string');
    const results = ctx.document.body.search(p.anchorText, { matchCase: p.matchCase || false });
    results.load('text');
    await ctx.sync();
    if (results.items.length === 0) throw new Error('Anchor not found: ' + p.anchorText);
    const idx = p.occurrence || 0;
    if (idx >= results.items.length)
      throw new Error(`Occurrence ${idx} not found (only ${results.items.length} match${results.items.length === 1 ? '' : 'es'})`);
    const target = results.items[idx];
    // Insert a space after the anchor, then position cursor there for OOXML insertion
    const endRange = target.getRange('End');
    const spaceRange = endRange.insertText(' ', Word.InsertLocation.after);
    await ctx.sync();
    // Select end of space range, then insert OOXML at cursor
    spaceRange.getRange('End').select();
    await ctx.sync();
    const sel = ctx.document.getSelection();
    sel.insertOoxml(p.ooxml, Word.InsertLocation.end);
    sel.getRange('End').select();
    try {
      await ctx.sync();
    } catch (e: any) {
      if (e.message?.includes('GeneralException'))
        throw new Error('Invalid OOXML for inline equation insertion.');
      throw e;
    }
    return { success: true };
  },
};
