import type { CommandHandler } from './index';
import { checkSearchLength, sanitizeText } from './document';

declare const Word: any;

export const contentControlCommands: Record<string, CommandHandler> = {
  async getContentControls(ctx) {
    const ccs = ctx.document.body.getContentControls({ types: [Word.ContentControlType.richText, Word.ContentControlType.plainText, Word.ContentControlType.checkBox] });
    ccs.load('id,tag,title,type,text');
    await ctx.sync();
    return { count: ccs.items.length, controls: ccs.items.map((c: any) => ({ id: c.id, tag: c.tag, title: c.title, type: c.type, text: sanitizeText(c.text) })) };
  },

  async insertContentControl(ctx, p) {
    const ccType = p.type || 'RichText';
    let range: any;
    if (p.anchorText) {
      if (typeof p.anchorText === 'string' && p.anchorText.trim() === '')
        throw new Error('anchorText cannot be empty. Provide a non-empty search string or omit the parameter to use the current selection.');
      checkSearchLength(p.anchorText);
      const results = ctx.document.body.search(p.anchorText, { matchCase: p.matchCase || false });
      results.load('text');
      await ctx.sync();
      if (results.items.length === 0) throw new Error('Anchor not found: ' + p.anchorText);
      const idx = p.occurrence || 0;
      if (idx >= results.items.length)
        throw new Error(`Occurrence ${idx} not found (only ${results.items.length} match${results.items.length === 1 ? '' : 'es'})`);
      range = results.items[idx];
    } else {
      range = ctx.document.getSelection();
    }
    const cc = range.insertContentControl(ccType);
    if (p.title) cc.title = p.title;
    if (p.tag) cc.tag = p.tag;
    if (p.color) cc.color = p.color;
    cc.getRange('End').select();
    await ctx.sync();
    const result: any = { success: true };
    if (ccType === 'CheckBox') {
      result.warning = 'CheckBox content controls replace the anchor text with a checkbox widget. The original text is not preserved.';
    }
    return result;
  },

  async setContentControlText(ctx, p) {
    let target: any = null;
    if (p.tag) {
      const ccs = ctx.document.body.getContentControls({ types: [Word.ContentControlType.richText, Word.ContentControlType.plainText, Word.ContentControlType.checkBox] });
      ccs.load('id,tag,type');
      await ctx.sync();
      for (const cc of ccs.items) {
        if (cc.tag === p.tag) {
          if (cc.type === 'CheckBox')
            throw new Error('Cannot set text on a CheckBox content control. CheckBox controls only support checked/unchecked state.');
          target = cc;
          break;
        }
      }
    }
    if (!target && p.id) {
      const allCcs = ctx.document.body.getContentControls({ types: [Word.ContentControlType.richText, Word.ContentControlType.plainText, Word.ContentControlType.checkBox] });
      allCcs.load('id,type');
      await ctx.sync();
      for (const cc of allCcs.items) {
        if (cc.id === p.id) {
          if (cc.type === 'CheckBox')
            throw new Error('Cannot set text on a CheckBox content control. CheckBox controls only support checked/unchecked state.');
          target = cc;
          break;
        }
      }
    }
    if (!target) throw new Error('Content control not found. Provide "id" or "tag" to identify the control. Use word_get_content_controls to list available controls.');
    const inserted = target.insertText(p.text, Word.InsertLocation.replace);
    inserted.getRange('End').select();
    await ctx.sync();
    return { success: true };
  },
};
