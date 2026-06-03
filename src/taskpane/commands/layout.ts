import type { CommandHandler } from './index';

declare const Office: any;

export const layoutCommands: Record<string, CommandHandler> = {
  async getPageLayout(ctx, p) {
    const sections = ctx.document.sections;
    sections.load('items');
    await ctx.sync();
    const idx = p.sectionIndex || 0;
    if (idx < 0 || idx >= sections.items.length)
      throw new Error(`Section index out of range. Document has ${sections.items.length} section(s) (0-indexed).`);
    const section = sections.items[idx];
    const ps = section.pageSetup;
    ps.load('orientation,topMargin,bottomMargin,leftMargin,rightMargin,paperSize,headerDistance,footerDistance');
    await ctx.sync();
    return { orientation: ps.orientation, topMargin: ps.topMargin, bottomMargin: ps.bottomMargin, leftMargin: ps.leftMargin, rightMargin: ps.rightMargin, paperSize: ps.paperSize, headerDistance: ps.headerDistance, footerDistance: ps.footerDistance };
  },

  async setPageLayout(ctx, p) {
    if (p.orientation !== undefined) {
      const validOrientations = ['Portrait', 'Landscape'];
      if (!validOrientations.includes(p.orientation))
        throw new Error(`Invalid orientation: "${p.orientation}". Valid values: Portrait, Landscape.`);
    }
    if (p.topMargin !== undefined && p.topMargin < 0) throw new Error('topMargin must be non-negative (in points)');
    if (p.bottomMargin !== undefined && p.bottomMargin < 0) throw new Error('bottomMargin must be non-negative (in points)');
    if (p.leftMargin !== undefined && p.leftMargin < 0) throw new Error('leftMargin must be non-negative (in points)');
    if (p.rightMargin !== undefined && p.rightMargin < 0) throw new Error('rightMargin must be non-negative (in points)');
    const sections = ctx.document.sections;
    sections.load('items');
    await ctx.sync();
    const idx = p.sectionIndex || 0;
    if (idx < 0 || idx >= sections.items.length)
      throw new Error(`Section index out of range. Document has ${sections.items.length} section(s) (0-indexed).`);
    const section = sections.items[idx];
    const pageSetup = section.pageSetup;
    if (p.orientation) pageSetup.orientation = p.orientation;
    if (p.topMargin !== undefined) pageSetup.topMargin = p.topMargin;
    if (p.bottomMargin !== undefined) pageSetup.bottomMargin = p.bottomMargin;
    if (p.leftMargin !== undefined) pageSetup.leftMargin = p.leftMargin;
    if (p.rightMargin !== undefined) pageSetup.rightMargin = p.rightMargin;
    if (p.paperSize) pageSetup.paperSize = p.paperSize;
    await ctx.sync();
    return { success: true };
  },

  async getSections(ctx) {
    const sections = ctx.document.sections;
    sections.load('items');
    await ctx.sync();
    const items: any[] = [];
    for (let i = 0; i < sections.items.length; i++) {
      const ps = sections.items[i].pageSetup;
      ps.load('orientation,topMargin,bottomMargin,leftMargin,rightMargin,paperSize');
      await ctx.sync();
      items.push({ index: i, orientation: ps.orientation, topMargin: ps.topMargin, bottomMargin: ps.bottomMargin, leftMargin: ps.leftMargin, rightMargin: ps.rightMargin, paperSize: ps.paperSize });
    }
    return { count: items.length, sections: items };
  },

  async insertPageBreak(ctx, p) {
    const body = ctx.document.body;
    if (p.paragraphIndex !== undefined) {
      if (typeof p.paragraphIndex !== 'number' || !Number.isInteger(p.paragraphIndex) || p.paragraphIndex < 0) throw new Error('paragraphIndex must be a non-negative integer.');
      const paragraphs = body.paragraphs;
      paragraphs.load('text');
      await ctx.sync();
      if (p.paragraphIndex >= paragraphs.items.length)
        throw new Error(`Paragraph index out of range. Valid indices: 0-${paragraphs.items.length - 1} (document has ${paragraphs.items.length} paragraphs).`);
      const para = paragraphs.items[p.paragraphIndex];
      para.insertBreak('Page', 'After');
      para.getRange('End').select();
    } else {
      const lastPara = body.paragraphs.getLast();
      lastPara.insertBreak('Page', 'After');
      lastPara.getRange('End').select();
    }
    try {
      await ctx.sync();
    } catch (e: any) {
      if (e.message?.includes('GeneralException'))
        throw new Error(`Cannot insert page break at paragraph ${p.paragraphIndex !== undefined ? p.paragraphIndex : 'end'}. The paragraph may be inside a table cell (page breaks are not allowed inside tables).`);
      throw e;
    }
    return { success: true };
  },

  async insertSectionBreak(ctx, p) {
    const body = ctx.document.body;
    const breakType = p.breakType || 'SectionNext';
    const validBreakTypes = ['SectionNext', 'SectionContinuous', 'SectionEven', 'SectionOdd'];
    if (!validBreakTypes.includes(breakType))
      throw new Error(`Invalid breakType: "${breakType}". Valid values: ${validBreakTypes.join(', ')}.`);
    if (p.paragraphIndex !== undefined) {
      if (typeof p.paragraphIndex !== 'number' || !Number.isInteger(p.paragraphIndex) || p.paragraphIndex < 0) throw new Error('paragraphIndex must be a non-negative integer.');
      const paragraphs = body.paragraphs;
      paragraphs.load('text');
      await ctx.sync();
      if (p.paragraphIndex >= paragraphs.items.length)
        throw new Error(`Paragraph index out of range. Valid indices: 0-${paragraphs.items.length - 1} (document has ${paragraphs.items.length} paragraphs).`);
      const para = paragraphs.items[p.paragraphIndex];
      para.insertBreak(breakType, 'After');
      para.getRange('End').select();
      try {
        await ctx.sync();
      } catch (e: any) {
        if (e.message?.includes('GeneralException'))
          throw new Error(`Cannot insert section break at paragraph ${p.paragraphIndex}. The paragraph may be inside a table cell (section/page breaks are not allowed inside tables).`);
        throw e;
      }
    } else {
      const lastPara = body.paragraphs.getLast();
      lastPara.insertBreak(breakType, 'After');
      lastPara.getRange('End').select();
      await ctx.sync();
    }
    return { success: true };
  },

  async getPageInfo(ctx) {
    if (!Office.context.requirements.isSetSupported('WordApiDesktop', '1.2'))
      throw new Error('Page API not available on this platform (requires desktop Word).');
    const bodyRange = ctx.document.body.getRange();
    const pages = bodyRange.pages;
    pages.load('items');
    await ctx.sync();
    const allParas = ctx.document.body.paragraphs;
    allParas.load('text,style');
    await ctx.sync();
    const pageDetails: any[] = [];
    let nextStart = 0;
    for (let i = 0; i < pages.items.length; i++) {
      const page = pages.items[i];
      page.load('index,height,width');
      const paras = page.getRange().paragraphs;
      paras.load('text,style');
      await ctx.sync();
      let firstIdx = -1;
      let lastIdx = -1;
      for (let j = 0; j < paras.items.length; j++) {
        const pp = paras.items[j];
        for (let k = nextStart; k < allParas.items.length; k++) {
          if (allParas.items[k].text === pp.text && allParas.items[k].style === pp.style) {
            if (firstIdx === -1) firstIdx = k;
            lastIdx = k;
            nextStart = k + 1;
            break;
          }
        }
      }
      pageDetails.push({ pageIndex: page.index, height: page.height, width: page.width, paragraphCount: paras.items.length, firstParagraphIndex: firstIdx, lastParagraphIndex: lastIdx });
    }
    return { pageCount: pages.items.length, pages: pageDetails };
  },

  async insertTableOfContents(ctx, p) {
    const body = ctx.document.body;
    const range = body.getRange(p.location || 'Start');
    range.insertField(p.location || 'Start', 'TOC', p.switches || '\\o "1-3" \\h \\z \\u', true);
    await ctx.sync();
    return { success: true };
  },
};
