import type { CommandHandler } from './index';

export const imageCommands: Record<string, CommandHandler> = {
  async insertImage(ctx, p) {
    if (!p.base64 || typeof p.base64 !== 'string' || p.base64.trim() === '')
      throw new Error('Invalid image data. Ensure the base64 string is a valid PNG, JPEG, or GIF image.');
    const body = ctx.document.body;
    const picture = body.insertInlinePictureFromBase64(p.base64, p.location || 'End');
    if (p.width) picture.width = p.width;
    if (p.height) picture.height = p.height;
    if (p.altText) picture.altTextDescription = p.altText;
    try {
      await ctx.sync();
    } catch (e: any) {
      if (e.message?.includes('GeneralException'))
        throw new Error('Invalid image data. Ensure the base64 string is a valid PNG, JPEG, or GIF image.');
      throw e;
    }
    const imgRange = picture.getRange('End');
    imgRange.select();
    await ctx.sync();
    return { success: true };
  },

  async getImages(ctx) {
    const pics = ctx.document.body.inlinePictures;
    pics.load('altTextDescription,altTextTitle,width,height,hyperlink');
    await ctx.sync();
    const items = pics.items.map((pic: any, i: number) => ({
      index: i, width: pic.width, height: pic.height, altText: pic.altTextDescription, altTitle: pic.altTextTitle, hyperlink: pic.hyperlink,
    }));
    return { count: items.length, images: items };
  },

  async deleteImage(ctx, p) {
    if (p.index < 0) throw new Error('Index must be non-negative');
    const pics = ctx.document.body.inlinePictures;
    pics.load('altTextDescription');
    await ctx.sync();
    if (p.index >= pics.items.length)
      throw new Error(`Image index out of range. Document has ${pics.items.length} image(s) (0-indexed).`);
    pics.items[p.index].delete();
    await ctx.sync();
    return { success: true };
  },
};
