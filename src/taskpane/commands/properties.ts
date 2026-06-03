import type { CommandHandler } from './index';

export const propertyCommands: Record<string, CommandHandler> = {
  async getCustomProperties(ctx) {
    const customProps = ctx.document.properties.customProperties;
    customProps.load('key,value,type');
    await ctx.sync();
    const items = customProps.items.map((cp: any) => ({ key: cp.key, value: cp.value, type: cp.type }));
    return { count: items.length, properties: items };
  },

  async setCustomProperty(ctx, p) {
    if (!p.key || typeof p.key !== 'string' || p.key.trim() === '')
      throw new Error('key must be a non-empty string');
    // BUG-09: Validate key length to prevent silent truncation
    if (p.key.length > 255)
      throw new Error(`key must be 255 characters or fewer (got ${p.key.length}).`);
    if (p.value === undefined || p.value === null || (typeof p.value === 'string' && p.value.trim() === ''))
      throw new Error('value must be a non-empty string');
    ctx.document.properties.customProperties.add(p.key, p.value);
    await ctx.sync();
    return { success: true };
  },

  async deleteCustomProperty(ctx, p) {
    const customProps = ctx.document.properties.customProperties;
    customProps.load('key');
    await ctx.sync();
    for (const cp of customProps.items) {
      if (cp.key === p.key) {
        cp.delete();
        await ctx.sync();
        return { success: true };
      }
    }
    throw new Error('Custom property not found: ' + p.key);
  },
};
