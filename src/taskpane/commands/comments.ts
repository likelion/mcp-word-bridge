import type { CommandHandler } from './index';
import { checkOccurrence, checkNonEmptyString, checkAnchorText, sanitizeText } from './document';

export const commentCommands: Record<string, CommandHandler> = {
  async addComment(ctx, p) {
    checkOccurrence(p.occurrence);
    checkNonEmptyString(p.comment, 'comment');
    checkAnchorText(p.anchorText);
    const results = ctx.document.body.search(p.anchorText, { matchCase: p.matchCase || false });
    results.load('text');
    await ctx.sync();
    if (results.items.length === 0) throw new Error('Anchor not found: ' + p.anchorText);
    const idx = p.occurrence || 0;
    if (idx >= results.items.length)
      throw new Error(`Occurrence ${idx} not found (only ${results.items.length} match${results.items.length === 1 ? '' : 'es'})`);
    const target = results.items[idx];
    target.insertComment(p.comment);
    await ctx.sync();
    return { success: true };
  },

  async getComments(ctx) {
    const comments = ctx.document.body.getComments();
    comments.load('id,authorName,content,creationDate,resolved');
    await ctx.sync();
    const items = comments.items.map((c: any) => ({ id: c.id, author: c.authorName, content: c.content, date: c.creationDate, resolved: c.resolved }));
    return { count: items.length, comments: items };
  },

  async getCommentsWithAnchor(ctx) {
    const comments = ctx.document.body.getComments();
    comments.load('id,authorName,content,creationDate,resolved');
    await ctx.sync();
    const items: any[] = [];
    for (const c of comments.items) {
      const range = c.getRange();
      range.load('text');
      await ctx.sync();
      // Sanitize: strip control chars and truncate long anchor text
      const rawText = (range.text || '').replace(/[\u0001-\u0008\u000B\u000C\u000E-\u001F\u0013\u0014\u0015]/g, '');
      const anchorText = rawText.length > 200 ? rawText.substring(0, 200) + '...' : rawText;
      items.push({ id: c.id, author: c.authorName, content: c.content, date: c.creationDate, resolved: c.resolved, anchorText });
    }
    return { count: items.length, comments: items };
  },

  async replyToComment(ctx, p) {
    const comments = ctx.document.body.getComments();
    comments.load('id');
    await ctx.sync();
    let target: any = null;
    for (const c of comments.items) {
      if (String(c.id) === String(p.commentId)) { target = c; break; }
    }
    if (!target) throw new Error('Comment not found: ' + p.commentId);
    target.reply(p.text);
    await ctx.sync();
    return { success: true };
  },

  async resolveComment(ctx, p) {
    const comments = ctx.document.body.getComments();
    comments.load('id');
    await ctx.sync();
    for (const c of comments.items) {
      if (String(c.id) === String(p.commentId)) {
        c.resolved = true;
        await ctx.sync();
        return { success: true };
      }
    }
    throw new Error('Comment not found: ' + p.commentId);
  },

  async deleteComment(ctx, p) {
    const comments = ctx.document.body.getComments();
    comments.load('id');
    await ctx.sync();
    for (const c of comments.items) {
      if (String(c.id) === String(p.commentId)) {
        c.delete();
        await ctx.sync();
        return { success: true };
      }
    }
    throw new Error('Comment not found: ' + p.commentId);
  },

  async getCommentReplies(ctx, p) {
    const comments = ctx.document.body.getComments();
    comments.load('id,replies');
    await ctx.sync();
    for (const c of comments.items) {
      if (String(c.id) === String(p.commentId)) {
        const replies = c.replies;
        replies.load('authorName,content,creationDate');
        await ctx.sync();
        const items = replies.items.map((r: any) => ({ author: r.authorName, content: r.content, date: r.creationDate }));
        return { count: items.length, replies: items };
      }
    }
    throw new Error('Comment not found: ' + p.commentId);
  },

  async getCommentAnchor(ctx, p) {
    const comments = ctx.document.body.getComments();
    comments.load('id');
    await ctx.sync();
    let target: any = null;
    for (const c of comments.items) {
      if (String(c.id) === String(p.commentId)) { target = c; break; }
    }
    if (!target) throw new Error('Comment not found: ' + p.commentId);
    const range = target.getRange();
    range.load('text');
    await ctx.sync();
    return { commentId: p.commentId, anchorText: range.text };
  },
};
