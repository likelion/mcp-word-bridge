import type { CommandHandler } from './index';
import { checkSearchLength } from './document';

export const commentCommands: Record<string, CommandHandler> = {
  async addComment(ctx, p) {
    if (p.occurrence !== undefined && p.occurrence < 0) throw new Error('occurrence must be non-negative (0-indexed)');
    if (!p.comment || typeof p.comment !== 'string' || p.comment.trim() === '')
      throw new Error('comment text must be a non-empty string');
    if (!p.anchorText || typeof p.anchorText !== 'string' || p.anchorText.trim() === '')
      throw new Error('anchorText cannot be empty. Provide a non-empty search string.');
    checkSearchLength(p.anchorText);
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
      items.push({ id: c.id, author: c.authorName, content: c.content, date: c.creationDate, resolved: c.resolved, anchorText: range.text });
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
