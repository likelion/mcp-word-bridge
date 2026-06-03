import type { ToolDefinition } from '../types';
import { forwardTool } from './helpers';

export const addComment = forwardTool(
  'word_add_comment',
  '[Comments] Add a review comment anchored to a text match.',
  {
    properties: {
      anchorText: { type: 'string', description: 'Text to search for as comment anchor' },
      comment: { type: 'string', description: 'Comment text' },
      occurrence: { type: 'number', description: '0=first, 1=second, etc. Default: 0' },
      matchCase: { type: 'boolean', description: 'Default: false' },
    },
    required: ['anchorText', 'comment'],
  },
  'addComment',
);

export const getComments = forwardTool(
  'word_get_comments',
  '[Comments] Get all comments with ID, author, content, date, resolved status, and anchor text.',
  { properties: {} },
  'getCommentsWithAnchor',
);

export const getCommentReplies = forwardTool(
  'word_get_comment_replies',
  '[Comments] Get all replies for a specific comment by its ID.',
  {
    properties: {
      commentId: { type: 'string' },
    },
    required: ['commentId'],
  },
  'getCommentReplies',
);

export const replyToComment = forwardTool(
  'word_reply_to_comment',
  '[Comments] Reply to a comment thread by its ID.',
  {
    properties: {
      commentId: { type: 'string' },
      text: { type: 'string' },
    },
    required: ['commentId', 'text'],
  },
  'replyToComment',
);

export const resolveComment = forwardTool(
  'word_resolve_comment',
  '[Comments] Mark a comment as resolved. Preferred over delete to preserve audit trail.',
  {
    properties: {
      commentId: { type: 'string' },
    },
    required: ['commentId'],
  },
  'resolveComment',
);

export const deleteComment = forwardTool(
  'word_delete_comment',
  '[Comments] Permanently delete a comment and all its replies.',
  {
    properties: {
      commentId: { type: 'string' },
    },
    required: ['commentId'],
  },
  'deleteComment',
);

export const commentTools: ToolDefinition[] = [
  addComment,
  getComments,
  getCommentReplies,
  replyToComment,
  resolveComment,
  deleteComment,
];
