import { describe, test, expect, beforeAll, afterAll } from 'vitest';
import { getClient, skip, cleanup } from './setup';
import type { LiveTestClient } from './client';

let client: LiveTestClient;
beforeAll(async () => { client = await getClient(); }, 60000);
afterAll(cleanup);

describe('Comments', () => {
  beforeAll(async () => {
    if (skip()) return;
    await client.resetDocument();
    await client.call('word_insert_paragraph', { text: 'First section to review' });
    await client.call('word_insert_paragraph', { text: 'Second important paragraph' });
    await client.call('word_add_comment', { anchorText: 'First section', comment: 'Comment A' });
    await client.call('word_add_comment', { anchorText: 'important', comment: 'Comment B' });
  });

  test('comments created with anchor text', async () => {
    if (skip()) return;
    const c = await client.call('word_get_comments');
    expect(c.count).toBe(2);
    expect(c.comments[0].anchorText).toBe('First section');
  });

  test('reply to comment', async () => {
    if (skip()) return;
    const c = await client.call('word_get_comments');
    await client.call('word_reply_to_comment', { commentId: c.comments[0].id, text: 'Reply' });
    const replies = await client.call('word_get_comment_replies', { commentId: c.comments[0].id });
    expect(replies.count).toBe(1);
  });

  test('resolve and delete comment', async () => {
    if (skip()) return;
    const c = await client.call('word_get_comments');
    await client.call('word_resolve_comment', { commentId: c.comments[1].id });
    const updated = await client.call('word_get_comments');
    expect(updated.comments.find((x: any) => x.id === c.comments[1].id).resolved).toBe(true);
    await client.call('word_delete_comment', { commentId: c.comments[0].id });
    const after = await client.call('word_get_comments');
    expect(after.count).toBe(1);
  });
});
