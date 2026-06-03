import { describe, test, expect, beforeAll, afterAll } from 'vitest';
import { getClient, skip, cleanup } from './setup';
import type { LiveTestClient } from './client';

let client: LiveTestClient;
beforeAll(async () => { client = await getClient(); }, 60000);
afterAll(cleanup);

describe('Change Tracking', () => {
  beforeAll(async () => {
    if (skip()) return;
    await client.resetDocument();
  });

  test('enable tracking and accept change', async () => {
    if (skip()) return;
    await client.call('word_set_change_tracking', { mode: 'TrackAll' });
    await client.call('word_insert_paragraph', { text: 'Tracked addition' });
    const changes = await client.call('word_get_tracked_changes');
    expect(changes.count).toBeGreaterThan(0);
    const r = await client.call('word_accept_tracked_change', { index: 0 });
    expect(r.success).toBe(true);
    await client.call('word_set_change_tracking', { mode: 'Off' });
  });
});

describe('Content Controls & Properties', () => {
  beforeAll(async () => {
    if (skip()) return;
    await client.resetDocument();
    await client.call('word_insert_paragraph', { text: 'Editable field here' });
    await client.call('word_insert_content_control', { anchorText: 'field', type: 'RichText', tag: 'cc1', title: 'Field1' });
    await client.call('word_set_custom_property', { key: 'version', value: '1.0' });
  });

  test('content control created', async () => {
    if (skip()) return;
    const ccs = await client.call('word_get_content_controls');
    expect(ccs.count).toBe(1);
    expect(ccs.controls[0].tag).toBe('cc1');
  });

  test('set content control text', async () => {
    if (skip()) return;
    await client.call('word_set_content_control_text', { tag: 'cc1', text: 'New Value' });
    const ccs = await client.call('word_get_content_controls');
    expect(ccs.controls[0].text).toBe('New Value');
  });

  test('custom properties', async () => {
    if (skip()) return;
    const props = await client.call('word_get_custom_properties');
    expect(props.properties.find((p: any) => p.key === 'version').value).toBe('1.0');
    await client.call('word_delete_custom_property', { key: 'version' });
    const after = await client.call('word_get_custom_properties');
    expect(after.properties.find((p: any) => p.key === 'version')).toBeUndefined();
  });
});
