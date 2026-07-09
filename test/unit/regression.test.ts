/**
 * Regression tests for input validation, error handling, and edge cases.
 */
import { describe, test, expect } from 'vitest';
import { buildToolRegistry } from '../../src/server/tools';
import {
  checkNonNegative,
  normalizeAlignment,
  checkNoSpecialCodes,
  checkSpacingBounds,
  checkPropertyKeyLength,
  MAX_SPACING_POINTS,
  MAX_CUSTOM_PROPERTY_KEY_LENGTH,
} from '../../src/server/validation';

const { handlers } = buildToolRegistry();

// --- Helpers ---

/** Mock bridge that tracks sent actions and supports custom responses */
function mockBridge(responses: Record<string, any> = {}): any {
  const calls: Array<{ action: string; params: any }> = [];
  return {
    calls,
    send: async (action: string, params: any = {}) => {
      calls.push({ action, params });
      if (responses[action]) return responses[action];
      if (action === 'getParagraphs') return { total: 5, count: 5, paragraphs: [] };
      if (action === 'getParaOoxml') return { ooxml: '<pkg:package></pkg:package>' };
      if (action === 'getContentControls') return { count: 0, controls: [] };
      return { success: true };
    },
  };
}

/** Shorthand mock with just paragraph count */
function mockBridgeWithParas(count = 5): any {
  return mockBridge({ getParagraphs: { total: count, count, paragraphs: [] } });
}

// =============================================================================
// search_and_replace allows Word special codes (consistent with word_search)
// =============================================================================
describe('search_and_replace allows Word special codes', () => {
  const handler = handlers.get('word_search_and_replace')!;

  test('accepts ^p in find string', async () => {
    const bridge = mockBridge({ searchAndReplace: { replacements: 1 } });
    const result = await handler({ find: '^p', replace: 'x' }, bridge);
    const parsed = JSON.parse(result.content[0].text);
    expect(parsed.replacements).toBe(1);
  });

  test('accepts ^13 in find string', async () => {
    const bridge = mockBridge({ searchAndReplace: { replacements: 1 } });
    const result = await handler({ find: '^13', replace: 'x' }, bridge);
    const parsed = JSON.parse(result.content[0].text);
    expect(parsed.replacements).toBe(1);
  });

  test('accepts ^w in find string', async () => {
    const bridge = mockBridge({ searchAndReplace: { replacements: 1 } });
    const result = await handler({ find: '^w', replace: 'x' }, bridge);
    const parsed = JSON.parse(result.content[0].text);
    expect(parsed.replacements).toBe(1);
  });

  test('accepts ^t in find string', async () => {
    const bridge = mockBridge({ searchAndReplace: { replacements: 1 } });
    const result = await handler({ find: '^t', replace: 'x' }, bridge);
    const parsed = JSON.parse(result.content[0].text);
    expect(parsed.replacements).toBe(1);
  });

  test('accepts ^11 in find string', async () => {
    const bridge = mockBridge({ searchAndReplace: { replacements: 1 } });
    const result = await handler({ find: '^11', replace: 'x' }, bridge);
    const parsed = JSON.parse(result.content[0].text);
    expect(parsed.replacements).toBe(1);
  });

  test('accepts ^p in replace string', async () => {
    const bridge = mockBridge({ searchAndReplace: { replacements: 1 } });
    const result = await handler({ find: 'hello', replace: '^p' }, bridge);
    const parsed = JSON.parse(result.content[0].text);
    expect(parsed.replacements).toBe(1);
  });

  test('accepts ^w in replace string', async () => {
    const bridge = mockBridge({ searchAndReplace: { replacements: 1 } });
    const result = await handler({ find: 'hello', replace: 'a^wb' }, bridge);
    const parsed = JSON.parse(result.content[0].text);
    expect(parsed.replacements).toBe(1);
  });

  test('accepts normal text without caret codes', async () => {
    const bridge = mockBridge({ searchAndReplace: { replacements: 2 } });
    const result = await handler({ find: 'hello', replace: 'world' }, bridge);
    const parsed = JSON.parse(result.content[0].text);
    expect(parsed.replacements).toBe(2);
  });

  test('accepts caret character not followed by special code', async () => {
    const bridge = mockBridge({ searchAndReplace: { replacements: 1 } });
    const result = await handler({ find: 'x^y', replace: 'z' }, bridge);
    const parsed = JSON.parse(result.content[0].text);
    expect(parsed.replacements).toBe(1);
  });

  test('rejects empty find string', async () => {
    await expect(handler({ find: '', replace: 'x' }, mockBridge()))
      .rejects.toThrow('cannot be empty');
  });

  test('rejects whitespace-only find string', async () => {
    await expect(handler({ find: '   ', replace: 'x' }, mockBridge()))
      .rejects.toThrow('cannot be empty');
  });
});

// =============================================================================
// search_and_replace bookmark preservation
// =============================================================================
describe('search_and_replace bookmark preservation', () => {
  const handler = handlers.get('word_search_and_replace')!;

  test('passes preserveBookmarks parameter to bridge', async () => {
    const bridge = mockBridge({ searchAndReplace: { replacements: 1 } });
    await handler({ find: 'hello', replace: 'world', preserveBookmarks: true }, bridge);
    expect(bridge.calls[0].params.preserveBookmarks).toBe(true);
  });

  test('returns bookmarksLost in response when bridge reports it', async () => {
    const bridge = mockBridge({
      searchAndReplace: { replacements: 2, bookmarksLost: 1, warning: '1 bookmark(s) destroyed by this replacement: bm1. Use preserveBookmarks: true to re-create them on the replacement text.' },
    });
    const result = await handler({ find: 'foo', replace: 'bar' }, bridge);
    const parsed = JSON.parse(result.content[0].text);
    expect(parsed.replacements).toBe(2);
    expect(parsed.bookmarksLost).toBe(1);
    expect(parsed.warning).toContain('bookmark');
  });

  test('returns bookmarksRestored in response when preserveBookmarks is true', async () => {
    const bridge = mockBridge({
      searchAndReplace: { replacements: 3, bookmarksRestored: 2 },
    });
    const result = await handler({ find: 'a', replace: 'b', preserveBookmarks: true }, bridge);
    const parsed = JSON.parse(result.content[0].text);
    expect(parsed.replacements).toBe(3);
    expect(parsed.bookmarksRestored).toBe(2);
    expect(parsed.warning).toBeUndefined();
  });

  test('no bookmark fields when no bookmarks affected', async () => {
    const bridge = mockBridge({ searchAndReplace: { replacements: 5 } });
    const result = await handler({ find: 'x', replace: 'y' }, bridge);
    const parsed = JSON.parse(result.content[0].text);
    expect(parsed.replacements).toBe(5);
    expect(parsed.bookmarksLost).toBeUndefined();
    expect(parsed.bookmarksRestored).toBeUndefined();
    expect(parsed.warning).toBeUndefined();
  });

  test('preserveBookmarks defaults to false (not included in params)', async () => {
    const bridge = mockBridge({ searchAndReplace: { replacements: 1 } });
    await handler({ find: 'a', replace: 'b' }, bridge);
    expect(bridge.calls[0].params.preserveBookmarks).toBeUndefined();
  });
});

// =============================================================================
// search_and_replace case-preservation skip
// =============================================================================
describe('search_and_replace case-preservation', () => {
  const handler = handlers.get('word_search_and_replace')!;

  test('returns skipped count when matches are identical to replace text', async () => {
    const bridge = mockBridge({ searchAndReplace: { replacements: 1, skipped: 2 } });
    const result = await handler({ find: 'hello', replace: 'hello', matchCase: false }, bridge);
    const parsed = JSON.parse(result.content[0].text);
    expect(parsed.replacements).toBe(1);
    expect(parsed.skipped).toBe(2);
  });

  test('no skipped field when all matches are replaced', async () => {
    const bridge = mockBridge({ searchAndReplace: { replacements: 3 } });
    const result = await handler({ find: 'Hello', replace: 'World' }, bridge);
    const parsed = JSON.parse(result.content[0].text);
    expect(parsed.replacements).toBe(3);
    expect(parsed.skipped).toBeUndefined();
  });

  test('skipped field is zero when not present in response', async () => {
    const bridge = mockBridge({ searchAndReplace: { replacements: 0 } });
    const result = await handler({ find: 'nonexistent', replace: 'x' }, bridge);
    const parsed = JSON.parse(result.content[0].text);
    expect(parsed.replacements).toBe(0);
    expect(parsed.skipped).toBeUndefined();
  });
});

// =============================================================================
// checkNoSpecialCodes validation helper
// =============================================================================
describe('checkNoSpecialCodes', () => {
  test('rejects ^p', () => {
    expect(() => checkNoSpecialCodes('^p', 'find')).toThrow('Word special code "^p"');
  });

  test('rejects ^w', () => {
    expect(() => checkNoSpecialCodes('^w', 'find')).toThrow('Word special code "^w"');
  });

  test('rejects ^t', () => {
    expect(() => checkNoSpecialCodes('^t', 'find')).toThrow('Word special code "^t"');
  });

  test('rejects ^13', () => {
    expect(() => checkNoSpecialCodes('^13', 'find')).toThrow('Word special code "^13"');
  });

  test('rejects ^11', () => {
    expect(() => checkNoSpecialCodes('^11', 'find')).toThrow('Word special code "^11"');
  });

  test('rejects ^^ (escaped caret)', () => {
    expect(() => checkNoSpecialCodes('^^', 'find')).toThrow('Word special code');
  });

  test('rejects ^~ (non-breaking hyphen)', () => {
    expect(() => checkNoSpecialCodes('^~', 'find')).toThrow('Word special code');
  });

  test('accepts text without caret', () => {
    expect(() => checkNoSpecialCodes('normal text', 'find')).not.toThrow();
  });

  test('accepts caret not followed by special letter', () => {
    expect(() => checkNoSpecialCodes('x^y', 'find')).not.toThrow();
    expect(() => checkNoSpecialCodes('^z', 'find')).not.toThrow();
    expect(() => checkNoSpecialCodes('^', 'find')).not.toThrow();
  });

  test('accepts caret followed by non-special number', () => {
    expect(() => checkNoSpecialCodes('^2', 'find')).not.toThrow();
    expect(() => checkNoSpecialCodes('^99', 'find')).not.toThrow();
  });
});

// =============================================================================
// set_content_control_text rejects duplicate tags
// =============================================================================
describe('set_content_control_text duplicate tag detection', () => {
  const handler = handlers.get('word_set_content_control_text')!;

  test('rejects when multiple controls share same tag', async () => {
    const bridge = mockBridge({
      getContentControls: {
        count: 2,
        controls: [
          { id: 100, tag: 'cc1', title: 'A', type: 'RichText', text: 'a' },
          { id: 200, tag: 'cc1', title: 'B', type: 'RichText', text: 'b' },
        ],
      },
    });
    await expect(handler({ tag: 'cc1', text: 'new' }, bridge))
      .rejects.toThrow('Multiple content controls (2) share tag "cc1"');
  });

  test('includes matching IDs in error message', async () => {
    const bridge = mockBridge({
      getContentControls: {
        count: 3,
        controls: [
          { id: 10, tag: 'dup', title: 'A', type: 'RichText', text: 'a' },
          { id: 20, tag: 'dup', title: 'B', type: 'RichText', text: 'b' },
          { id: 30, tag: 'other', title: 'C', type: 'RichText', text: 'c' },
        ],
      },
    });
    await expect(handler({ tag: 'dup', text: 'new' }, bridge))
      .rejects.toThrow('Matching IDs: 10, 20');
  });

  test('passes through when only one control matches tag', async () => {
    const bridge = mockBridge({
      getContentControls: {
        count: 2,
        controls: [
          { id: 100, tag: 'unique', title: 'A', type: 'RichText', text: 'a' },
          { id: 200, tag: 'other', title: 'B', type: 'RichText', text: 'b' },
        ],
      },
      setContentControlText: { success: true },
    });
    const result = await handler({ tag: 'unique', text: 'new' }, bridge);
    const parsed = JSON.parse(result.content[0].text);
    expect(parsed.success).toBe(true);
  });

  test('rejects when neither id nor tag is provided', async () => {
    await expect(handler({ text: 'new' }, mockBridge()))
      .rejects.toThrow('Provide "id" or "tag"');
  });

  test('rejects with specific message when tag has zero matches', async () => {
    const bridge = mockBridge({
      getContentControls: {
        count: 2,
        controls: [
          { id: 100, tag: 'existing', title: 'A', type: 'RichText', text: 'a' },
          { id: 200, tag: 'other', title: 'B', type: 'RichText', text: 'b' },
        ],
      },
    });
    await expect(handler({ tag: 'nonexistent', text: 'new' }, bridge))
      .rejects.toThrow('Content control with tag "nonexistent" not found');
  });

  test('skips duplicate check when using id directly', async () => {
    const bridge = mockBridge({ setContentControlText: { success: true } });
    const result = await handler({ id: 42, text: 'new' }, bridge);
    const parsed = JSON.parse(result.content[0].text);
    expect(parsed.success).toBe(true);
    expect(bridge.calls.find((c: any) => c.action === 'getContentControls')).toBeUndefined();
  });
});

// =============================================================================
// insert_content_control forwards tag parameter for duplicate detection
// =============================================================================
describe('insert_content_control tag forwarding', () => {
  const handler = handlers.get('word_insert_content_control')!;

  test('passes tag parameter to bridge', async () => {
    const bridge = mockBridge({ insertContentControl: { success: true } });
    await handler({ anchorText: 'hello', tag: 'my_tag', type: 'RichText' }, bridge);
    const call = bridge.calls.find((c: any) => c.action === 'insertContentControl');
    expect(call).toBeDefined();
    expect(call.params.tag).toBe('my_tag');
  });

  test('bridge response warning is returned to caller', async () => {
    const bridge = mockBridge({ insertContentControl: { success: true, warning: 'Another content control already uses tag "my_tag". Duplicate tags prevent tag-based lookups — use word_set_content_control_text with "id" instead.' } });
    const result = await handler({ anchorText: 'hello', tag: 'my_tag' }, bridge);
    const parsed = JSON.parse(result.content[0].text);
    expect(parsed.warning).toContain('Duplicate tags');
  });
});

// =============================================================================
// Float index and count rejection
// =============================================================================
describe('Float index and count rejection', () => {
  const moveHandler = handlers.get('word_move_paragraph')!;
  const copyHandler = handlers.get('word_copy_paragraph')!;

  test('checkNonNegative rejects 1.5', () => {
    expect(() => checkNonNegative(1.5, 'index')).toThrow('non-negative integer');
  });

  test('checkNonNegative rejects 0.1', () => {
    expect(() => checkNonNegative(0.1, 'index')).toThrow('non-negative integer');
  });

  test('checkNonNegative rejects 1.999', () => {
    expect(() => checkNonNegative(1.999, 'index')).toThrow('non-negative integer');
  });

  test('checkNonNegative rejects NaN', () => {
    expect(() => checkNonNegative(NaN, 'index')).toThrow('non-negative integer');
  });

  test('checkNonNegative rejects Infinity', () => {
    expect(() => checkNonNegative(Infinity, 'index')).toThrow('non-negative integer');
  });

  test('move_paragraph rejects float fromIndex', async () => {
    await expect(moveHandler({ fromIndex: 1.5, toIndex: 3 }, mockBridgeWithParas()))
      .rejects.toThrow('non-negative integer');
  });

  test('move_paragraph rejects float toIndex', async () => {
    await expect(moveHandler({ fromIndex: 0, toIndex: 2.5 }, mockBridgeWithParas()))
      .rejects.toThrow('non-negative integer');
  });

  test('move_paragraph rejects float count', async () => {
    await expect(moveHandler({ fromIndex: 0, toIndex: 3, count: 1.7 }, mockBridgeWithParas()))
      .rejects.toThrow('count must be an integer');
  });

  test('move_paragraph rejects NaN count', async () => {
    await expect(moveHandler({ fromIndex: 0, toIndex: 3, count: NaN }, mockBridgeWithParas()))
      .rejects.toThrow('count must be an integer');
  });

  test('copy_paragraph rejects float fromIndex', async () => {
    await expect(copyHandler({ fromIndex: 0.5, toIndex: 2 }, mockBridgeWithParas()))
      .rejects.toThrow('non-negative integer');
  });

  test('copy_paragraph rejects float toIndex', async () => {
    await expect(copyHandler({ fromIndex: 0, toIndex: 3.14 }, mockBridgeWithParas()))
      .rejects.toThrow('non-negative integer');
  });

  test('copy_paragraph rejects float count', async () => {
    await expect(copyHandler({ fromIndex: 0, toIndex: 2, count: 1.5 }, mockBridgeWithParas()))
      .rejects.toThrow('count must be an integer');
  });

  test('copy_paragraph accepts integer count', async () => {
    const bridge = mockBridge();
    const result = await copyHandler({ fromIndex: 0, toIndex: 2, count: 2 }, bridge);
    const parsed = JSON.parse(result.content[0].text);
    expect(parsed.success).toBe(true);
  });
});

// =============================================================================
// setParagraphSpacing bounds validation
// =============================================================================
describe('setParagraphSpacing bounds validation', () => {
  const handler = handlers.get('word_set_paragraph_spacing')!;

  test('rejects leftIndent exceeding maximum', async () => {
    await expect(handler({ index: 0, leftIndent: 99999 }, mockBridge()))
      .rejects.toThrow('exceeds maximum');
  });

  test('rejects rightIndent exceeding maximum', async () => {
    await expect(handler({ index: 0, rightIndent: 99999 }, mockBridge()))
      .rejects.toThrow('exceeds maximum');
  });

  test('rejects firstLineIndent exceeding maximum', async () => {
    await expect(handler({ index: 0, firstLineIndent: 99999 }, mockBridge()))
      .rejects.toThrow('out of range');
  });

  test('rejects lineSpacing exceeding maximum', async () => {
    await expect(handler({ index: 0, lineSpacing: 99999 }, mockBridge()))
      .rejects.toThrow('exceeds maximum');
  });

  test('rejects spaceBefore exceeding maximum', async () => {
    await expect(handler({ index: 0, spaceBefore: 99999 }, mockBridge()))
      .rejects.toThrow('exceeds maximum');
  });

  test('rejects spaceAfter exceeding maximum', async () => {
    await expect(handler({ index: 0, spaceAfter: 99999 }, mockBridge()))
      .rejects.toThrow('exceeds maximum');
  });

  test('accepts valid spacing at maximum', async () => {
    const bridge = mockBridge({ setParagraphSpacing: { success: true } });
    const result = await handler({ index: 0, leftIndent: MAX_SPACING_POINTS }, bridge);
    const parsed = JSON.parse(result.content[0].text);
    expect(parsed.success).toBe(true);
  });

  test('accepts normal spacing values', async () => {
    const bridge = mockBridge({ setParagraphSpacing: { success: true } });
    const result = await handler({ index: 0, leftIndent: 36, spaceBefore: 12 }, bridge);
    const parsed = JSON.parse(result.content[0].text);
    expect(parsed.success).toBe(true);
  });

  test('rejects non-integer index', async () => {
    await expect(handler({ index: 1.5, leftIndent: 36 }, mockBridge()))
      .rejects.toThrow('non-negative integer');
  });

  test('rejects negative leftIndent', async () => {
    await expect(handler({ index: 0, leftIndent: -50 }, mockBridge()))
      .rejects.toThrow('non-negative');
  });

  test('rejects negative rightIndent', async () => {
    await expect(handler({ index: 0, rightIndent: -10 }, mockBridge()))
      .rejects.toThrow('non-negative');
  });

  test('allows negative firstLineIndent (hanging indent)', async () => {
    const bridge = mockBridge({ setParagraphSpacing: { success: true } });
    const result = await handler({ index: 0, firstLineIndent: -36 }, bridge);
    const parsed = JSON.parse(result.content[0].text);
    expect(parsed.success).toBe(true);
  });

  test('rejects firstLineIndent below -1584', async () => {
    await expect(handler({ index: 0, firstLineIndent: -2000 }, mockBridge()))
      .rejects.toThrow('out of range');
  });
});

// =============================================================================
// checkSpacingBounds helper
// =============================================================================
describe('checkSpacingBounds', () => {
  test('accepts 0', () => {
    expect(() => checkSpacingBounds(0, 'leftIndent')).not.toThrow();
  });

  test('accepts normal values', () => {
    expect(() => checkSpacingBounds(72, 'leftIndent')).not.toThrow();
    expect(() => checkSpacingBounds(144, 'rightIndent')).not.toThrow();
  });

  test('accepts maximum value', () => {
    expect(() => checkSpacingBounds(MAX_SPACING_POINTS, 'leftIndent')).not.toThrow();
  });

  test('rejects value above maximum', () => {
    expect(() => checkSpacingBounds(MAX_SPACING_POINTS + 1, 'leftIndent')).toThrow('exceeds maximum');
  });

  test('rejects absurdly large value', () => {
    expect(() => checkSpacingBounds(99999, 'leftIndent')).toThrow('exceeds maximum');
  });

  test('error message includes value and field name', () => {
    expect(() => checkSpacingBounds(5000, 'leftIndent')).toThrow('leftIndent value 5000');
  });
});

// =============================================================================
// set_custom_property key length validation
// =============================================================================
describe('set_custom_property key length validation', () => {
  const handler = handlers.get('word_set_custom_property')!;

  test('rejects key longer than 255 characters', async () => {
    const longKey = 'k'.repeat(280);
    await expect(handler({ key: longKey, value: 'val' }, mockBridge()))
      .rejects.toThrow('255 characters or fewer (got 280)');
  });

  test('accepts key at exactly 255 characters', async () => {
    const bridge = mockBridge({ setCustomProperty: { success: true } });
    const key255 = 'k'.repeat(255);
    const result = await handler({ key: key255, value: 'val' }, bridge);
    const parsed = JSON.parse(result.content[0].text);
    expect(parsed.success).toBe(true);
  });

  test('rejects empty key', async () => {
    await expect(handler({ key: '', value: 'val' }, mockBridge()))
      .rejects.toThrow('non-empty');
  });

  test('rejects whitespace-only key', async () => {
    await expect(handler({ key: '   ', value: 'val' }, mockBridge()))
      .rejects.toThrow('non-empty');
  });
});

// =============================================================================
// checkPropertyKeyLength helper
// =============================================================================
describe('checkPropertyKeyLength', () => {
  test('accepts key within limit', () => {
    expect(() => checkPropertyKeyLength('normalKey')).not.toThrow();
    expect(() => checkPropertyKeyLength('k'.repeat(255))).not.toThrow();
  });

  test('rejects key exceeding limit', () => {
    expect(() => checkPropertyKeyLength('k'.repeat(256))).toThrow('255 characters or fewer');
  });

  test('error message includes actual length', () => {
    expect(() => checkPropertyKeyLength('k'.repeat(300))).toThrow('got 300');
  });

  test('MAX_CUSTOM_PROPERTY_KEY_LENGTH is 255', () => {
    expect(MAX_CUSTOM_PROPERTY_KEY_LENGTH).toBe(255);
  });
});

// =============================================================================
// Case-insensitive alignment normalization
// =============================================================================
describe('Case-insensitive alignment normalization', () => {
  test('accepts lowercase "left"', () => {
    expect(normalizeAlignment('left')).toBe('Left');
  });

  test('accepts uppercase "LEFT"', () => {
    expect(normalizeAlignment('LEFT')).toBe('Left');
  });

  test('accepts mixed case "Center"', () => {
    expect(normalizeAlignment('Center')).toBe('Centered');
  });

  test('accepts "CENTER"', () => {
    expect(normalizeAlignment('CENTER')).toBe('Centered');
  });

  test('accepts "right"', () => {
    expect(normalizeAlignment('right')).toBe('Right');
  });

  test('accepts "justified"', () => {
    expect(normalizeAlignment('justified')).toBe('Justified');
  });

  test('accepts "JUSTIFY"', () => {
    expect(normalizeAlignment('JUSTIFY')).toBe('Justified');
  });

  test('still rejects invalid values', () => {
    expect(() => normalizeAlignment('middle')).toThrow('Invalid alignment');
    expect(() => normalizeAlignment('diagonal')).toThrow('Invalid alignment');
  });
});

// =============================================================================
// Document outline maxLevel filtering
// =============================================================================
describe('Document outline maxLevel filtering', () => {
  const outlineHandler = handlers.get('word_get_document_outline')!;

  const outlineBridge: any = {
    send: async () => ({
      count: 3,
      paragraphs: [
        { index: 0, text: 'H1', style: 'Heading 1', outlineLevel: 1, isTocEntry: false },
        { index: 1, text: 'H2', style: 'Heading 2', outlineLevel: 2, isTocEntry: false },
        { index: 2, text: 'Body', style: 'Normal', outlineLevel: 10, isTocEntry: false },
      ],
    }),
  };

  test('maxLevel 1 returns only level 1 headings', async () => {
    const result = await outlineHandler({ maxLevel: 1 }, outlineBridge);
    const parsed = JSON.parse(result.content[0].text);
    expect(parsed.count).toBe(1);
    expect(parsed.outline[0].text).toBe('H1');
  });

  test('maxLevel 3 (default) returns both headings', async () => {
    const result = await outlineHandler({}, outlineBridge);
    const parsed = JSON.parse(result.content[0].text);
    expect(parsed.count).toBe(2);
  });

  test('maxLevel 9 returns both headings', async () => {
    const result = await outlineHandler({ maxLevel: 9 }, outlineBridge);
    const parsed = JSON.parse(result.content[0].text);
    expect(parsed.count).toBe(2);
  });

  test('maxLevel 0 rejects with validation error', async () => {
    await expect(outlineHandler({ maxLevel: 0 }, outlineBridge))
      .rejects.toThrow('maxLevel must be an integer between 1 and 9');
  });

  test('maxLevel 10 rejects with validation error', async () => {
    await expect(outlineHandler({ maxLevel: 10 }, outlineBridge))
      .rejects.toThrow('maxLevel must be an integer between 1 and 9');
  });

  test('maxLevel -1 rejects with validation error', async () => {
    await expect(outlineHandler({ maxLevel: -1 }, outlineBridge))
      .rejects.toThrow('maxLevel must be an integer between 1 and 9');
  });

  test('maxLevel 1.5 rejects with validation error', async () => {
    await expect(outlineHandler({ maxLevel: 1.5 }, outlineBridge))
      .rejects.toThrow('maxLevel must be an integer between 1 and 9');
  });
});

// =============================================================================
// Move/copy paragraph validation completeness
// =============================================================================
describe('Move/copy paragraph validation completeness', () => {
  const moveHandler = handlers.get('word_move_paragraph')!;
  const copyHandler = handlers.get('word_copy_paragraph')!;

  test('move rejects count=0', async () => {
    await expect(moveHandler({ fromIndex: 0, toIndex: 3, count: 0 }, mockBridgeWithParas()))
      .rejects.toThrow('at least 1');
  });

  test('move rejects toIndex inside source range', async () => {
    await expect(moveHandler({ fromIndex: 1, toIndex: 2, count: 3 }, mockBridgeWithParas()))
      .rejects.toThrow('inside the source range');
  });

  test('move rejects same from and to (count=1)', async () => {
    await expect(moveHandler({ fromIndex: 2, toIndex: 2 }, mockBridgeWithParas()))
      .rejects.toThrow('must be different');
  });

  test('copy accepts same from and to (duplicates in place)', async () => {
    const bridge = mockBridge();
    const result = await copyHandler({ fromIndex: 0, toIndex: 0 }, bridge);
    expect(result.content[0].text).toContain('success');
  });

  test('move rejects fromIndex + count exceeding paragraph count', async () => {
    await expect(moveHandler({ fromIndex: 3, toIndex: 0, count: 5 }, mockBridgeWithParas()))
      .rejects.toThrow('exceeds paragraph count');
  });
});

// =============================================================================
// Batch partial failure semantics
// =============================================================================
describe('Batch stops at first failure', () => {
  const batchHandler = handlers.get('word_batch')!;

  test('first op succeeds, second fails, third not executed', async () => {
    const calls: string[] = [];
    const batchBridge: any = {
      send: async (action: string, params: any) => {
        if (action === 'batchExecute') {
          const results = params.operations.map((op: any, i: number) => {
            calls.push(op.action);
            if (op.action === 'getDocumentText') return { index: i, success: true, result: { text: 'hi' } };
            if (op.action === 'deleteParagraph') return { index: i, success: false, error: 'Index out of range' };
            return { index: i, success: true, result: {} };
          });
          return { results };
        }
        return {};
      },
    };

    const result = await batchHandler({
      operations: [
        { tool: 'word_get_text', args: {} },
        { tool: 'word_delete_paragraph', args: { index: 999 } },
        { tool: 'word_get_text', args: {} },
      ],
    }, batchBridge);

    const parsed = JSON.parse(result.content[0].text);
    expect(parsed.completed).toBe(1);
    expect(parsed.failed).toBe(1);
    expect(parsed.results.length).toBe(2);
  });
});

// =============================================================================
// move_paragraph rejects table-internal paragraphs
// =============================================================================
describe('move_paragraph rejects table-internal paragraphs', () => {
  const { handlers } = buildToolRegistry();
  const moveHandler = handlers.get('word_move_paragraph')!;

  function bridgeWithTablePara(): any {
    return {
      send: async (action: string) => {
        if (action === 'getParagraphs') {
          return {
            count: 5,
            paragraphs: [
              { index: 0, text: 'Normal', style: 'Normal', inTable: false, isTocEntry: false, outlineLevel: 10 },
              { index: 1, text: 'Cell A', style: 'Normal', inTable: true, isTocEntry: false, outlineLevel: 10 },
              { index: 2, text: 'Cell B', style: 'Normal', inTable: true, isTocEntry: false, outlineLevel: 10 },
              { index: 3, text: 'After', style: 'Normal', inTable: false, isTocEntry: false, outlineLevel: 10 },
              { index: 4, text: 'End', style: 'Normal', inTable: false, isTocEntry: false, outlineLevel: 10 },
            ],
          };
        }
        return { ooxml: '<pkg:package></pkg:package>' };
      },
    };
  }

  test('rejects moving a paragraph that is inside a table', async () => {
    await expect(moveHandler({ fromIndex: 1, toIndex: 4 }, bridgeWithTablePara()))
      .rejects.toThrow('inside a table cell');
  });

  test('rejects moving a range that includes table paragraphs', async () => {
    await expect(moveHandler({ fromIndex: 1, toIndex: 4, count: 2 }, bridgeWithTablePara()))
      .rejects.toThrow('inside a table cell');
  });

  test('allows moving non-table paragraphs', async () => {
    const calls: any[] = [];
    const bridge: any = {
      send: async (action: string, params: any) => {
        calls.push({ action, params });
        if (action === 'getParagraphs') {
          return {
            count: 5,
            paragraphs: [
              { index: 0, text: 'A', style: 'Normal', inTable: false, isTocEntry: false, outlineLevel: 10 },
              { index: 1, text: 'B', style: 'Normal', inTable: false, isTocEntry: false, outlineLevel: 10 },
              { index: 2, text: 'C', style: 'Normal', inTable: false, isTocEntry: false, outlineLevel: 10 },
              { index: 3, text: 'D', style: 'Normal', inTable: false, isTocEntry: false, outlineLevel: 10 },
              { index: 4, text: 'E', style: 'Normal', inTable: false, isTocEntry: false, outlineLevel: 10 },
            ],
          };
        }
        if (action === 'getParaOoxml') return { ooxml: '<pkg:package></pkg:package>' };
        return { success: true };
      },
    };
    const result = await moveHandler({ fromIndex: 0, toIndex: 4 }, bridge);
    const parsed = JSON.parse(result.content[0].text);
    expect(parsed.success).toBe(true);
  });
});

// =============================================================================
// move_paragraph detects no-op when destination equals source position
// =============================================================================
describe('move_paragraph detects no-op', () => {
  const { handlers } = buildToolRegistry();
  const moveHandler = handlers.get('word_move_paragraph')!;

  function bridgeWithNParas(n: number): any {
    const paragraphs = Array.from({ length: n }, (_, i) => ({
      index: i, text: `P${i}`, style: 'Normal', inTable: false, isTocEntry: false, outlineLevel: 10,
    }));
    return {
      send: async (action: string) => {
        if (action === 'getParagraphs') return { total: n, count: n, paragraphs };
        return { ooxml: '<pkg:package></pkg:package>' };
      },
    };
  }

  test('returns warning when moving range immediately after itself', async () => {
    // fromIndex=0, toIndex=3, count=3: moves [P0,P1,P2] to after P3
    // This IS a reorder (P3 ends up before the block), so no no-op warning
    const result = await moveHandler({ fromIndex: 0, toIndex: 3, count: 3 }, bridgeWithNParas(5));
    const parsed = JSON.parse(result.content[0].text);
    expect(parsed.success).toBe(true);
    expect(parsed.moved).not.toBeNull();
  });

  test('returns warning for single paragraph no-op (fromIndex=0, toIndex=1, After)', async () => {
    // fromIndex=0, toIndex=1: moves P0 to after P1 — this IS a reorder
    const result = await moveHandler({ fromIndex: 0, toIndex: 1 }, bridgeWithNParas(5));
    const parsed = JSON.parse(result.content[0].text);
    expect(parsed.success).toBe(true);
    expect(parsed.moved).not.toBeNull();
  });

  test('rejects moving the trailing empty paragraph', async () => {
    const paragraphs = Array.from({ length: 5 }, (_, i) => ({
      index: i, text: i === 4 ? '' : `P${i}`, style: 'Normal', inTable: false, isTocEntry: false, outlineLevel: 10,
    }));
    const bridge: any = {
      send: async (action: string) => {
        if (action === 'getParagraphs') return { total: 5, count: 5, paragraphs };
        return { ooxml: '<pkg:package></pkg:package>' };
      },
    };
    await expect(moveHandler({ fromIndex: 4, toIndex: 0 }, bridge))
      .rejects.toThrow('Cannot move the last paragraph');
  });

  test('allows moving last paragraph if it has content', async () => {
    const result = await moveHandler({ fromIndex: 4, toIndex: 0 }, bridgeWithNParas(5));
    const parsed = JSON.parse(result.content[0].text);
    expect(parsed.success).toBe(true);
  });

  test('detects no-op: move to After the preceding paragraph (fromIndex=1, toIndex=0, After)', async () => {
    // P1 is already after P0 — moving P1 to "After P0" is a no-op
    const result = await moveHandler({ fromIndex: 1, toIndex: 0 }, bridgeWithNParas(5));
    const parsed = JSON.parse(result.content[0].text);
    expect(parsed.warning).toContain('No move performed');
    expect(parsed.moved).toBeNull();
  });

  test('detects no-op: move to Before the following paragraph (fromIndex=1, toIndex=2, Before)', async () => {
    // P1 is already before P2 — moving P1 to "Before P2" is a no-op
    const result = await moveHandler({ fromIndex: 1, toIndex: 2, location: 'Before' }, bridgeWithNParas(5));
    const parsed = JSON.parse(result.content[0].text);
    expect(parsed.warning).toContain('No move performed');
    expect(parsed.moved).toBeNull();
  });
});

// =============================================================================
// copy_paragraph handles self-copy by adjusting target index
// =============================================================================
describe('copy_paragraph self-copy handling', () => {
  const { handlers } = buildToolRegistry();
  const copyHandler = handlers.get('word_copy_paragraph')!;

  test('self-copy adjusts effective target to after source range', async () => {
    const calls: any[] = [];
    const bridge: any = {
      send: async (action: string, params: any) => {
        calls.push({ action, params });
        if (action === 'getParagraphs') {
          return {
            count: 3,
            paragraphs: [
              { index: 0, text: 'A', style: 'Normal', inTable: false, isTocEntry: false, outlineLevel: 10 },
              { index: 1, text: 'B', style: 'Normal', inTable: false, isTocEntry: false, outlineLevel: 10 },
              { index: 2, text: 'C', style: 'Normal', inTable: false, isTocEntry: false, outlineLevel: 10 },
            ],
          };
        }
        if (action === 'getParaOoxml') return { ooxml: '<pkg:package></pkg:package>' };
        return { success: true };
      },
    };
    const result = await copyHandler({ fromIndex: 0, toIndex: 0 }, bridge);
    const parsed = JSON.parse(result.content[0].text);
    expect(parsed.success).toBe(true);
    // The insertOoxmlAtIndex call should use index 0 with location 'After'
    const insertCall = calls.find(c => c.action === 'insertOoxmlAtIndex');
    expect(insertCall).toBeDefined();
    expect(insertCall.params.index).toBe(0);
    expect(insertCall.params.location).toBe('After');
  });

  test('copy to different index uses original target', async () => {
    const calls: any[] = [];
    const bridge: any = {
      send: async (action: string, params: any) => {
        calls.push({ action, params });
        if (action === 'getParagraphs') {
          return {
            count: 5,
            paragraphs: Array.from({ length: 5 }, (_, i) => ({
              index: i, text: `P${i}`, style: 'Normal', inTable: false, isTocEntry: false, outlineLevel: 10,
            })),
          };
        }
        if (action === 'getParaOoxml') return { ooxml: '<pkg:package></pkg:package>' };
        return { success: true };
      },
    };
    await copyHandler({ fromIndex: 0, toIndex: 3 }, bridge);
    const insertCall = calls.find(c => c.action === 'insertOoxmlAtIndex');
    expect(insertCall.params.index).toBe(3);
    expect(insertCall.params.location).toBe('After');
  });

  test('rejects copy from table-internal paragraphs', async () => {
    const bridge: any = {
      send: async (action: string) => {
        if (action === 'getParagraphs') {
          return {
            count: 3,
            paragraphs: [
              { index: 0, text: 'Cell', style: 'Normal', inTable: true, isTocEntry: false, outlineLevel: 10 },
              { index: 1, text: 'Normal', style: 'Normal', inTable: false, isTocEntry: false, outlineLevel: 10 },
              { index: 2, text: 'End', style: 'Normal', inTable: false, isTocEntry: false, outlineLevel: 10 },
            ],
          };
        }
        return { ooxml: '<pkg:package></pkg:package>' };
      },
    };
    await expect(copyHandler({ fromIndex: 0, toIndex: 2 }, bridge))
      .rejects.toThrow('inside a table cell');
  });
});

// =============================================================================
// search tool description documents Word special codes
// =============================================================================
describe('search tool description includes special codes note', () => {
  const { tools } = buildToolRegistry();

  test('word_search description mentions search codes', () => {
    const searchTool = tools.find(t => t.name === 'word_search');
    expect(searchTool).toBeDefined();
    expect(searchTool!.description).toContain('^p');
    expect(searchTool!.description).toContain('^t');
    expect(searchTool!.description).toContain('^^');
  });

  test('word_insert_footnote description warns about insertion order', () => {
    const fnTool = tools.find(t => t.name === 'word_insert_footnote');
    expect(fnTool).toBeDefined();
    expect(fnTool!.description).toContain('reverse');
  });
});

// =============================================================================
// insert_table headerRowCount defaults to 0
// =============================================================================
describe('insert_table headerRowCount default', () => {
  const handler = handlers.get('word_insert_table')!;

  test('forwards headerRowCount=0 when not specified', async () => {
    const bridge = mockBridge({ insertTable: { success: true } });
    await handler({ rows: 2, cols: 2 }, bridge);
    const call = bridge.calls.find((c: any) => c.action === 'insertTable');
    expect(call).toBeDefined();
    // The parameter is not set by the server (taskpane handles default)
    // Verify it's forwarded as-is (undefined means taskpane will use ?? 0)
    expect(call.params.headerRowCount).toBeUndefined();
  });

  test('forwards explicit headerRowCount to bridge', async () => {
    const bridge = mockBridge({ insertTable: { success: true } });
    await handler({ rows: 2, cols: 2, headerRowCount: 1 }, bridge);
    const call = bridge.calls.find((c: any) => c.action === 'insertTable');
    expect(call.params.headerRowCount).toBe(1);
  });

  test('forwards headerRowCount=0 explicitly', async () => {
    const bridge = mockBridge({ insertTable: { success: true } });
    await handler({ rows: 2, cols: 2, headerRowCount: 0 }, bridge);
    const call = bridge.calls.find((c: any) => c.action === 'insertTable');
    expect(call.params.headerRowCount).toBe(0);
  });

  test('tool schema documents default as 0', () => {
    const { tools } = buildToolRegistry();
    const tableTool = tools.find(t => t.name === 'word_insert_table');
    expect(tableTool).toBeDefined();
    const desc = (tableTool!.schema.properties as any).headerRowCount.description;
    expect(desc).toContain('default: 0');
  });
});

// =============================================================================
// Batch: server-composed tool execution path
// =============================================================================
describe('Batch executes server-composed tools', () => {
  const batchHandler = handlers.get('word_batch')!;

  test('executes server-composed tool (move_paragraph) within batch', async () => {
    const paragraphs = Array.from({ length: 5 }, (_, i) => ({
      index: i, text: `P${i}`, style: 'Normal', inTable: false, isTocEntry: false, outlineLevel: 10,
    }));
    const bridge: any = {
      send: async (action: string) => {
        if (action === 'getParagraphs') return { total: 5, count: 5, paragraphs };
        if (action === 'batchExecute') return { results: [] };
        if (action === 'getParaOoxml') return { ooxml: '<pkg:package></pkg:package>' };
        return { success: true };
      },
    };
    const result = await batchHandler({
      operations: [{ tool: 'word_move_paragraph', args: { fromIndex: 4, toIndex: 0 } }],
    }, bridge);
    const parsed = JSON.parse(result.content[0].text);
    expect(parsed.completed).toBe(1);
    expect(parsed.results[0].success).toBe(true);
  });

  test('stops on server-composed tool failure', async () => {
    const bridge: any = {
      send: async (action: string) => {
        if (action === 'batchExecute') return { results: [] };
        if (action === 'getParagraphs') return { total: 2, count: 2, paragraphs: [] };
        return { success: true };
      },
    };
    const result = await batchHandler({
      operations: [
        { tool: 'word_move_paragraph', args: { fromIndex: 99, toIndex: 0 } },
        { tool: 'word_get_text', args: {} },
      ],
    }, bridge);
    const parsed = JSON.parse(result.content[0].text);
    expect(parsed.failed).toBe(1);
    expect(parsed.completed).toBe(0);
    expect(parsed.results[0].success).toBe(false);
    expect(parsed.results[0].error).toContain('out of range');
  });

  test('native batch failure stops execution', async () => {
    const bridge: any = {
      send: async (action: string) => {
        if (action === 'batchExecute') return {
          results: [{ index: 0, success: false, error: 'simulated failure' }],
        };
        return { success: true };
      },
    };
    const result = await batchHandler({
      operations: [
        { tool: 'word_get_text', args: {} },
        { tool: 'word_get_text', args: {} },
      ],
    }, bridge);
    const parsed = JSON.parse(result.content[0].text);
    expect(parsed.failed).toBe(1);
    expect(parsed.completed).toBe(0);
    expect(parsed.results[0].error).toBe('simulated failure');
  });
});

// =============================================================================
// Document outline outlineLevel fallback
// =============================================================================
describe('Document outline outlineLevel fallback', () => {
  const outlineHandler = handlers.get('word_get_document_outline')!;

  test('uses outlineLevel when style does not match Heading N', async () => {
    const paragraphs = [
      { index: 0, text: 'Custom styled heading', style: 'CustomStyle', outlineLevel: 2, isTocEntry: false, inTable: false, isListItem: false },
      { index: 1, text: 'Body text', style: 'Normal', outlineLevel: 10, isTocEntry: false, inTable: false, isListItem: false },
    ];
    const bridge: any = {
      send: async () => ({ total: 2, count: 2, paragraphs }),
    };
    const result = await outlineHandler({}, bridge);
    const parsed = JSON.parse(result.content[0].text);
    expect(parsed.count).toBe(1);
    expect(parsed.outline[0].text).toBe('Custom styled heading');
    expect(parsed.outline[0].level).toBe(2);
  });

  test('skips TOC entries even with valid outlineLevel', async () => {
    const paragraphs = [
      { index: 0, text: 'TOC entry\t1', style: 'TOC 1', outlineLevel: 1, isTocEntry: true, inTable: false, isListItem: false },
      { index: 1, text: 'Real heading', style: 'Heading 1', outlineLevel: 1, isTocEntry: false, inTable: false, isListItem: false },
    ];
    const bridge: any = {
      send: async () => ({ total: 2, count: 2, paragraphs }),
    };
    const result = await outlineHandler({}, bridge);
    const parsed = JSON.parse(result.content[0].text);
    expect(parsed.count).toBe(1);
    expect(parsed.outline[0].text).toBe('Real heading');
  });
});

// =============================================================================
// Equations handler: display mode and error paths
// =============================================================================
describe('Equations handler', () => {
  const eqHandler = handlers.get('word_insert_equation')!;

  test('rejects empty latex string', async () => {
    const bridge: any = { send: async () => ({}) };
    await expect(eqHandler({ latex: '' }, bridge))
      .rejects.toThrow('latex must be a non-empty string');
  });

  test('rejects invalid latex', async () => {
    const bridge: any = { send: async () => ({}) };
    const result = await eqHandler({ latex: '\\frac{' }, bridge);
    expect(result.isError).toBe(true);
    expect(result.content[0].text).toContain('LaTeX parse error');
  });

  test('display mode inserts OOXML at end', async () => {
    const calls: string[] = [];
    const bridge: any = { send: async (action: string) => { calls.push(action); return {}; } };
    const result = await eqHandler({ latex: 'x^2' }, bridge);
    const parsed = JSON.parse(result.content[0].text);
    expect(parsed.success).toBe(true);
    expect(parsed.displayMode).toBe(true);
    expect(calls).toContain('insertOoxml');
  });

  test('inline mode without anchor inserts at selection', async () => {
    const calls: string[] = [];
    const bridge: any = { send: async (action: string) => { calls.push(action); return {}; } };
    const result = await eqHandler({ latex: 'y', displayMode: false }, bridge);
    const parsed = JSON.parse(result.content[0].text);
    expect(parsed.success).toBe(true);
    expect(parsed.displayMode).toBe(false);
    expect(calls).toContain('insertOoxmlAtSelection');
  });

  test('inline mode with anchor uses insertOoxmlAfterMatch', async () => {
    const calls: Array<{ action: string; params: any }> = [];
    const bridge: any = {
      send: async (action: string, params: any) => {
        calls.push({ action, params });
        return { success: true };
      },
    };
    const result = await eqHandler({ latex: 'z', displayMode: false, anchorText: 'hello' }, bridge);
    const parsed = JSON.parse(result.content[0].text);
    expect(parsed.success).toBe(true);
    expect(calls.some(c => c.action === 'insertOoxmlAfterMatch')).toBe(true);
    const matchCall = calls.find(c => c.action === 'insertOoxmlAfterMatch')!;
    expect(matchCall.params.anchorText).toBe('hello');
    expect(matchCall.params.occurrence).toBe(0);
    expect(matchCall.params.ooxml).toBeDefined();
  });

  test('inline mode with anchor passes occurrence and matchCase', async () => {
    const calls: Array<{ action: string; params: any }> = [];
    const bridge: any = {
      send: async (action: string, params: any) => {
        calls.push({ action, params });
        return { success: true };
      },
    };
    await eqHandler({ latex: 'z', displayMode: false, anchorText: 'hi', occurrence: 2, matchCase: true }, bridge);
    const matchCall = calls.find(c => c.action === 'insertOoxmlAfterMatch')!;
    expect(matchCall.params.occurrence).toBe(2);
    expect(matchCall.params.matchCase).toBe(true);
  });

  test('inline mode with nonexistent anchor propagates bridge error', async () => {
    const bridge: any = {
      send: async (action: string) => {
        if (action === 'insertOoxmlAfterMatch') throw new Error('Anchor not found: nope');
        return {};
      },
    };
    await expect(eqHandler({ latex: 'a', displayMode: false, anchorText: 'nope' }, bridge))
      .rejects.toThrow('Anchor not found');
  });
});

// =============================================================================
// format_text server-side validation
// =============================================================================
describe('format_text server-side validation', () => {
  const handler = handlers.get('word_format_text')!;
  const bridge = mockBridge();

  test('size=0 as only property throws "size must be positive"', async () => {
    await expect(handler({ text: 'hello', size: 0 }, bridge))
      .rejects.toThrow('size must be positive');
  });

  test('size=0 with bold=true still throws "size must be positive"', async () => {
    await expect(handler({ text: 'hello', size: 0, bold: true }, bridge))
      .rejects.toThrow('size must be positive');
  });

  test('no formatting properties throws correct error', async () => {
    await expect(handler({ text: 'hello' }, bridge))
      .rejects.toThrow('At least one formatting property');
  });

  test('size=1 is accepted', async () => {
    const result = await handler({ text: 'hello', size: 1 }, bridge);
    const parsed = JSON.parse(result.content[0].text);
    expect(parsed.success).toBe(true);
  });

  test('size=1639 is rejected', async () => {
    await expect(handler({ text: 'hello', size: 1639 }, bridge))
      .rejects.toThrow('must not exceed 1638');
  });

  test('invalid hex color is rejected', async () => {
    await expect(handler({ text: 'hello', color: '#ZZZZZZ' }, bridge))
      .rejects.toThrow('valid hex color');
  });

  test('invalid highlightColor is rejected', async () => {
    await expect(handler({ text: 'hello', highlightColor: 'Purple' }, bridge))
      .rejects.toThrow('Invalid highlightColor');
  });

  test('empty text is rejected', async () => {
    await expect(handler({ text: '', bold: true }, bridge))
      .rejects.toThrow('must be a non-empty string');
  });
});

// =============================================================================
// reply_to_comment rejects empty text
// =============================================================================
describe('reply_to_comment server-side validation', () => {
  const handler = handlers.get('word_reply_to_comment')!;
  const bridge = mockBridge();

  test('empty text is rejected', async () => {
    await expect(handler({ commentId: '123', text: '' }, bridge))
      .rejects.toThrow('must be a non-empty string');
  });

  test('whitespace-only text is rejected', async () => {
    await expect(handler({ commentId: '123', text: '   ' }, bridge))
      .rejects.toThrow('must be a non-empty string');
  });

  test('valid text is accepted', async () => {
    const result = await handler({ commentId: '123', text: 'Good point!' }, bridge);
    const parsed = JSON.parse(result.content[0].text);
    expect(parsed.success).toBe(true);
  });
});

// =============================================================================
// set_paragraph_spacing requires at least one property
// =============================================================================
describe('set_paragraph_spacing requires at least one property', () => {
  const handler = handlers.get('word_set_paragraph_spacing')!;
  const bridge = mockBridge();

  test('no properties throws correct error', async () => {
    await expect(handler({ index: 0 }, bridge))
      .rejects.toThrow('At least one spacing or indent property');
  });

  test('lineSpacing alone is accepted', async () => {
    const result = await handler({ index: 0, lineSpacing: 12 }, bridge);
    const parsed = JSON.parse(result.content[0].text);
    expect(parsed.success).toBe(true);
  });

  test('firstLineIndent alone is accepted', async () => {
    const result = await handler({ index: 0, firstLineIndent: 36 }, bridge);
    const parsed = JSON.parse(result.content[0].text);
    expect(parsed.success).toBe(true);
  });
});

// =============================================================================
// move/copy rejects table-internal destination
// =============================================================================
describe('move/copy rejects table-internal destination', () => {
  const { handlers: h } = buildToolRegistry();
  const moveHandler = h.get('word_move_paragraph')!;
  const copyHandler = h.get('word_copy_paragraph')!;

  function bridgeWithTable(): any {
    return {
      send: async (action: string) => {
        if (action === 'getParagraphs') {
          return {
            total: 5,
            count: 5,
            paragraphs: [
              { index: 0, text: 'Before', inTable: false },
              { index: 1, text: 'Cell', inTable: true },
              { index: 2, text: 'Cell2', inTable: true },
              { index: 3, text: 'After', inTable: false },
              { index: 4, text: 'End', inTable: false },
            ],
          };
        }
        if (action === 'getParaOoxml') return { ooxml: '<pkg:package></pkg:package>' };
        return { success: true };
      },
    };
  }

  test('move to table cell destination is rejected', async () => {
    await expect(moveHandler({ fromIndex: 0, toIndex: 1 }, bridgeWithTable()))
      .rejects.toThrow('inside a table cell');
  });

  test('copy to table cell destination is rejected', async () => {
    await expect(copyHandler({ fromIndex: 3, toIndex: 2 }, bridgeWithTable()))
      .rejects.toThrow('inside a table cell');
  });

  test('move to non-table destination is allowed', async () => {
    const result = await moveHandler({ fromIndex: 0, toIndex: 3 }, bridgeWithTable());
    const parsed = JSON.parse(result.content[0].text);
    expect(parsed.success).toBe(true);
  });
});

// =============================================================================
// insert_text_at_match rejects empty text
// =============================================================================
describe('insert_text_at_match validates text', () => {
  const handler = handlers.get('word_insert_text_at_match')!;

  test('rejects empty text', async () => {
    await expect(handler({ text: '', after: 'hello' }, mockBridge()))
      .rejects.toThrow('non-empty');
  });

  test('rejects whitespace-only text', async () => {
    await expect(handler({ text: '   ', after: 'hello' }, mockBridge()))
      .rejects.toThrow('non-empty');
  });

  test('accepts non-empty text', async () => {
    const bridge = mockBridge();
    const result = await handler({ text: 'world', after: 'hello' }, bridge);
    const parsed = JSON.parse(result.content[0].text);
    expect(parsed.success).toBe(true);
  });
});

// =============================================================================
// set_paragraph_style requires at least one of style/alignment
// =============================================================================
describe('set_paragraph_style requires style or alignment', () => {
  const handler = handlers.get('word_set_paragraph_style')!;

  test('rejects when neither style nor alignment provided', async () => {
    await expect(handler({ index: 0 }, mockBridge()))
      .rejects.toThrow('At least one');
  });

  test('accepts style alone', async () => {
    const bridge = mockBridge();
    const result = await handler({ index: 0, style: 'Normal' }, bridge);
    const parsed = JSON.parse(result.content[0].text);
    expect(parsed.success).toBe(true);
  });

  test('accepts alignment alone', async () => {
    const bridge = mockBridge();
    const result = await handler({ index: 0, alignment: 'Center' }, bridge);
    const parsed = JSON.parse(result.content[0].text);
    expect(parsed.success).toBe(true);
  });
});

// =============================================================================
// Table tool server-side validators
// =============================================================================
describe('table tool server-side validators', () => {
  const insertTable = handlers.get('word_insert_table')!;
  const setCell = handlers.get('word_set_table_cell')!;
  const deleteRow = handlers.get('word_delete_table_row')!;
  const deleteTable = handlers.get('word_delete_table')!;
  const mergeCells = handlers.get('word_merge_table_cells')!;
  const splitCell = handlers.get('word_split_table_cell')!;
  const shading = handlers.get('word_set_table_cell_shading')!;

  test('insertTable rejects rows <= 0', async () => {
    await expect(insertTable({ rows: 0, cols: 3 }, mockBridge()))
      .rejects.toThrow('rows must be a positive integer');
  });

  test('insertTable rejects cols > 63', async () => {
    await expect(insertTable({ rows: 1, cols: 64 }, mockBridge()))
      .rejects.toThrow('cols must not exceed 63');
  });

  test('insertTable rejects data row count mismatch', async () => {
    await expect(insertTable({ rows: 2, cols: 2, data: [['a', 'b']] }, mockBridge()))
      .rejects.toThrow('Data rows');
  });

  test('insertTable rejects data column count mismatch', async () => {
    await expect(insertTable({ rows: 1, cols: 2, data: [['a']] }, mockBridge()))
      .rejects.toThrow('columns but expected');
  });

  test('insertTable accepts valid input', async () => {
    const bridge = mockBridge();
    const result = await insertTable({ rows: 2, cols: 2, data: [['a', 'b'], ['c', 'd']] }, bridge);
    const parsed = JSON.parse(result.content[0].text);
    expect(parsed.success).toBe(true);
  });

  test('setTableCell rejects negative tableIndex', async () => {
    await expect(setCell({ tableIndex: -1, row: 0, col: 0, text: 'x' }, mockBridge()))
      .rejects.toThrow('non-negative integer');
  });

  test('setTableCell rejects negative row', async () => {
    await expect(setCell({ tableIndex: 0, row: -1, col: 0, text: 'x' }, mockBridge()))
      .rejects.toThrow('non-negative integer');
  });

  test('deleteTableRow rejects negative indices', async () => {
    await expect(deleteRow({ tableIndex: -1, rowIndex: 0 }, mockBridge()))
      .rejects.toThrow('non-negative integer');
  });

  test('deleteTable rejects negative index', async () => {
    await expect(deleteTable({ index: -1 }, mockBridge()))
      .rejects.toThrow('non-negative integer');
  });

  test('deleteTable forwards deleteCaption flag to bridge', async () => {
    const bridge = mockBridge();
    await deleteTable({ index: 0, deleteCaption: false }, bridge);
    const call = bridge.calls.find((c: any) => c.action === 'deleteTable');
    expect(call).toBeDefined();
    expect(call.params.deleteCaption).toBe(false);
  });

  test('mergeCells rejects topRow > bottomRow', async () => {
    await expect(mergeCells({ tableIndex: 0, topRow: 3, firstCell: 0, bottomRow: 1, lastCell: 2 }, mockBridge()))
      .rejects.toThrow('topRow');
  });

  test('mergeCells rejects single cell', async () => {
    await expect(mergeCells({ tableIndex: 0, topRow: 0, firstCell: 0, bottomRow: 0, lastCell: 0 }, mockBridge()))
      .rejects.toThrow('single cell');
  });

  test('splitCell rejects negative col', async () => {
    await expect(splitCell({ tableIndex: 0, row: 0, col: -1 }, mockBridge()))
      .rejects.toThrow('non-negative integer');
  });

  test('splitCell rejects colCount <= 0', async () => {
    await expect(splitCell({ tableIndex: 0, row: 0, col: 0, colCount: 0 }, mockBridge()))
      .rejects.toThrow('colCount must be a positive integer');
  });

  test('shading rejects invalid hex color', async () => {
    await expect(shading({ tableIndex: 0, row: 0, col: 0, color: 'red' }, mockBridge()))
      .rejects.toThrow('hex color');
  });

  test('shading accepts valid hex color', async () => {
    const bridge = mockBridge();
    const result = await shading({ tableIndex: 0, row: 0, col: 0, color: '#FF0000' }, bridge);
    const parsed = JSON.parse(result.content[0].text);
    expect(parsed.success).toBe(true);
  });
});

// =============================================================================
// delete_paragraph description documents fallback strategies
// =============================================================================
describe('delete_paragraph tool metadata', () => {
  const { tools } = buildToolRegistry();
  const tool = tools.find(t => t.name === 'word_delete_paragraph')!;

  test('description mentions fallback strategies', () => {
    expect(tool.description).toContain('fallback');
  });

  test('description mentions list formatting', () => {
    expect(tool.description).toContain('list formatting');
  });

  test('description mentions range deletion', () => {
    expect(tool.description).toContain('range deletion');
  });
});

// =============================================================================
// insert_text_at_match description documents 255-char anchor text limit
// =============================================================================
describe('insert_text_at_match tool metadata', () => {
  const { tools } = buildToolRegistry();
  const tool = tools.find(t => t.name === 'word_insert_text_at_match')!;

  test('description documents 255-char anchor text limit', () => {
    expect(tool.description).toContain('255');
  });
});

// =============================================================================
// delete_paragraph forwards to bridge
// =============================================================================
describe('delete_paragraph server-side forwarding', () => {
  const handler = handlers.get('word_delete_paragraph')!;

  test('forwards index directly to bridge (validation is taskpane-side)', async () => {
    const bridge = mockBridge({ deleteParagraph: { success: true } });
    const result = await handler({ index: 2 }, bridge);
    const parsed = JSON.parse(result.content[0].text);
    expect(parsed.success).toBe(true);
    expect(bridge.calls[0]).toEqual({ action: 'deleteParagraph', params: { index: 2 } });
  });

  test('passes through bridge error messages', async () => {
    const bridge: any = {
      calls: [] as any[],
      send: async (action: string, params: any) => {
        bridge.calls.push({ action, params });
        throw new Error('Cannot delete paragraph 5. Word refused the operation.');
      },
    };
    await expect(handler({ index: 5 }, bridge))
      .rejects.toThrow('Cannot delete paragraph 5');
  });
});

// =============================================================================
// insert_paragraph_at_index server-side forwarding
// =============================================================================
describe('insert_paragraph_at_index server-side forwarding', () => {
  const handler = handlers.get('word_insert_paragraph_at_index')!;

  test('forwards valid params to bridge', async () => {
    const bridge = mockBridge({ insertParagraphAtIndex: { success: true } });
    const result = await handler({ index: 3, text: 'test content', style: 'Heading 1' }, bridge);
    const parsed = JSON.parse(result.content[0].text);
    expect(parsed.success).toBe(true);
    expect(bridge.calls[0].action).toBe('insertParagraphAtIndex');
    expect(bridge.calls[0].params.index).toBe(3);
    expect(bridge.calls[0].params.text).toBe('test content');
    expect(bridge.calls[0].params.style).toBe('Heading 1');
  });

  test('forwards long text without server-side rejection', async () => {
    const longText = 'A'.repeat(600) + ' — with em-dashes (and parens) "quotes"';
    const bridge = mockBridge({ insertParagraphAtIndex: { success: true } });
    const result = await handler({ index: 0, text: longText }, bridge);
    const parsed = JSON.parse(result.content[0].text);
    expect(parsed.success).toBe(true);
    expect(bridge.calls[0].params.text).toBe(longText);
  });

  test('passes through bridge error messages', async () => {
    const bridge: any = {
      calls: [] as any[],
      send: async (action: string, params: any) => {
        bridge.calls.push({ action, params });
        throw new Error('Cannot insert paragraph at index 378. Primary insert failed.');
      },
    };
    await expect(handler({ index: 378, text: 'test' }, bridge))
      .rejects.toThrow('Cannot insert paragraph at index 378');
  });
});

// =============================================================================
// Track Changes validation
// =============================================================================
describe('Track Changes input validation', () => {
  const acceptHandler = handlers.get('word_accept_tracked_change')!;
  const rejectHandler = handlers.get('word_reject_tracked_change')!;
  const rangeHandler = handlers.get('word_accept_tracked_changes_in_range')!;

  test('accept rejects negative index', async () => {
    await expect(acceptHandler({ index: -1 }, mockBridge()))
      .rejects.toThrow('non-negative integer');
  });

  test('accept rejects float index', async () => {
    await expect(acceptHandler({ index: 1.5 }, mockBridge()))
      .rejects.toThrow('non-negative integer');
  });

  test('reject rejects negative index', async () => {
    await expect(rejectHandler({ index: -3 }, mockBridge()))
      .rejects.toThrow('non-negative integer');
  });

  test('reject rejects NaN index', async () => {
    await expect(rejectHandler({ index: NaN }, mockBridge()))
      .rejects.toThrow('non-negative integer');
  });

  test('accept_in_range rejects negative startIndex', async () => {
    await expect(rangeHandler({ startIndex: -1 }, mockBridge()))
      .rejects.toThrow('non-negative integer');
  });

  test('accept_in_range rejects float startIndex', async () => {
    await expect(rangeHandler({ startIndex: 2.5 }, mockBridge()))
      .rejects.toThrow('non-negative integer');
  });

  test('accept_in_range rejects endIndex <= startIndex', async () => {
    await expect(rangeHandler({ startIndex: 5, endIndex: 5 }, mockBridge()))
      .rejects.toThrow('endIndex must be greater than startIndex');
  });

  test('accept_in_range rejects endIndex < startIndex', async () => {
    await expect(rangeHandler({ startIndex: 10, endIndex: 3 }, mockBridge()))
      .rejects.toThrow('endIndex must be greater than startIndex');
  });

  test('accept_in_range rejects negative endIndex', async () => {
    await expect(rangeHandler({ startIndex: 0, endIndex: -1 }, mockBridge()))
      .rejects.toThrow('non-negative integer');
  });

  test('accept_in_range accepts valid range', async () => {
    const bridge = mockBridge({ acceptTrackedChangesInRange: { success: true } });
    const result = await rangeHandler({ startIndex: 0, endIndex: 5 }, bridge);
    const parsed = JSON.parse(result.content[0].text);
    expect(parsed.success).toBe(true);
  });

  test('accept_in_range accepts startIndex without endIndex', async () => {
    const bridge = mockBridge({ acceptTrackedChangesInRange: { success: true } });
    const result = await rangeHandler({ startIndex: 3 }, bridge);
    const parsed = JSON.parse(result.content[0].text);
    expect(parsed.success).toBe(true);
  });

  test('accept forwards valid index to bridge', async () => {
    const bridge = mockBridge({ acceptTrackedChange: { success: true } });
    await acceptHandler({ index: 0 }, bridge);
    expect(bridge.calls[0].action).toBe('acceptTrackedChange');
    expect(bridge.calls[0].params.index).toBe(0);
  });

  test('reject forwards valid index to bridge', async () => {
    const bridge = mockBridge({ rejectTrackedChange: { success: true } });
    await rejectHandler({ index: 2 }, bridge);
    expect(bridge.calls[0].action).toBe('rejectTrackedChange');
    expect(bridge.calls[0].params.index).toBe(2);
  });
});
