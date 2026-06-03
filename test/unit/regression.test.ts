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
      if (action === 'getParagraphs') return { count: 5, paragraphs: [] };
      if (action === 'getParaOoxml') return { ooxml: '<pkg:package></pkg:package>' };
      if (action === 'getContentControls') return { count: 0, controls: [] };
      return { success: true };
    },
  };
}

/** Shorthand mock with just paragraph count */
function mockBridgeWithParas(count = 5): any {
  return mockBridge({ getParagraphs: { count, paragraphs: [] } });
}

// =============================================================================
// search_and_replace rejects Word special find codes
// =============================================================================
describe('search_and_replace rejects Word special codes', () => {
  const handler = handlers.get('word_search_and_replace')!;

  test('rejects ^p in find string', async () => {
    await expect(handler({ find: '^p', replace: 'x' }, mockBridge()))
      .rejects.toThrow('Word special code');
  });

  test('rejects ^13 in find string', async () => {
    await expect(handler({ find: '^13', replace: 'x' }, mockBridge()))
      .rejects.toThrow('Word special code');
  });

  test('rejects ^w in find string', async () => {
    await expect(handler({ find: '^w', replace: 'x' }, mockBridge()))
      .rejects.toThrow('Word special code');
  });

  test('rejects ^t in find string', async () => {
    await expect(handler({ find: '^t', replace: 'x' }, mockBridge()))
      .rejects.toThrow('Word special code');
  });

  test('rejects ^11 in find string', async () => {
    await expect(handler({ find: '^11', replace: 'x' }, mockBridge()))
      .rejects.toThrow('Word special code');
  });

  test('rejects ^p in replace string', async () => {
    await expect(handler({ find: 'hello', replace: '^p' }, mockBridge()))
      .rejects.toThrow('Word special code');
  });

  test('rejects ^w in replace string', async () => {
    await expect(handler({ find: 'hello', replace: 'a^wb' }, mockBridge()))
      .rejects.toThrow('Word special code');
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

  test('skips duplicate check when using id directly', async () => {
    const bridge = mockBridge({ setContentControlText: { success: true } });
    const result = await handler({ id: 42, text: 'new' }, bridge);
    const parsed = JSON.parse(result.content[0].text);
    expect(parsed.success).toBe(true);
    expect(bridge.calls.find((c: any) => c.action === 'getContentControls')).toBeUndefined();
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
      .rejects.toThrow('exceeds maximum');
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
        if (action === 'getParagraphs') return { count: n, paragraphs };
        return { ooxml: '<pkg:package></pkg:package>' };
      },
    };
  }

  test('returns warning when moving range immediately after itself', async () => {
    const result = await moveHandler({ fromIndex: 0, toIndex: 3, count: 3 }, bridgeWithNParas(5));
    const parsed = JSON.parse(result.content[0].text);
    expect(parsed.warning).toContain('No move performed');
    expect(parsed.moved).toBeNull();
  });

  test('returns warning for single paragraph no-op (fromIndex=0, toIndex=1, After)', async () => {
    const result = await moveHandler({ fromIndex: 0, toIndex: 1 }, bridgeWithNParas(5));
    const parsed = JSON.parse(result.content[0].text);
    expect(parsed.warning).toContain('No move performed');
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
