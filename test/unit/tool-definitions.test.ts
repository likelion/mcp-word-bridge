import { describe, test, expect } from 'vitest';
import { buildToolRegistry } from '../../src/server/tools';

describe('Tool Registry', () => {
  const { tools, handlers } = buildToolRegistry();

  test('declares exactly 91 tools', () => {
    expect(tools.length).toBe(91);
  });

  test('every tool has name, description, and schema', () => {
    for (const t of tools) {
      expect(t.name).toBeTruthy();
      expect(t.description).toBeTruthy();
      expect(t.schema).toBeDefined();
      expect(t.schema.properties).toBeDefined();
    }
  });

  test('all tool names follow word_ prefix', () => {
    for (const t of tools) {
      expect(t.name).toMatch(/^word_/);
    }
  });

  test('no duplicate tool names', () => {
    const names = tools.map(t => t.name);
    expect(new Set(names).size).toBe(names.length);
  });

  test('required fields reference defined properties', () => {
    for (const t of tools) {
      if (t.schema.required) {
        for (const r of t.schema.required) {
          expect(t.schema.properties).toHaveProperty(r);
        }
      }
    }
  });

  test('every tool has a handler in the registry', () => {
    for (const t of tools) {
      expect(handlers.has(t.name)).toBe(true);
    }
  });

  test('all descriptions have category tags', () => {
    for (const t of tools) {
      expect(t.description).toMatch(/^\[.+\]/);
    }
  });

  test('word_move_paragraph requires fromIndex and toIndex', () => {
    const t = tools.find(t => t.name === 'word_move_paragraph')!;
    expect(t.schema.required).toContain('fromIndex');
    expect(t.schema.required).toContain('toIndex');
  });

  test('word_copy_paragraph requires fromIndex and toIndex', () => {
    const t = tools.find(t => t.name === 'word_copy_paragraph')!;
    expect(t.schema.required).toContain('fromIndex');
    expect(t.schema.required).toContain('toIndex');
  });

  test('word_batch requires operations', () => {
    const t = tools.find(t => t.name === 'word_batch')!;
    expect(t.schema.required).toContain('operations');
  });

  test('word_insert_equation requires latex', () => {
    const t = tools.find(t => t.name === 'word_insert_equation')!;
    expect(t.schema.required).toContain('latex');
  });

  test('forward tools have bridgeAction property', () => {
    const forwardTools = tools.filter((t: any) => 'bridgeAction' in t);
    expect(forwardTools.length).toBeGreaterThan(80); // Most tools are forwards
    for (const t of forwardTools) {
      expect(typeof (t as any).bridgeAction).toBe('string');
      expect((t as any).bridgeAction.length).toBeGreaterThan(0);
    }
  });

  test('server-composed tools do NOT have bridgeAction', () => {
    const serverComposed = ['word_move_paragraph', 'word_copy_paragraph', 'word_get_document_outline', 'word_insert_equation', 'word_batch'];
    for (const name of serverComposed) {
      const t = tools.find(t => t.name === name)!;
      expect('bridgeAction' in t).toBe(false);
    }
  });

  test('word_search exposes all search options', () => {
    const t = tools.find(t => t.name === 'word_search')!;
    const props = Object.keys(t.schema.properties);
    expect(props).toContain('matchCase');
    expect(props).toContain('matchWholeWord');
    expect(props).toContain('matchWildcards');
    expect(props).toContain('matchPrefix');
    expect(props).toContain('matchSuffix');
    expect(props).toContain('ignorePunct');
    expect(props).toContain('ignoreSpace');
  });

  test('word_search_and_replace exposes all search options', () => {
    const t = tools.find(t => t.name === 'word_search_and_replace')!;
    const props = Object.keys(t.schema.properties);
    expect(props).toContain('matchWildcards');
    expect(props).toContain('matchPrefix');
    expect(props).toContain('matchSuffix');
    expect(props).toContain('ignorePunct');
    expect(props).toContain('ignoreSpace');
  });
});
