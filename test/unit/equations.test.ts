import { describe, test, expect } from 'vitest';
import { fixDelimiters, fixNaryOperands, latexToOmml, buildEquationOoxml } from '../../src/lib/equations';

describe('latexToOmml', () => {
  const { mml2omml } = require('mathml2omml');

  test('converts simple fraction', () => {
    const result = latexToOmml('\\frac{a}{b}', true, mml2omml);
    expect(result.omml).toContain('<m:f>');
    expect(result.displayMode).toBe(true);
  });

  test('converts superscript', () => {
    const result = latexToOmml('x^2', true, mml2omml);
    expect(result.omml).toContain('<m:sSup>');
  });

  test('throws on invalid LaTeX', () => {
    expect(() => latexToOmml('\\frac{', true, mml2omml)).toThrow('LaTeX parse error');
  });

  test('throws on empty input', () => {
    expect(() => latexToOmml('', true, mml2omml)).toThrow('"latex"');
  });

  test('strips xmlns declarations', () => {
    const result = latexToOmml('x', true, mml2omml);
    expect(result.omml).not.toContain('xmlns:m=');
  });

  test('escapes < and & in text nodes', () => {
    const result = latexToOmml('a < b', true, mml2omml);
    expect(result.omml).not.toMatch(/<m:t[^>]*>[^<]*<[^/]/);
  });
});

describe('buildEquationOoxml', () => {
  test('display mode wraps in oMathPara with centered paragraph', () => {
    const ooxml = buildEquationOoxml('<m:r><m:t>x</m:t></m:r>', true);
    expect(ooxml).toContain('<m:oMathPara>');
    expect(ooxml).toContain('w:val="center"');
  });

  test('display mode includes trailing Normal paragraph reset', () => {
    const ooxml = buildEquationOoxml('<m:r><m:t>x</m:t></m:r>', true);
    expect(ooxml).toContain('w:val="Normal"');
    expect(ooxml).toContain('w:val="left"');
  });

  test('inline mode does not wrap in oMathPara', () => {
    const ooxml = buildEquationOoxml('<m:r><m:t>x</m:t></m:r>', false);
    expect(ooxml).not.toContain('<m:oMathPara>');
  });

  test('output is valid OOXML package structure', () => {
    const ooxml = buildEquationOoxml('<m:r><m:t>x</m:t></m:r>', true);
    expect(ooxml).toContain('pkg:package');
    expect(ooxml).toContain('pkg:part');
    expect(ooxml).toContain('word/document.xml');
  });
});

describe('fixDelimiters', () => {
  test('wraps matched parentheses in m:d', () => {
    const input = '<m:r><m:t xml:space="preserve">(</m:t></m:r><m:r><m:t>x</m:t></m:r><m:r><m:t xml:space="preserve">)</m:t></m:r>';
    const result = fixDelimiters(input);
    expect(result).toContain('<m:d>');
    expect(result).toContain('<m:e>');
  });
});

describe('fixNaryOperands', () => {
  test('returns unchanged input when no nary operators', () => {
    const input = '<mrow><mi>x</mi></mrow>';
    expect(fixNaryOperands(input)).toBe(input);
  });
});
