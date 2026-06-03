/**
 * Equation conversion pipeline: LaTeX → MathML → OMML → OOXML.
 * Pure functions, no I/O, no server dependencies.
 */

export const DELIM_PAIRS: Record<string, string> = {
  '(': ')', '[': ']', '{': '}', '|': '|', '\u2016': '\u2016',
  '\u2308': '\u2309', '\u230A': '\u230B', '\u27E8': '\u27E9',
};
export const OPEN_CHARS = new Set(Object.keys(DELIM_PAIRS));

interface DelimInfo { index: number; length: number; char: string }
interface DelimPair { open: DelimInfo; close: DelimInfo }

export function fixDelimiters(omml: string): string {
  const DELIM_CHARS_ALWAYS = '()[]{}\u2308\u2309\u230A\u230B\u27E8\u27E9';
  const DELIM_CHARS_BOUNDARY = '|\u2016';

  omml = omml.replace(
    /<m:r>(<m:rPr>[^]*?<\/m:rPr>)?<m:t([^>]*)>([^<]+)<\/m:t><\/m:r>/g,
    (match, rPr: string | undefined, attrs: string, text: string) => {
      if (text.length <= 1) return match;
      const allDelims = DELIM_CHARS_ALWAYS + DELIM_CHARS_BOUNDARY;
      const hasDelim = [...text].some(c => allDelims.includes(c));
      if (!hasDelim) return match;
      const rPrStr = rPr || '';
      const segments: string[] = [];
      let current = '';
      const chars = [...text];
      for (let i = 0; i < chars.length; i++) {
        const ch = chars[i]!;
        if (DELIM_CHARS_ALWAYS.includes(ch)) {
          if (current) segments.push(current);
          segments.push(ch);
          current = '';
        } else if (DELIM_CHARS_BOUNDARY.includes(ch) && (i === 0 || i === chars.length - 1)) {
          if (current) segments.push(current);
          segments.push(ch);
          current = '';
        } else {
          current += ch;
        }
      }
      if (current) segments.push(current);
      if (segments.length <= 1) return match;
      return segments.map(s => '<m:r>' + rPrStr + '<m:t' + attrs + '>' + s + '</m:t></m:r>').join('');
    },
  );

  const delimRun = /<m:r><m:t xml:space="preserve">([()[\]{}|\u2016\u2308\u2309\u230A\u230B\u27E8\u27E9]|)<\/m:t><\/m:r>/g;
  const delims: DelimInfo[] = [];
  let m;
  while ((m = delimRun.exec(omml)) !== null) {
    delims.push({ index: m.index, length: m[0].length, char: m[1]! });
  }
  if (delims.length < 2) return omml;

  const pairs: DelimPair[] = [];
  const stack: DelimInfo[] = [];
  for (const d of delims) {
    if (d.char === '|' || d.char === '\u2016') {
      const stackIdx = stack.findIndex(s => s.char === d.char);
      if (stackIdx >= 0) {
        pairs.push({ open: stack[stackIdx]!, close: d });
        stack.splice(stackIdx, 1);
        continue;
      }
    }
    if (OPEN_CHARS.has(d.char)) {
      stack.push(d);
    } else {
      const expectedOpen = d.char === ')' ? '(' : d.char === ']' ? '[' : d.char === '}' ? '{' :
        d.char === '\u2309' ? '\u2308' : d.char === '\u230B' ? '\u230A' : d.char === '\u27E9' ? '\u27E8' : null;
      for (let i = stack.length - 1; i >= 0; i--) {
        if (stack[i]!.char === expectedOpen) {
          pairs.push({ open: stack[i]!, close: d });
          stack.splice(i, 1);
          break;
        }
      }
    }
  }

  pairs.sort((a, b) => b.open.index - a.open.index);
  for (const pair of pairs) {
    const { open, close } = pair;
    const openMatch = omml.indexOf('<m:r><m:t xml:space="preserve">' + open.char + '</m:t></m:r>', open.index > 10 ? open.index - 10 : 0);
    const closeSearch = open.char === close.char ? openMatch + 50 : 0;
    const closeMatch = omml.indexOf('<m:r><m:t xml:space="preserve">' + close.char + '</m:t></m:r>', closeSearch > openMatch ? closeSearch : openMatch + 1);
    if (openMatch < 0 || closeMatch < 0 || closeMatch <= openMatch) continue;

    const openLen = ('<m:r><m:t xml:space="preserve">' + open.char + '</m:t></m:r>').length;
    const closeLen = ('<m:r><m:t xml:space="preserve">' + close.char + '</m:t></m:r>').length;
    const content = omml.substring(openMatch + openLen, closeMatch);

    let dPrContent = '';
    if (open.char !== '(') dPrContent += '<m:begChr m:val="' + open.char + '"/>';
    if (close.char !== ')') dPrContent += '<m:endChr m:val="' + close.char + '"/>';
    const dPr = dPrContent ? '<m:dPr>' + dPrContent + '</m:dPr>' : '';
    const replacement = '<m:d>' + dPr + '<m:e>' + content + '</m:e></m:d>';
    omml = omml.substring(0, openMatch) + replacement + omml.substring(closeMatch + closeLen);
  }
  return omml;
}

interface NaryInfo { start: number; tag: string }

export function fixNaryOperands(mml: string): string {
  const NARY_CHARS = ['\u2211', '\u222B', '\u220F', '\u222E', '\u22C3', '\u22C2', '\u2A01', '\u2A02'];
  const naryStarts: NaryInfo[] = [];

  for (const ch of NARY_CHARS) {
    let searchFrom = 0;
    while (true) {
      const moIdx = mml.indexOf('>' + ch + '</mo>', searchFrom);
      if (moIdx < 0) break;
      const before = mml.substring(Math.max(0, moIdx - 200), moIdx);
      const msubMatch = before.match(/.*(<msub(?:sup)?>)<mo[^>]*$/);
      if (msubMatch) {
        const msubStart = moIdx - (before.length - before.lastIndexOf(msubMatch[1]!));
        naryStarts.push({ start: msubStart, tag: msubMatch[1]!.replace('<', '').replace('>', '') });
      }
      searchFrom = moIdx + 1;
    }
  }
  if (naryStarts.length === 0) return mml;
  naryStarts.sort((a, b) => b.start - a.start);

  for (const nary of naryStarts) {
    const tag = nary.tag;
    const closeTag = '</' + tag + '>';
    const openTag = '<' + tag + '>';
    let depth = 0;
    let i = nary.start;
    let closeEnd = -1;
    while (i < mml.length) {
      if (mml.startsWith(openTag, i)) { depth++; i += openTag.length; }
      else if (mml.startsWith(closeTag, i)) { depth--; if (depth === 0) { closeEnd = i + closeTag.length; break; } i += closeTag.length; }
      else { i++; }
    }
    if (closeEnd < 0) continue;

    const rest = mml.substring(closeEnd);
    if (!rest || rest.startsWith('</mrow>') || rest.startsWith('</math>') || rest.startsWith('<mrow>')) continue;

    let opDepth = 0;
    let opEnd = 0;
    let j = 0;
    while (j < rest.length) {
      if (rest.startsWith('<mrow', j)) { opDepth++; j = rest.indexOf('>', j) + 1; }
      else if (rest.startsWith('</mrow>', j)) { if (opDepth === 0) { opEnd = j; break; } opDepth--; j += 7; }
      else if (rest.startsWith('</math>', j)) { opEnd = j; break; }
      else if (opDepth === 0 && rest.startsWith('<mo', j)) {
        const moEnd = rest.indexOf('>', j);
        if (moEnd >= 0) {
          const moTag = rest.substring(j, moEnd + 1);
          if (moTag.includes('fence="true"') && moTag.includes('postfix')) { opEnd = j; break; }
          const moClose = rest.indexOf('</mo>', moEnd);
          if (moClose >= 0) {
            const moContent = rest.substring(moEnd + 1, moClose);
            if (['+', '\u2212', '=', '<', '>', '\u2264', '\u2265', '\u2260'].includes(moContent.trim())) { opEnd = j; break; }
          }
        }
        j = rest.indexOf('>', j) + 1;
      } else { j++; }
    }
    if (opEnd === 0 && j >= rest.length) opEnd = rest.length;
    if (opEnd <= 0) continue;

    const operand = rest.substring(0, opEnd);
    if (!operand.trim()) continue;
    mml = mml.substring(0, closeEnd) + '<mrow>' + operand + '</mrow>' + mml.substring(closeEnd + opEnd);
  }
  return mml;
}

type Mml2OmmlFn = (mml: string) => string;

export interface LatexResult {
  omml: string;
  displayMode: boolean;
}

/**
 * Convert LaTeX to OMML (full pipeline).
 */
export function latexToOmml(latex: string, displayMode: boolean, mml2omml: Mml2OmmlFn): LatexResult {
  if (!latex || typeof latex !== 'string') throw new Error('"latex" parameter is required and must be a string.');
  if (!mml2omml) throw new Error('mml2omml function is required');

  const temml = require('temml/dist/temml.cjs');
  let mathml: string;
  try {
    mathml = temml.renderToString(latex);
  } catch (e) {
    throw new Error('LaTeX parse error: ' + (e as Error).message);
  }
  if (mathml.includes('temml-error') || !mathml.startsWith('<math')) {
    const errMatch = mathml.match(/ParseError:\s*(.+?)(<|$)/);
    const errMsg = errMatch ? errMatch[1]!.trim() : 'Invalid LaTeX expression';
    throw new Error('LaTeX parse error: ' + errMsg);
  }

  // Preprocess MathML
  let cleanMml = mathml;
  cleanMml = cleanMml.replace(/<mspace[^>]*\/>/g, '');
  cleanMml = cleanMml.replace(/<mspace[^>]*><\/mspace>/g, '');
  cleanMml = cleanMml.replace(/ class="[^"]*"/g, '');
  cleanMml = cleanMml.replace(/<mrow style="[^"]*"><\/mrow>/g, '');
  cleanMml = cleanMml.replace(/[\uFE00-\uFE0F]/g, '');

  if (displayMode) {
    cleanMml = cleanMml.replace(/<msub><mi>(lim|lim sup|lim inf|min|max|inf|sup)<\/mi>/g, '<munder><mi>$1</mi>');
    cleanMml = cleanMml.replace(/(<munder><mi>(?:lim|lim sup|lim inf|min|max|inf|sup)<\/mi><[^]*?)<\/msub>/g, '$1</munder>');
  }

  cleanMml = fixNaryOperands(cleanMml);

  let omml: string;
  try {
    omml = mml2omml(cleanMml);
  } catch (e) {
    throw new Error('MathML→OMML conversion error: ' + (e as Error).message);
  }

  // Post-process OMML
  omml = omml.replace(/ xmlns:m="[^"]*"/g, '').replace(/ xmlns:w="[^"]*"/g, '');
  omml = omml.replace(/<m:e\/>/g, '');

  // Fix text node escaping
  const parts = omml.split(/(<m:t>|<m:t\s[^>]*>|<\/m:t>)/);
  let inText = false;
  for (let i = 0; i < parts.length; i++) {
    const part = parts[i]!;
    if (part === '<m:t>' || (part.startsWith('<m:t ') && part.endsWith('>'))) { inText = true; continue; }
    if (part === '</m:t>') { inText = false; continue; }
    if (inText && (part.includes('<') || part.includes('&'))) {
      parts[i] = part.replace(/&(?!amp;|lt;|gt;|quot;|apos;)/g, '&amp;').replace(/</g, '&lt;');
    }
  }
  omml = parts.join('');

  // Fix trailing spaces in \text{} runs
  omml = omml.replace(
    /(<m:r><m:rPr><m:nor\/><\/m:rPr><m:t[^>]*>)([^<]+)(<\/m:t><\/m:r><m:r>)/g,
    (match, prefix: string, text: string, suffix: string) => {
      if (!text.endsWith(' ')) return prefix + text + ' ' + suffix;
      return match;
    },
  );

  omml = fixDelimiters(omml);

  if (displayMode) {
    omml = omml.replace(/<m:limLoc m:val="subSup"\/>/g, '<m:limLoc m:val="undOvr"/>');
  }

  return { omml, displayMode };
}

/**
 * Wrap OMML in a complete OOXML flat OPC package ready for Word insertion.
 */
export function buildEquationOoxml(omml: string, displayMode: boolean): string {
  const mathContent = displayMode ? '<m:oMathPara>' + omml + '</m:oMathPara>' : omml;

  let bodyContent: string;
  if (displayMode) {
    bodyContent = '<w:p><w:pPr><w:jc w:val="center"/></w:pPr>' + mathContent + '</w:p>' +
      '<w:p><w:pPr><w:pStyle w:val="Normal"/><w:jc w:val="left"/><w:rPr>' +
      '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="24"/>' +
      '</w:rPr></w:pPr></w:p>';
  } else {
    bodyContent = '<w:p>' + mathContent +
      '<w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>' +
      '<w:sz w:val="24"/><w:i w:val="0"/></w:rPr><w:t>\u200B</w:t></w:r></w:p>';
  }

  return '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">' +
    '<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
    '</Relationships></pkg:xmlData></pkg:part>' +
    '<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData>' +
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">' +
    '<w:body>' + bodyContent + '</w:body></w:document></pkg:xmlData></pkg:part></pkg:package>';
}
