/**
 * Word special-code expansion for replacement/insertion strings.
 *
 * The Word desktop Find/Replace UI interprets codes like ^p (paragraph mark),
 * ^t (tab), ^l (line break), ^s (non-breaking space) in replacement strings.
 * The Word JavaScript API's insertText() does NOT interpret these — it inserts
 * them as literal text. This module bridges the gap by parsing codes and
 * executing the corresponding API operations.
 *
 * Supported codes (structural — require API calls):
 *   ^p  → paragraph break (insertParagraph)
 *   ^l  → line break / soft return (insertBreak 'Line')
 *   ^m  → page break (insertBreak 'Page')
 *   ^n  → column break (insertBreak 'Column')
 *
 * Supported codes (character — inserted via insertText):
 *   ^t  → tab (\t)
 *   ^s  → non-breaking space (\u00A0)
 *   ^~  → non-breaking hyphen (\u2011)
 *   ^-  → optional/soft hyphen (\u00AD)
 *   ^+  → em dash (\u2014)
 *   ^=  → en dash (\u2013)
 *   ^^  → literal caret (^)
 */

declare const Word: any;

export interface TextSegment {
  type: 'text' | 'paragraph' | 'lineBreak' | 'pageBreak' | 'columnBreak';
  content?: string; // only for 'text' type
}

// Placeholder for escaped caret (^^) during parsing
const CARET_PLACEHOLDER = '\u0000';

/**
 * Parse a string containing Word special codes into typed segments.
 * Returns an ordered list of segments to execute sequentially.
 */
export function parseSpecialCodes(input: string): TextSegment[] {
  if (!input) return [{ type: 'text', content: '' }];

  // Pass 1: Replace ^^ with placeholder to avoid interference
  const working = input.replace(/\^\^/g, CARET_PLACEHOLDER);

  // Pass 2: Split on recognized codes
  // Single-letter codes: ^p, ^l, ^t, ^s, ^m, ^n (case-insensitive)
  // Symbol codes: ^~, ^-, ^+, ^=
  const segments: TextSegment[] = [];
  const codePattern = /\^([pltsmnPLTSMN~\-+=])/g;
  let lastIndex = 0;
  let match: RegExpExecArray | null;

  while ((match = codePattern.exec(working)) !== null) {
    // Collect text before this code
    if (match.index > lastIndex) {
      const text = working.slice(lastIndex, match.index);
      segments.push({ type: 'text', content: restoreCaret(text) });
    }

    // Emit the code segment
    const code = match[1]!;
    switch (code.toLowerCase()) {
      case 'p':
        segments.push({ type: 'paragraph' });
        break;
      case 'l':
        segments.push({ type: 'lineBreak' });
        break;
      case 'm':
        segments.push({ type: 'pageBreak' });
        break;
      case 'n':
        segments.push({ type: 'columnBreak' });
        break;
      case 't':
        segments.push({ type: 'text', content: '\t' });
        break;
      case 's':
        segments.push({ type: 'text', content: '\u00A0' });
        break;
      default:
        // Symbol codes (not case-sensitive — they're symbols)
        switch (code) {
          case '~':
            segments.push({ type: 'text', content: '\u2011' }); // non-breaking hyphen
            break;
          case '-':
            segments.push({ type: 'text', content: '\u00AD' }); // optional/soft hyphen
            break;
          case '+':
            segments.push({ type: 'text', content: '\u2014' }); // em dash
            break;
          case '=':
            segments.push({ type: 'text', content: '\u2013' }); // en dash
            break;
        }
    }

    lastIndex = match.index + match[0].length;
  }

  // Remaining text after last code
  if (lastIndex < working.length) {
    const text = working.slice(lastIndex);
    segments.push({ type: 'text', content: restoreCaret(text) });
  }

  // If no segments were produced (empty input edge case), return empty text
  if (segments.length === 0) {
    return [{ type: 'text', content: '' }];
  }

  return segments;
}

/** Restore caret placeholder back to literal ^ */
function restoreCaret(text: string): string {
  return text.replace(/\u0000/g, '^');
}

/**
 * Check if a string contains any special codes that need expansion.
 * Used as a fast-path check to skip parsing for plain text.
 */
export function hasSpecialCodes(text: string): boolean {
  // Match ^ followed by a recognized code letter/symbol, but not ^^
  return /\^(?!\^)[pltsmn~\-+=]/i.test(text) || /\^\^/.test(text);
}

/**
 * Insert text with special-code expansion into a Word range.
 *
 * For plain text (no codes), behaves identically to range.insertText().
 * For text with codes, executes a sequence of API operations to produce
 * the correct document structure.
 *
 * @param range - The Word range to insert relative to
 * @param text - Text potentially containing special codes
 * @param location - Word.InsertLocation (replace, before, after)
 * @param ctx - Word request context for sync
 * @returns The range covering the end of the inserted content
 */
export async function insertWithCodes(
  range: any,
  text: string,
  location: string,
  ctx: any,
): Promise<any> {
  // Fast path: no special codes — plain insertText
  if (!hasSpecialCodes(text)) {
    const inserted = range.insertText(text, location);
    await ctx.sync();
    return inserted;
  }

  const segments = parseSpecialCodes(text);

  // If after parsing we have a single text segment, use plain insertText
  if (segments.length === 1 && segments[0]!.type === 'text') {
    const inserted = range.insertText(segments[0]!.content!, location);
    await ctx.sync();
    return inserted;
  }

  // Multi-segment insertion:
  // For replace mode, first clear the matched range then insert after it.
  // For before/after, first segment uses the original location.
  // Subsequent segments always chain via 'after' from the cursor.
  let cursor: any = range;
  let isFirst = true;

  for (const seg of segments) {
    // On the first segment with replace, clear the range and switch to 'after'
    if (isFirst && location === Word.InsertLocation.replace) {
      cursor = cursor.insertText('', Word.InsertLocation.replace);
      await ctx.sync();
      // From here on, everything goes 'after' the now-empty cursor
    }

    const loc = isFirst && location !== Word.InsertLocation.replace
      ? location
      : Word.InsertLocation.after;

    switch (seg.type) {
      case 'text': {
        if (seg.content) {
          cursor = cursor.insertText(seg.content, loc);
          await ctx.sync();
        }
        break;
      }
      case 'paragraph': {
        const para = cursor.insertParagraph('', loc);
        await ctx.sync();
        // Get the end of the new paragraph as our cursor
        cursor = para.getRange(Word.RangeLocation.end);
        await ctx.sync();
        break;
      }
      case 'lineBreak': {
        cursor.insertBreak(Word.BreakType.line, loc);
        await ctx.sync();
        cursor = cursor.getRange(Word.RangeLocation.after);
        await ctx.sync();
        break;
      }
      case 'pageBreak': {
        cursor.insertBreak(Word.BreakType.page, loc);
        await ctx.sync();
        cursor = cursor.getRange(Word.RangeLocation.after);
        await ctx.sync();
        break;
      }
      case 'columnBreak': {
        cursor.insertBreak(Word.BreakType.column, loc);
        await ctx.sync();
        cursor = cursor.getRange(Word.RangeLocation.after);
        await ctx.sync();
        break;
      }
    }

    isFirst = false;
  }

  return cursor;
}
