export const usageGuide = `# MCP Word Bridge — Usage Guide

Controls a live Word document. All operations execute immediately.

## Quick Start

1. **Read before writing** — call \`word_get_document_outline\` or \`word_get_paragraphs\` to understand structure
2. **Use the right tool:**
   - Append content → \`word_insert_paragraph\` (Start/End)
   - Insert at position → \`word_insert_paragraph_at_index\` (Before/After by index)
   - Edit existing text → \`word_replace_paragraph_text\` (by index, preferred)
   - Bulk find/replace → \`word_search_and_replace\` (all occurrences)
   - Insert adjacent to text → \`word_insert_text_at_match\` (searches then inserts)
3. **Batch operations** — use \`word_batch\` for multiple operations in one call (faster, fewer round-trips)
4. **Save explicitly** — call \`word_save\` after significant changes

## Reading

- \`word_get_document_outline\` — heading tree (fast structural overview)
- \`word_get_paragraphs\` — paragraphs with text, style, alignment. Paginate with start/end.
- \`word_get_paragraph_by_index\` — full details of one paragraph (font, spacing, indent)
- \`word_get_text\` — plain-text dump (no structure)
- \`word_search\` — locate text before operating on it

## Editing

- \`word_replace_paragraph_text\` — replace by index (safe for collaboration, preserves style)
- \`word_search_and_replace\` — bulk find/replace across document
- \`word_insert_text_at_match\` — insert before/after a search match (use \`occurrence\` for Nth)
- \`word_move_paragraph\` — reorder content (preserves footnotes, formatting, hyperlinks)
- \`word_copy_paragraph\` — duplicate content with full fidelity
- Verify edits with \`word_search\` or \`word_get_paragraphs\`

## Batch

\`\`\`json
{"operations": [
  {"tool": "word_insert_paragraph", "args": {"text": "Hello", "style": "Heading 1"}},
  {"tool": "word_insert_paragraph", "args": {"text": "World"}},
  {"tool": "word_format_text", "args": {"text": "Hello", "bold": true}}
]}
\`\`\`
Runs sequentially. Stops on first error. Maximum 50 per batch. Prefer batching over individual calls.

## Search

- Case-insensitive by default. Pass \`matchCase: true\` for exact case.
- Text parameters must be ≤255 characters.
- Use \`occurrence\` (0-indexed) to target the Nth match when multiple exist.

## Comments — Important

- \`word_get_comments\` returns comments with their anchor text
- **Replacing text that anchors a comment deletes the comment** — always check comments first
- Safe pattern: \`word_get_comments\` → \`word_reply_to_comment\`/\`word_resolve_comment\` → then replace

## Tables

- All indices 0-based: tableIndex, row, col
- \`word_list_tables\` for metadata → \`word_get_table_data\` for cell values
- Cannot insert page/section breaks inside table cells

## Footnotes & Endnotes

- \`word_insert_footnote\` — anchor to a text match
- \`word_insert_footnote_at_index\` — anchor to a paragraph by index (no search needed)

## Page Layout

- Margins in points (72 pt = 1 inch)
- \`lineSpacing\` is in points, not a multiplier (12pt font: 12=single, 18=1.5x, 24=double)

## Content Controls

- RichText/PlainText: wraps anchor text (non-destructive)
- CheckBox: REPLACES anchor text with a checkbox (cannot set text after)

## TOC

After inserting a Table of Contents, heading text appears twice (in TOC and body). Search matches TOC entries first — use \`occurrence\` to target the body instance.

## Equations

- \`word_insert_equation\` takes LaTeX, inserts a native editable Word equation
- \`displayMode: true\` (default) = centered block equation
- \`displayMode: false\` = inline; provide \`anchorText\` to position after a match
- Supports: fractions, roots, integrals, sums, matrices, Greek letters, AMS math
- Examples: \`\\\\frac{a}{b}\`, \`\\\\int_0^\\\\infty e^{-x} dx\`, \`\\\\sum_{i=1}^n x_i\`

## Change Tracking

- Call \`word_set_change_tracking({mode:"TrackAll"})\` BEFORE edits to track them
- Tracked edits appear as revisions in Word (visible to collaborators)

## Best Practices

1. Read structure first (\`word_get_document_outline\`)
2. Check comments before bulk replacements (\`word_get_comments\`)
3. Batch multiple operations (\`word_batch\`)
4. Prefer index-based tools over search-based in collaborative editing
5. Use \`word_move_paragraph\` to reorder (not delete+insert)
6. Use \`word_copy_paragraph\` to duplicate (preserves everything)
7. Save explicitly after significant changes
8. Resolve comments rather than deleting (preserves audit trail)
`;
