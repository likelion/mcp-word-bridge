# MCP Word Bridge

[![Tests](https://github.com/likelion/mcp-word-bridge/actions/workflows/tests.yml/badge.svg)](https://github.com/likelion/mcp-word-bridge/actions/workflows/tests.yml)
[![codecov](https://codecov.io/gh/likelion/mcp-word-bridge/branch/main/graph/badge.svg)](https://codecov.io/gh/likelion/mcp-word-bridge)

MCP server for live Word document editing via Office Add-in. Enables programmatic editing of Word documents through the Word JavaScript API, with changes appearing as user edits in co-authoring sessions.

## Quick Start

### 1. Add to your MCP client

```json
{
  "mcpServers": {
    "word-bridge": {
      "command": "npx",
      "args": ["-y", "mcp-word-bridge"]
    }
  }
}
```

### 2. Install the Word add-in

```bash
npx mcp-word-bridge-install
```

Restart Word → Home → Add-ins → **MCP Word Bridge**

### 3. Trust the TLS certificate (first time only)

On first run the server auto-generates a self-signed localhost certificate and prints the trust command. Run it once:

```bash
# macOS
security add-trusted-cert -r trustRoot -k ~/Library/Keychains/login.keychain-db "$(npm root -g)/mcp-word-bridge/certs/cert.pem"

# Windows
certutil -user -addstore Root "%APPDATA%\npm\node_modules\mcp-word-bridge\certs\cert.pem"
```

### 4. Open Word and activate

Open Word → Home → Add-ins → MCP Word Bridge. The taskpane shows "Word: connected ✓" and "WebSocket: connected ✓".

That's it. The MCP server starts automatically when your MCP client loads the config, and stops when it unloads.

## Environment Variables

| Variable | Default | Description |
|----------|---------|-------------|
| `MCP_WORD_BRIDGE_PORT` | `3000` | HTTPS port for the bridge server |

## How It Works

```
MCP Client ←stdio→ Server ←WebSocket→ Taskpane (Office Add-in) ←→ Word JS API ←→ Document
```

Single process. The MCP client spawns the server, which starts both the HTTPS bridge (for the add-in) and the MCP protocol handler (on stdio). Everything starts and stops together.

## Tools (92)

### Document
| Tool | Description |
|------|-------------|
| `word_get_text` | Get full plain text of the document |
| `word_get_document_properties` | Get metadata (title, author, path, timestamps) |
| `word_set_document_properties` | Set metadata fields |
| `word_save` | Save document to disk |
| `word_clear` | Clear all document body content |
| `word_create_document` | Create and open a new blank document in a new Word window |
| `word_get_word_count` | Get word, character, and paragraph counts |
| `word_get_styles` | List available styles |
| `word_get_coauthors` | Get co-authoring status and active authors |
| `word_set_change_tracking` | Enable/disable track changes |
| `word_get_document_outline` | Get heading hierarchy as a structured outline tree |

### Paragraphs
| Tool | Description |
|------|-------------|
| `word_get_paragraphs` | Get paragraphs with style, alignment, TOC flag (paginated) |
| `word_get_paragraph_by_index` | Get full details of one paragraph (font, spacing, indent) |
| `word_insert_paragraph` | Append/prepend a styled paragraph at document Start or End |
| `word_insert_paragraph_at_index` | Insert a paragraph before or after a specific paragraph index |
| `word_delete_paragraph` | Delete a paragraph by index |
| `word_replace_paragraph_text` | Replace paragraph text by index (preserves style) |
| `word_move_paragraph` | Move one or more consecutive paragraphs to another position |
| `word_copy_paragraph` | Copy one or more consecutive paragraphs to another position |
| `word_set_paragraph_style` | Change style or alignment of a paragraph |
| `word_set_paragraph_spacing` | Set line spacing, before/after, and indentation |

### Search & Text
| Tool | Description |
|------|-------------|
| `word_search` | Find text matches (case-insensitive by default) |
| `word_search_and_replace` | Find and replace all occurrences |
| `word_insert_text_at_match` | Insert text before or after a search match |
| `word_get_selection_info` | Get current selection with font details |
| `word_insert_text_at_selection` | Insert or replace text at cursor |
| `word_insert_line_break` | Insert a soft line break (Shift+Enter) |

### Formatting
| Tool | Description |
|------|-------------|
| `word_format_text` | Apply bold, italic, color, size, font to a text match |
| `word_clear_formatting` | Reset formatting to paragraph style defaults |
| `word_get_font_info` | Inspect font properties of a text match |

### Tables
| Tool | Description |
|------|-------------|
| `word_insert_table` | Insert a table with data and optional style |
| `word_list_tables` | List table metadata (count, dimensions, style) |
| `word_get_table_data` | Get cell values from a specific table |
| `word_set_table_cell` | Set text in a cell |
| `word_add_table_row` | Add a row with optional values |
| `word_delete_table_row` | Delete a row by index |
| `word_merge_table_cells` | Merge a rectangular range of cells |
| `word_split_table_cell` | Split a cell into rows/columns |
| `word_set_table_style` | Apply a built-in table style |
| `word_set_table_cell_shading` | Set cell background color |

### Lists
| Tool | Description |
|------|-------------|
| `word_insert_list` | Insert a bulleted or numbered list |
| `word_get_list_info` | Get list level and numbering for a paragraph |
| `word_set_list_level` | Change indent level of a list item |

### Comments
| Tool | Description |
|------|-------------|
| `word_add_comment` | Add a comment anchored to text |
| `word_get_comments` | Get all comments with author, date, status, and anchor text |
| `word_get_comment_replies` | Get replies for a comment |
| `word_reply_to_comment` | Reply to a comment |
| `word_resolve_comment` | Mark a comment as resolved |
| `word_delete_comment` | Delete a comment and its replies |

### Footnotes & Endnotes
| Tool | Description |
|------|-------------|
| `word_insert_footnote` | Insert footnote anchored to a text match |
| `word_insert_footnote_at_index` | Insert footnote at end of a paragraph |
| `word_insert_endnote` | Insert endnote anchored to a text match |
| `word_get_footnotes` | Get all footnotes with index and text |
| `word_get_endnotes` | Get all endnotes with index and text |
| `word_delete_footnote` | Delete a footnote by index |
| `word_delete_endnote` | Delete an endnote by index |

### Track Changes
| Tool | Description |
|------|-------------|
| `word_get_tracked_changes` | Get all tracked changes with type, author, text |
| `word_accept_tracked_change` | Accept a change by index |
| `word_reject_tracked_change` | Reject a change by index |
| `word_accept_all_tracked_changes` | Accept all changes |
| `word_reject_all_tracked_changes` | Reject all changes |

### Content Controls
| Tool | Description |
|------|-------------|
| `word_get_content_controls` | Get all controls with id, tag, type, text |
| `word_insert_content_control` | Wrap text in a RichText, PlainText, or CheckBox control |
| `word_set_content_control_text` | Set text in a control by ID or tag |

### Bookmarks
| Tool | Description |
|------|-------------|
| `word_get_bookmarks` | List all bookmark names |
| `word_insert_bookmark` | Create a bookmark at a text match |
| `word_delete_bookmark` | Delete a bookmark by name |
| `word_go_to_bookmark` | Navigate to a bookmark and select it |
| `word_get_bookmark_text` | Get text content of a bookmark |

### Hyperlinks
| Tool | Description |
|------|-------------|
| `word_insert_hyperlink` | Add a hyperlink to existing text |
| `word_get_hyperlinks` | List all hyperlinks with URLs and text |
| `word_remove_hyperlink` | Remove link from text (keeps the text) |

### Headers & Footers
| Tool | Description |
|------|-------------|
| `word_get_header_footer` | Get header or footer text |
| `word_set_header_footer` | Set header or footer text |

### Images
| Tool | Description |
|------|-------------|
| `word_insert_image` | Insert image from base64 data |
| `word_get_images` | List images with dimensions and alt text |
| `word_delete_image` | Delete an image by index |

### Page Layout & Sections
| Tool | Description |
|------|-------------|
| `word_get_page_layout` | Get margins, orientation, paper size |
| `word_set_page_layout` | Set margins, orientation, paper size |
| `word_get_sections` | List all sections with page setup |
| `word_insert_page_break` | Insert a page break after a paragraph |
| `word_insert_section_break` | Insert a section break after a paragraph |
| `word_get_page_info` | Get page count and per-page paragraph ranges (Desktop only) |

### Custom Properties
| Tool | Description |
|------|-------------|
| `word_get_custom_properties` | Get all custom key-value properties |
| `word_set_custom_property` | Create or update a custom property |
| `word_delete_custom_property` | Delete a custom property |

### Advanced
| Tool | Description |
|------|-------------|
| `word_insert_html` | Insert HTML converted to native Word formatting |
| `word_insert_ooxml` | Insert raw Office Open XML |
| `word_insert_table_of_contents` | Insert a TOC based on headings |
| `word_get_fields` | Get all fields (hyperlinks, TOC, page numbers) |

### Equations
| Tool | Description |
|------|-------------|
| `word_insert_equation` | Insert a LaTeX equation as a native editable Word equation |

### Batch
| Tool | Description |
|------|-------------|
| `word_batch` | Execute up to 50 operations in a single call |

## Equations

`word_insert_equation` converts LaTeX math to native Word equations:

```
LaTeX → temml → MathML → mathml2omml → OMML → OOXML → Word
```

- **Display mode** (default): centered block equation
- **Inline mode** (`displayMode: false`): within a paragraph, optionally positioned via `anchorText`
- Supports: fractions, roots, integrals, sums, matrices, Greek letters, AMS math
- Equations are fully editable in Word's built-in equation editor

## Resources

The server exposes an MCP resource at `word-bridge://usage-guide` containing patterns and best practices for LLMs operating on Word documents.

## TLS Certificate

Auto-generated on first run. The server prints the trust command:

```
[mcp-word-bridge] ✓ Certificate generated.
[mcp-word-bridge] Trust it: security add-trusted-cert -r trustRoot -k ~/Library/Keychains/login.keychain-db "/path/to/certs/cert.pem"
```

Run it once. To regenerate: delete `certs/cert.pem` and `certs/key.pem`, then restart.

## Development

```bash
git clone https://github.com/likelion/mcp-word-bridge.git
cd mcp-word-bridge
npm install
npm run build        # build server + taskpane
npm run dev          # watch mode
npm run typecheck    # TypeScript type checking
npm test             # unit tests (67 tests, <500ms)
npm run test:live    # integration tests (requires Word)
```

### Project Structure

```
src/
├── server/           # MCP server (Node.js)
│   ├── tools/        # One file per tool category
│   ├── bridge.ts     # WebSocket bridge to taskpane
│   ├── mcp.ts        # MCP protocol handlers
│   └── validation.ts # Shared input validators
├── taskpane/         # Runs inside Word's add-in webview
│   ├── commands/     # One file per command category
│   └── log.ts        # Taskpane UI logger
├── shared/           # Types shared between server and taskpane
│   ├── protocol.ts   # WebSocket message types
│   └── constants.ts  # Limits and timeouts
└── lib/
    └── equations.ts  # LaTeX→OMML conversion pipeline
```

## License

MIT
