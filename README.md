> [English](./README.md) | [中文](./README.zh-CN.md)

# Word MCP Server

MCP server for real-time Microsoft Word document manipulation via COM automation (using winax native addon). Launches Word visibly so you can see changes as they happen.

> This tool is intended for users who already have a valid Microsoft Word license.
> It does not include, crack, or distribute Microsoft Office software.

## Features

- **110+ MCP tools** — document lifecycle, content input, formatting, tables, charts, media, document structure, selection & clipboard, high-level Manager API, semantic navigation, variable replacement
- **Real-time visible** — Word window opens in foreground, every change is immediately visible
- **Ctrl+Z friendly** — all operations use Word COM API, undo stack works normally
- **Disconnect-safe** — MCP disconnection keeps Word open without losing edits
- **Auto backup** — `.bak` file created before every save
- **Chart data isolation** — child process with 15s timeout prevents Excel COM freeze
- **Format auto-detection** — file extension determines save format (docx/pdf/rtf/txt...)

## Prerequisites

- Windows OS with Microsoft Word installed (Office 2016+)
- Node.js 20+
- MSVC Build Tools (required only during `npm install` to compile winax)

## Quick Start

```bash
cd word-mcp-server
npm install && npm run build
```

Add to your MCP client config:

```json
{
  "mcpServers": {
    "word": {
      "command": "node",
      "args": ["C:\\path\\to\\word-mcp-server\\build\\parent.js"]
    }
  }
}
```

> If `npm install` fails due to missing C++ compiler, install Visual Studio 2022 Build Tools with the "Desktop development with C++" workload.

## Tools (110+ total)

### Document Lifecycle
| Tool | Description |
|------|-------------|
| `word_get_status` | Query current Word/document state (NO_WORD/NO_DOC/DOC_ACTIVE/DIALOG) |
| `word_document` | Universal entry: open by path or use active doc |
| `word_create` | Create new document (visible in Word) |
| `word_create_from_template` | Create new document from .dotx template |
| `word_open` | Open existing .docx file |
| `word_save` | Save current document (auto `.bak` backup) |
| `word_save_as` | Save with new path (auto-detect format) |
| `word_close` | Close current document |
| `word_get_info` | Get word/paragraph/page count etc. |
| `word_get_structure` | Get heading outline with paragraph indices |
| `word_quit` | Quit Microsoft Word entirely |

### Content Input
| Tool | Description |
|------|-------------|
| `word_type_text` | Type text at cursor (smooth/instant mode) |
| `word_insert_paragraph` | Insert paragraph breaks |
| `word_insert_page_break` | Insert page break |
| `word_insert_horizontal_line` | Insert horizontal line |
| `word_insert_list` | Insert bullet or numbered list |
| `word_set_columns` | Set section column count (newsletter layout) |
| `word_insert_section_break` | Insert section break (nextPage/continuous/evenPage/oddPage) |
| `word_insert_file` | Insert content of another .docx at cursor |
| `word_backspace` | Delete characters before cursor |
| `word_write_markdown` | Write formatted content from Markdown (headings, bold, italic, strikethrough, code, links, lists, tables, blockquotes, hr) |
| `word_get_text` | Read full document text |
| `word_get_paragraph` | Read a specific paragraph's text |
| `word_get_table_data` | Extract table content as structured data (rows×columns grid) |
| `word_get_comments` | List all comments with author info |
| `word_get_bookmarks` | List all bookmarks |
| `word_get_lists` | List all bullet/numbered lists with hierarchy |
| `word_get_sections` | List sections with page setup info (orientation, columns, page size) |
| `word_export_to_pdf` | Export document to PDF |

### Links, References & Comments
| Tool | Description |
|------|-------------|
| `word_add_hyperlink` | Add hyperlink (supports screenTip & subAddress) |
| `word_add_footnote` | Add footnote (auto-numbered) |
| `word_add_comment` | Add comment |

### Find, Selection & Navigation
| Tool | Description |
|------|-------------|
| `word_find_text` | Find text and position cursor at the match |
| `word_find_replace` | Find and replace text |
| `word_go_to` | Navigate to page/section/bookmark/end |
| `word_go_to_paragraph` | Navigate to specific paragraph by 1-based index |
| `word_select_all` | Select all content |
| `word_select_text` | Select text by character range |
| `word_select_current_word` | Select word at cursor |
| `word_select_current_paragraph` | Select paragraph at cursor |
| `word_delete` | Delete selected content |

### Clipboard & Undo
| Tool | Description |
|------|-------------|
| `word_copy` | Copy to clipboard |
| `word_cut` | Cut to clipboard |
| `word_paste` | Paste from clipboard |
| `word_undo` | Undo last action(s) |
| `word_redo` | Redo undone action(s) |
| `word_get_cursor_info` | Get cursor position & selection info |

### Formatting
| Tool | Description |
|------|-------------|
| `word_set_font` | Set font name/size/bold/italic/color/strikethrough/superscript/subscript |
| `word_set_paragraph` | Set alignment/indent/spacing |
| `word_apply_style` | Apply named style (Heading 1, Title, etc.) |
| `word_set_page_setup` | Set margins/orientation/paper size |
| `word_set_properties` | Set document metadata (title, author, etc.) |
| `word_list_styles` | List available styles |

### Tables
| Tool | Description |
|------|-------------|
| `word_insert_table` | Insert table with optional data |
| `word_edit_cell` | Edit table cell text |
| `word_edit_cells` | Batch-fill multiple cells with 2D data array |
| `word_add_table_row` | Add row (optional data) |
| `word_delete_table_row` | Delete row |
| `word_add_table_column` | Add column (left of specified column or append) |
| `word_delete_table_column` | Delete column |
| `word_set_table_borders` | Set table border style/color/width |
| `word_set_table_shading` | Set table or header row background color |
| `word_merge_table_cells` | Merge cell range |
| `word_set_column_width` | Set column width in points |
| `word_set_row_height` | Set row height in points |
| `word_set_cell_font` | Set font for specific cell |
| `word_set_cell_vertical_alignment` | Set cell vertical alignment (top/center/bottom) |
| `word_apply_table_style` | Apply built-in Word table style by name |
| `word_table_to_text` | Convert table to plain text (remove table structure) |
| `word_text_to_table` | Convert selected delimited text to table |

### Media
| Tool | Description |
|------|-------------|
| `word_insert_image` | Insert image from file |
| `word_insert_chart` | Insert chart (column/bar/line/pie/area, child process sets data) |
| `word_insert_textbox` | Insert textbox (position/size/orientation) |

### Document Structure
| Tool | Description |
|------|-------------|
| `word_set_header` | Set page header |
| `word_set_footer` | Set page footer |
| `word_set_page_numbers` | Add page numbers |
| `word_insert_toc` | Insert table of contents |
| `word_add_bookmark` | Add bookmark |
| `word_set_watermark` | Set or remove watermark (e.g. "DRAFT") |

### Batch Mode
| Tool | Description |
|------|-------------|
| `word_batch_start` | Start batch mode: disables screen updating, accelerates operations 3-10x |
| `word_batch_end` | End batch mode: restores screen updating and refreshes window |

### Manager Layer (High-level Document Builder)
| Tool | Description |
|------|-------------|
| `word_mgr_create_document` | Create doc with auto close + page setup |
| `word_mgr_write_content` | Write Markdown with auto cursor & ScreenUpdating management |
| `word_mgr_apply_heading` | Apply heading style at end of document |
| `word_mgr_set_header` / `word_mgr_set_footer` | Set header/footer with auto context return |
| `word_mgr_set_page_numbers` | Add page numbers with auto context return |
| `word_mgr_insert_table` / `word_mgr_insert_chart` / `word_mgr_insert_image` | Insert elements with auto paragraph separator |
| `word_mgr_insert_list` / `word_mgr_insert_textbox` | Insert list/textbox with auto cursor reset |
| `word_mgr_add_bookmark` / `word_mgr_add_comment` / `word_mgr_add_footnote` / `word_mgr_add_hyperlink` | Add markers with auto deep cursor reset |
| `word_mgr_insert_section_break` | Section break with auto context return |
| `word_mgr_format_page` | Configure page layout in one call |
| `word_mgr_set_watermark` | Set watermark with auto context return |
| `word_mgr_save` | Save with optional PDF export |

### Semantic Navigation
| Tool | Description |
|------|-------------|
| `word_locate` | Resolve heading/paragraph/table/bookmark position (read-only) |
| `word_select_at` | Move cursor to semantic location (heading/paragraph/table/bookmark) |
| `word_insert_at` | Write text at semantic location |
| `word_write_markdown_at` | Write Markdown at semantic location |
| `word_edit_cell_at` | Edit table cell at semantic location |

### Variable Replacement
| Tool | Description |
|------|-------------|
| `word_replace_variables` | Replace {{placeholder}} variables in document with provided values |

## Usage Examples

### Write formatted content with Markdown

Use `word_write_markdown` to insert styled content in a single call:

``````md
# Section Title
Some **bold** and *italic* text with `inline code`, ~~strikethrough~~, and [a link](https://example.com).

- Top level item
  - Nested item with **bold**
- Another top item

1. First step
2. Second step

```
code block with pre-formatted text
```

| Name | Value |
|------|-------|
| Alpha | 100 |
| Beta | 200 |

> Notable quote here
---
``````

## Prompts

- `create_report` — Step-by-step for structured reports
- `format_document` — Guided formatting workflow

## Architecture

```
src/
├── index.ts                          # Entry point (McpServer + stdio transport)
├── parent.ts                         # Parent watchdog process (30s timeout, auto-restart)
├── child.ts                          # MCP server process (spawned by parent watchdog)
├── manager/                          # Document builder pattern (high-level API)
│   ├── types.ts                      # 11 composite operation parameter interfaces
│   └── word-manager.ts               # Document builder (auto cursor reset, context switching)
├── word/                             # Word COM adapter layer (via winax native addon)
│   ├── session.ts                    # Winax COM session manager (lazy start / lifecycle / auto-recover)
│   ├── types.ts                      # Shared interfaces
│   ├── application.ts                # Document lifecycle (create/open/save/close/quit)
│   ├── document.ts                   # Document info, path scanning
│   ├── document-registry.ts          # path→doc mapping, dedup, stale detection
│   ├── process-monitor.ts            # WINWORD.EXE process health monitor
│   ├── word-base.ts                  # Abstract base for Word COM operations
│   ├── word-text-editor.ts           # Text editing operations (type, insert, backspace, find, replace)
│   ├── word-document-structure.ts    # Document structure operations (header/footer/toc/bookmark/watermark)
│   ├── word-table-editor.ts          # Table editing operations (cell, row, column, merge)
│   ├── word-media-editor.ts          # Media operations (image, chart, textbox)
│   ├── formatting.ts                 # Formatting ops (font/paragraph/page/styles)
│   ├── position-map.ts               # Semantic positioning system (binary search + dirty markers)
│   ├── variable-replacer.ts          # {{placeholder}} replacement engine
│   ├── word-markdown.ts              # Markdown → Word converter (headings, bold, italic, tables, lists, etc.)
│   └── chart-data-worker.ts          # Child process: isolated chart data setting (15s timeout)
├── security/                         # Security & policy layer (4-layer defense)
│   ├── errors.ts                     # Typed error hierarchy + path redaction
│   ├── path-sanitizer.ts             # Path traversal protection (7 attack vectors)
│   ├── policy.ts                     # Security policy + SecurityManager
│   ├── rate-limiter.ts               # Sliding window rate limiting
│   └── audit.ts                      # Operation audit logging
└── server/                           # MCP layer
    ├── tools/                        # 110+ tool registrations (12 modules)
    │   ├── helper.ts                 # mcpCall wrapper: timeout, audit, rate-limit, error handling
    │   ├── document.ts               # 11 lifecycle tools
    │   ├── content.ts                # 26 content/selection/clipboard tools
    │   ├── formatting.ts             # 6 formatting tools
    │   ├── tables.ts                 # 17 table tools
    │   ├── media.ts                  # 3 media tools
    │   ├── structure.ts              # 7 structure tools
    │   ├── reader.ts                 # 8 reading tools
    │   ├── batch.ts                  # 2 batch mode tools
    │   ├── markdown.ts               # 1 markdown tool (helper)
    │   ├── semantic.ts               # 5 semantic navigation tools (locate, select_at, insert_at, etc.)
    │   ├── variable.ts               # 1 variable replacement tool
    │   └── manager.ts                # 11 high-level Manager tools (word_mgr_*)
    └── prompts/                      # Prompt registrations
```

### Communication Flow

```
MCP Client (Claude Desktop, etc.)
    │  JSON-RPC over stdio
    ▼
build/parent.js  (Watchdog — 30s timeout, auto-restart on crash)
    │  spawns + pipes stdin/stdout
    ▼
build/child.js   (McpServer)
    │  winax COM Automation (native)
    ├─ Main thread: 110+ tools (doc/content/format/table/media/structure/reader/batch/markdown/semantic/variable/manager)
    └─ Child process: chart data setting (fork + 15s timeout)
            │
            ▼
        WINWORD.EXE  (visible window)
```

> No PowerShell bridge required. Direct COM calls via winax for lower latency and higher reliability.

## Security

- Path traversal detection blocks `..` and `~` patterns
- Cross-drive traversal detection on Windows
- Windows system directories (`C:\Windows\`) blocked
- ADS paths (alternate data streams) blocked
- Optional allowed-directories whitelist
- Text length capped at 1M characters
- Network paths disabled by default
- Macros forcibly disabled via `AutomationSecurity = 3`
- Auto-backup creates `.bak` before overwriting

## Safety

- Word opens in visible window — you see all operations
- `Ctrl+Z` undo works for all operations
- Server disconnect does NOT close Word (unsaved work preserved)
- Automatic `.bak` backup before every save
- Chart data set via isolated child process (15s timeout, non-blocking)
- Rate limiting available

## Tests

10 test files (9 unit + 1 integration) — **116 tests, all passing** ✅

```bash
npm test
```

Unit tests mock the winax layer and don't require a real Word installation. Coverage includes:

| Category | Test files | Coverage |
|----------|-----------|----------|
| Security | `errors`, `path-sanitizer`, `policy`, `rate-limiter`, `audit` | 5 files, 30 tests |
| COM Session | `session` | 13 tests (mock winax) |
| Markdown Parser | `markdown` | 33 tests (pure functions, no mock) |
| Text Editor | `text-editor` | 16 tests (splitIntoBatches + COM mock) |
| Position Map | `position-map` | 21 tests (semantic resolution logic) |
| Integration | `smoke` | 3 tests (MCP protocol + tool listing) |

## Related Skills (Optional)

Word MCP runs independently. To achieve consistent styling (like fixed font, spacing, or heading conventions), you can create rule files in your project (e.g., `.opencode/skills/` or your Agent's supported format) to guide AI behavior. Whether to use them is up to you — the MCP functions fully without them.

## License

MIT
