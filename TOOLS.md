> [English](./TOOLS.md) | [中文](./TOOLS.zh-CN.md)

# Word MCP Server — Tool Reference

> 110+ tools across 12 modules.

---

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

---

Back to [README](./README.md)
