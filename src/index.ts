import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js"
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js"
import { WordSession } from "./word/session.js"
import { WordApplicationManager } from "./word/application.js"
import { WordDocument } from "./word/document.js"
import { WordTextEditor } from "./word/word-text-editor.js"
import { WordTableEditor } from "./word/word-table-editor.js"
import { WordMediaEditor } from "./word/word-media-editor.js"
import { WordDocumentStructure } from "./word/word-document-structure.js"
import { WordFormatting } from "./word/formatting.js"
import { VariableReplacer } from "./word/variable-replacer.js"
import { WordMarkdown } from "./word/word-markdown.js"
import { PositionMap } from "./word/position-map.js"
import { SecurityManager } from "./security/policy.js"

import { WordDocumentManager } from "./manager/word-manager.js"
import { registerDocumentTools } from "./server/tools/document.js"
import { registerContentTools } from "./server/tools/content.js"
import { registerFormattingTools } from "./server/tools/formatting.js"
import { registerTableTools } from "./server/tools/tables.js"
import { registerMediaTools } from "./server/tools/media.js"
import { registerStructureTools } from "./server/tools/structure.js"
import { registerMarkdownTool } from "./server/tools/markdown.js"
import { registerReaderTools } from "./server/tools/reader.js"
import { registerBatchTools } from "./server/tools/batch.js"
import { registerVariableTool } from "./server/tools/variable.js"
import { registerSemanticTools } from "./server/tools/semantic.js"
import { registerManagerTools } from "./server/tools/manager.js"
import { registerReportPrompts } from "./server/prompts/report-prompts.js"
import { registerStateMachinePrompt } from "./server/prompts/state-machine.js"
import { setStatusSession, setPositionMap } from "./server/tools/helper.js"

function createLogger(prefix: string) {
  return (...args: unknown[]) => {
    console.error(`[${prefix}]`, ...args)
  }
}

async function main() {
  const log = createLogger("word-mcp")

  log("Starting Word MCP Server...")

  const session = new WordSession()
  session.setOnLog((level, message) => {
    log(`[session/${level}] ${message}`)
  })
  setStatusSession(session)

  const appManager = new WordApplicationManager(session)
  const docOps = new WordDocument(session)
  const textEditor = new WordTextEditor(session)
  const tableEditor = new WordTableEditor(session)
  const mediaEditor = new WordMediaEditor(session)
  const documentStructure = new WordDocumentStructure(session)
  const formatting = new WordFormatting(session)
  const markdown = new WordMarkdown(session)
  const positionMap = new PositionMap(session)
  const variableReplacer = new VariableReplacer(session)
  const security = new SecurityManager()
  const documentManager = new WordDocumentManager(
    session, appManager, docOps, textEditor, tableEditor,
    mediaEditor, documentStructure, formatting, markdown,
  )

  setPositionMap(positionMap)

  const cleanup = async () => {
    log("Shutting down...")
    try { await appManager.quit() } catch { /* ignore */ }
    process.exit(0)
  }
  process.on("SIGINT", cleanup)
  process.on("SIGTERM", cleanup)
  process.on("unhandledRejection", (reason) => {
    log("Unhandled rejection:", reason)
  })
  process.on("uncaughtException", (err) => {
    log("Uncaught exception:", err)
    process.exit(1)
  })

  const server = new McpServer(
    {
      name: "word-mcp-server",
      version: "1.0.0",
    },
    {
      instructions: [
        "This server provides tools to interact with Microsoft Word documents in real time.",
        "Word will be visible on screen — changes happen in real time.",
        "",
        "",
        "--- GUIDANCE ---",
        "Prefer word_mgr_* tools (word_mgr_create_document, word_mgr_write_content, etc.) when building documents from scratch — they auto-manage cursor position and context switching.",
        "Fall back to raw word_* tools for: editing existing content, reading/inspecting, fine-grained table formatting, find/replace, clipboard, undo/redo, track changes, batch mode.",
        "",
        "--- STATE MACHINE ---",
        "The system can be in one of these states. ALWAYS call word_get_status to check before deciding:",
        "",
        "State 1: NO_WORD — Word.exe is not running.",
        "  → All document/content/table tools return ENGINE_NOT_RUNNING.",
        "  → Action: Restart MCP server.",
        "",
        "State 2: NO_DOC — Word is running but no document is active.",
        "  → Most content/table/formatting tools return NO_DOCUMENT.",
        "  → Actions: word_document(path) to open, word_document() for untitled, word_create for new.",
        "",
        "State 3: DOC_ACTIVE — A document is open and ready.",
        "  → All tools available. Use word_get_info / word_get_structure / word_get_cursor_info to orient.",
        "",
        "State 4: DIALOG — Word is showing a modal dialog (e.g. Save As, confirmation).",
        "  → Most tools will hang until the dialog is dismissed. Check Word window visually.",
        "  → Action: Dismiss the dialog manually in the Word window, then retry.",
        "",
        "--- FIRST PRINCIPLES: Action -> Effect ---",
        "word_get_status       Query current state. Always call first. Side effect: NONE.",
        "word_document         Switch to or open a file. Errors if file not found. Omit path for untitled.",
        "word_create           Create new blank doc. Only when NO_DOC state.",
        "word_create_from_template  Create doc from .dotx template (inherits styles/headers/watermarks).",
        "word_open             Open existing file from disk. Use word_document for smart open.",
        "word_close            Close active doc. Returns to NO_DOC state.",
        "word_save             Save current doc (with .bak backup).",
        "word_type_text        Insert text at cursor. Use mode='instant' for bulk writing.",
        "word_set_font         Format selected text or set forward-typing font.",
        "word_set_paragraph    Format current paragraph (alignment, indents, spacing).",
        "word_apply_style      Apply named style (Heading 1/2/3, Normal, Title, etc.).",
        "word_select_text      Select by character range. Use word_get_cursor_info to find ranges.",
        "word_find_text        Search + position cursor at match. Returns context.",
        "word_go_to_paragraph  Navigate by 1-based paragraph index (from word_get_structure).",
        "word_insert_table     Create table at cursor. Returns index=1 for subsequent operations.",
        "word_edit_cell        Fill/replace a table cell's content.",
        "word_undo/redo        Revert/restore recent actions.",
        "word_set_header/footer  Add headers/footers to current section.",
        "word_insert_image     Embed image at cursor. File must exist on disk.",
        "word_insert_chart     Insert chart with data (column/bar/line/pie/area).",
        "word_insert_file      Insert content of another .docx at cursor (merge documents).",
        "word_write_markdown   Write formatted content from Markdown (# headings, **bold**, *italic*, ~~strikethrough~~, `code`, ```code blocks```, [links](url), lists (nested), tables, > blockquotes, --- hr). One call replaces 5-15 manual steps.",
        "word_get_text         Read full document text. Essential for verifying written content.",
        "word_get_paragraph    Read a specific paragraph's text by index.",
        "word_get_table_data   Extract table content as structured data (rows×columns grid).",
        "word_get_comments     List all comments in the document with author info.",
        "word_get_bookmarks    List all bookmarks in the document.",
        "word_get_lists        List all bullet/numbered lists with hierarchy.",
        "word_get_sections     List sections with page setup info (orientation, columns, page size).",
        "word_export_to_pdf    Export document to PDF (creates file on disk, doc unchanged).",
        "word_batch_start      Start batch mode: freezes Word screen, accelerates operations 3-10x.",
        "word_batch_end        End batch mode: restores screen updating and refreshes window.",
        "",
        "--- ANTI-PATTERNS: Common Mistakes ---",
        "These incorrect sequences produce NO EFFECT or WRONG RESULTS:",
        "",
        "BLOCK 1: Set font before typing",
        "  WRONG: word_set_font({size:14}) → word_type_text('hello')",
        "  RIGHT: word_type_text('hello') → word_select_all() → word_set_font({size:14})",
        "",
        "BLOCK 2: Find + goTo paragraph (find already moves cursor)",
        "  WRONG: word_find_text('keyword') then word_go_to_paragraph(n)",
        "  RIGHT: word_find_text('keyword') — cursor already at match, type/edit directly",
        "",
        "BLOCK 3: Format without selecting first",
        "  WRONG: word_set_font({bold:true}) — only affects NEW text after cursor",
        "  RIGHT: word_select_all() then word_set_font({bold:true})",
        "",
        "BLOCK 4: word_open vs word_document",
        "  WRONG: word_open({path}) — opens new window every time, even if already open",
        "  RIGHT: word_document({path}) — switches to open doc, or opens if not",
        "",
        "BLOCK 5: Undo after save",
        "  WRONG: word_save() → word_undo() — undo history is cleared by save",
        "  RIGHT: word_undo() during editing, word_save() when done",
        "",
        "BLOCK 6: word_create with active document",
        "  WRONG: word_create({title}) — errors if a doc is already active",
        "  RIGHT: word_get_status() → word_close() first → word_create({title})",
        "",
        "BLOCK 7: Write strategy — batch new content, incremental for editing",
        "  Rule: Creating new content? Write ALL in ONE call. Editing? Do per-section.",
        "",
        "  BATCH (creating document from scratch):",
        "    RIGHT: Generate ENTIRE content, then call word_type_text ONCE (mode='instant')",
        "    RIGHT: For formatted docs, use word_write_markdown ONCE with full Markdown",
        "    WHY: COM writes 1000 chars in ~10ms. AI generates in seconds. One-shot is 3-4 total calls.",
        "",
        "  INCREMENTAL (editing existing document):",
        "    RIGHT: Navigate to target paragraph → word_type_text for the section being edited",
        "    RIGHT: Edit multiple scattered sections in sequence, one tool call per section",
        "    WHY: Editing requires precise positioning. Each section edit needs its own navigatie + write.",
        "",
        "--- WORKFLOW TEMPLATES ---",
        "",
        "# Edit specific sections of a document (incremental — one call per edit)",
        "  1. word_get_structure() → get paragraph indices",
        "  2. word_go_to_paragraph({index:5}) → navigate to target section",
        "  3. word_type_text({text: 'replacement', mode: 'instant'}) → write only that section",
        "  4. Repeat 2-3 for each section that needs editing",
        "",
        "# Edit an existing document",
        "  1. word_get_status({path:'C:\\docs\\file.docx'}) → confirm file exists",
        "  2. word_document({path:'C:\\docs\\file.docx'}) → opens it",
        "  3. word_get_structure() → get heading outline with paragraph indices",
        "  4. word_go_to_paragraph({index:3}) → navigate to target",
        "  5. word_type_text({text:'new content', mode:'instant'}) or word_edit_cell({row:1, column:1, text:'data'})",
        "  6. word_save() → persist changes",
        "",
        "# Read / navigate a document",
        "  1. word_get_status({path:'C:\\docs\\file.docx'}) → check state",
        "  2. word_document({path:'C:\\docs\\file.docx'}) → open",
        "  3. word_get_info() → stats (pages, words, paragraphs, table count)",
        "  4. word_get_structure() → heading outline with paragraph indices",
        "  5. word_get_text() → read full text content",
        "  6. word_get_paragraph({index:5}) → read a specific paragraph",
        "  7. word_get_table_data({index:1}) → read table content as grid",
        "  8. word_get_comments() → review existing comments",
        "  9. word_get_bookmarks() → list navigation anchors",
        " 10. word_get_lists() → review list structure",
        " 11. word_get_sections() → check page layout (orientation, columns, page size)",
        " 12. word_find_text({findText:'keyword'}) → search and position cursor",
        " 13. word_export_to_pdf() → create PDF preview",
        "",
        "# Format a table",
        "  1. word_insert_table({rows:5, columns:3, data:[['H1','H2','H3'],['a','b','c']]})",
        "  2. word_apply_table_style({styleName:'Light List Accent 1'})",
        "  3. word_set_column_width({column:1, width:120}) → resize columns",
        "  4. word_set_cell_font({row:1, column:1, bold:true}) → format header",
        "  5. word_set_table_shading({color:'#E8F0FE', target:'row'}) → color header row",
        "",
        "# Accelerated batch editing (3-10x faster)",
        "  1. word_batch_start() → freeze Word screen, start batch",
        "  2. word_type_text('...') → runs fast (no repaint)",
        "  3. word_set_font(...) → runs fast",
        "  4. word_insert_table(...) → runs fast",
        "  5. word_batch_end() → restore screen, Word refreshes with all changes",
        "",
        "# Merge documents",
        "  1. word_document({path:'C:\\main.docx'}) → open the base document",
        "  2. word_go_to({what:'end'}) → position cursor at end",
        "  3. word_insert_file({path:'C:\\appendix.docx'}) → insert content of another doc",
        "  4. word_insert_page_break() → add page break before next section",
        "  5. word_insert_file({path:'C:\\disclaimer.docx'}) → insert another document",
        "  6. word_save_as({path:'C:\\merged-output.docx'}) → save the merged result",
        "",
        "--- ERROR RECOVERY ---",
        "All errors include a >> Recovery: hint with specific next steps.",
        "Common patterns:",
        "- NO_DOCUMENT → Use word_document or word_create first.",
        "- TABLE_NOT_FOUND → Check table count with word_get_info.",
        "- FILE_NOT_FOUND → Use word_get_status to find files, or word_create.",
        "- ENGINE_TIMEOUT → Reduce operation size (fewer rows/columns, shorter text). Check Word dialogs.",
        "",
        "Path Safety:",
        "- File paths are validated against directory traversal attacks.",
        "- Only files in allowed directories can be accessed.",
        "- System directories (Windows) are blocked.",
      ].join("\n"),
    }
  )

  registerDocumentTools(server, appManager, docOps, security)
  registerContentTools(server, textEditor, security)
  registerFormattingTools(server, formatting, security)
  registerTableTools(server, tableEditor, security)
  registerMediaTools(server, mediaEditor, security)
  registerStructureTools(server, documentStructure, security)
  registerReportPrompts(server)
  registerStateMachinePrompt(server)
  registerMarkdownTool(server, markdown, textEditor, positionMap, security)
  registerVariableTool(server, variableReplacer, security)
  registerReaderTools(server, docOps, security)
  registerBatchTools(server, session, security)
  registerSemanticTools(server, textEditor, tableEditor, positionMap, security)
  registerManagerTools(server, documentManager, security)

  const transport = new StdioServerTransport()
  await server.connect(transport)

  log("Word MCP Server connected and ready")
}

main().catch((error) => {
  console.error("[word-mcp] Fatal error:", error)
  process.exit(1)
})
