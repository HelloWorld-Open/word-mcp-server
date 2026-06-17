import { z } from "zod"
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js"

export function registerStateMachinePrompt(server: McpServer): void {
  server.registerPrompt(
    "state_machine",
    {
      description: "Learn about the Word document state machine and operation ordering rules",
      argsSchema: {
        detail: z.enum(["full", "states", "ordering"]).optional().default("full"),
      },
    },
    ({ detail }) => ({
      messages: [{
        role: "user",
        content: {
          type: "text",
          text: detail === "ordering"
            ? `# Ordering Rules
Every Word tool description includes an Order: field showing the recommended call sequence.

General principles:
1. Before editing: word_stream_start (new doc) or word_document (open existing) or word_open (open file)
2. Before formatting: word_select_* or word_find_text to target content
3. Before table ops: word_insert_table to create the table
4. Before save: finish all edits
5. Before export: word_stream_end to save and optionally export PDF

Key workflows:
- New doc: word_stream_start → word_stream_block → word_set_page_region({target:'header', text:'Title', quiet:true}) → word_stream_end
- Edit existing: word_document → word_find_text/word_select_at → word_insert_at → word_save
- Table: word_insert_table → word_edit_cells → word_set_table_borders → word_set_cell_font
- Review: word_set_track_changes → word_insert_at → word_track_changes_apply({action:'accept'})

Streaming session active — 4 tools are blocked: word_document, word_open, word_close, word_quit.
Call word_stream_end to finish streaming, then those tools become available again.

Edit mode active — word_stream_start is blocked.
Call word_close before creating a new document.`
            : detail === "states"
            ? `# Word State Machine

The Word MCP server operates in one of 4 base states:

State 1: NO_WORD — Word.exe is not running
  →  word_get_status: State: NO_WORD
  →  Any tool automatically starts Word (via ensureAlive)
  →  Next: word_document() or word_stream_start()

State 2: NO_DOC — Word is running, no document active
  →  word_get_status: State: NO_DOC
  →  Editing/formatting tools return [NO_DOCUMENT]
  →  Next: word_stream_start() or word_document()

State 3: DOC_ACTIVE — Document open and editable
  →  word_get_status: State: DOC_ACTIVE
  →  All editing/formatting/table tools available
  →  Next: type, format, save, export

State 4: DIALOG — Word showing a modal dialog
  →  Tools will hang until dialog is closed
  →  word_get_status still works (detects dialog)
  →  Next: close the dialog manually in Word window
  →  If stuck: use word_quit() to restart

In addition to the 4 base states, two **sub-states** apply:

**STREAMING** (during word_stream_start..word_stream_end):
  →  doc suffix shows "[流式会话活跃]"
  →  word_document, word_open, word_close, word_quit are blocked
  →  Call word_stream_end to return to DOC_ACTIVE

**EDITING** (during word_document/word_open..word_close):
  →  doc suffix shows "[编辑模式]"
  →  word_stream_start is blocked
  →  Call word_close before creating a new document`
            : `# Word State Machine & Ordering Rules

## 4-State Model
State 1: NO_WORD — Word.exe not running (auto-start on any call)
State 2: NO_DOC — Word running, no active document
State 3: DOC_ACTIVE — Document open and editable
State 4: DIALOG — Modal dialog open (tools hang)

## 2 Sub-States

During word_stream_start..word_stream_end, a **STREAMING** sub-state is active:
- doc suffix shows "[流式会话活跃]"
- word_document, word_open, word_close, word_quit are blocked
- Call word_stream_end to return to DOC_ACTIVE

During word_document/word_open..word_close, an **EDITING** sub-state is active:
- doc suffix shows "[编辑模式]"
- word_stream_start is blocked
- Call word_close before creating a new document

## Operation Ordering
Every tool description includes "Order:" with the recommended sequence.

### New Document
word_stream_start → word_stream_block/word_insert_table → word_set_page_region({target:'header', text:'...', quiet:true}) → word_stream_end({save:true, exportPath})

### Edit Existing  
word_document/word_open → word_find_text/word_select_at → word_insert_at → word_save

### Table Work
word_insert_table → word_edit_cells → word_set_table_borders → word_apply_table_style → word_set_cell_font

### Track Changes Review
word_set_track_changes → word_insert_at → word_track_changes_apply({action:'accept'})/word_track_changes_apply({action:'reject'}) → word_save

### Template Variables
word_stream_start({templatePath}) → word_replace_variables → word_stream_end

### Batch Operations
Write all content via word_stream_block calls, one per chapter/section.`,
        },
      }],
    })
  )
}
