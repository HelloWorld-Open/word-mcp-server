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
1. Before editing: word_create (new doc) or word_document (open existing) or word_open (open file)
2. Before formatting: word_select_* or word_find_text to target content
3. Before table ops: word_insert_table to create the table
4. Before save: finish all edits
5. Before export: word_save first, then word_export_to_pdf

Key workflows:
- New doc: word_create → word_set_page_setup → word_type_text/word_write_markdown → word_save
- Edit existing: word_document → word_find_text/word_select_at → word_type_text → word_save
- Table: word_insert_table → word_edit_cells → word_set_table_borders → word_set_cell_font
- Review: word_set_track_changes → word_type_text → word_accept_changes`
            : detail === "states"
            ? `# Word State Machine

The Word MCP server operates in one of 4 states:

State 1: NO_WORD — Word.exe is not running
  →  word_get_status: State: NO_WORD
  →  Any tool automatically starts Word (via ensureAlive)
  →  Next: word_document() or word_create()

State 2: NO_DOC — Word is running, no document active
  →  word_get_status: State: NO_DOC
  →  Editing/formatting tools return [NO_DOCUMENT]
  →  Next: word_create() or word_document()

State 3: DOC_ACTIVE — Document open and editable
  →  word_get_status: State: DOC_ACTIVE
  →  All editing/formatting/table tools available
  →  Next: type, format, save, export

State 4: DIALOG — Word showing a modal dialog
  →  Tools will hang until dialog is closed
  →  word_get_status still works (detects dialog)
  →  Next: close the dialog manually in Word window
  →  If stuck: use word_quit() to restart`
            : `# Word State Machine & Ordering Rules

## 4-State Model
State 1: NO_WORD — Word.exe not running (auto-start on any call)
State 2: NO_DOC — Word running, no active document
State 3: DOC_ACTIVE — Document open and editable
State 4: DIALOG — Modal dialog open (tools hang)

## Operation Ordering
Every tool description includes "Order:" with the recommended sequence.

### New Document
word_create → word_set_page_setup → word_type_text/word_write_markdown → word_save → word_export_to_pdf

### Edit Existing  
word_document/word_open → word_find_text/word_select_at → word_type_text → word_save

### Table Work
word_insert_table → word_edit_cells → word_set_table_borders → word_apply_table_style → word_set_cell_font

### Track Changes Review
word_set_track_changes → word_type_text → word_accept_changes/word_reject_changes → word_save

### Template Variables
word_document(template) → word_replace_variables → word_save

### Batch Operations
word_type_text → word_batch_start → multi-edit → word_batch_end`,
        },
      }],
    })
  )
}
