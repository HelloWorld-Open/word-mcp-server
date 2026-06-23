import { z } from "zod"
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js"
import { WordCursor } from "../../word/word-cursor.js"
import { SecurityManager } from "../../security/policy.js"
import type { ServerContext } from "../server-context.js"
import { createRegTool } from "./shared.js"

export function registerCursorTools(
  server: McpServer,
  context: ServerContext,
  cursor: WordCursor,
  security: SecurityManager,
): void {
  const regTool = createRegTool(server, security, context)

  regTool("word_find_text",
    {
      description: "WHEN: need to locate a keyword or phrase before editing. WHAT: searches forward/backward from cursor and positions cursor at the first match. CONSTRAINT: cursor moves to match; does NOT modify text. Use word_find_replace if replacement is needed.",
      inputSchema: {
        findText: z.string().min(1).max(5000).describe("Text to search for (case-insensitive by default)"),
        matchCase: z.boolean().optional().describe("Case sensitive search (default: false)"),
        matchWholeWord: z.boolean().optional().describe("Match whole words only (default: false)"),
        direction: z.enum(["forward", "backward"]).optional().describe("Search direction (default: forward)"),
        wrap: z.boolean().optional().describe("Wrap around to beginning/end if not found (default: true)"),
      },
    },
    async ({ findText, matchCase, matchWholeWord, direction, wrap }) => {
      const result = await cursor.findText(findText, { matchCase, matchWholeWord, direction, wrap })
      if (!result) return "Action: Text not found"
      return `Action: Text found\nDetail: ${result}`
    },
  )

  regTool("word_find_replace",
    {
      description: "WHEN: need to replace specific words or phrases across the entire document. WHAT: searches for text and replaces with new text, optionally case-sensitive or whole-word. CONSTRAINT: modifies document content; use word_find_text first to preview matches. Supports replace-all or single replace.",
      inputSchema: {
        findText: z.string().min(1).max(5000).describe("Text to find"),
        replaceWith: z.string().max(5000).describe("Replacement text"),
        matchCase: z.boolean().optional().describe("Case sensitive search (default: false)"),
        matchWholeWord: z.boolean().optional().describe("Match whole words only (default: false)"),
        replaceAll: z.boolean().optional().describe("Replace all occurrences (default: true). Set to false to replace only the first occurrence."),
      },
    },
    async ({ findText, replaceWith, matchCase, matchWholeWord, replaceAll }) => {
      await cursor.findReplace(findText, replaceWith, { matchCase, matchWholeWord, replaceAll })
      const mode = replaceAll !== false ? "all" : "first"
      return `Action: Find & replace completed (${mode})\nDetail: "${findText}" → "${replaceWith}"\nNext: word_undo_redo({action:"undo"}) or word_save()`
    },
  )

  regTool("word_go_to_paragraph",
    {
      description: "WHEN: you know the exact paragraph index from word_get_structure. WHAT: moves cursor to the start of a specific paragraph by 1-based index. CONSTRAINT: requires valid paragraph index; use word_get_structure() first to obtain indices. For heading-based navigation, use word_select_at({by:'heading', match:'...'}).",
      inputSchema: {
        index: z.number().int().min(1).max(1000000).describe("Paragraph index (1-based, obtained from word_get_structure output like 'H1 ¶3')"),
      },
    },
    async ({ index }) => {
      await cursor.goToParagraph(index)
      return `Action: Navigated to paragraph ${index}\nNext: word_get_cursor_info() or word_type_text({text:"..."})`
    },
  )

  regTool("word_go_to",
    {
      description: "WHEN: need to jump to a structural location (page/section/line/bookmark/document-end). WHAT: navigates to the first/last/next/previous occurrence of the target type. CONSTRAINT: 'bookmark' requires an existing bookmark created via word_add_bookmark. For exact paragraph navigation, use word_go_to_paragraph.",
      inputSchema: {
        what: z.enum(["page", "section", "line", "bookmark", "end"]).optional().describe("Target (default: page)"),
        which: z.enum(["first", "last", "next", "previous"]).optional().describe("Which occurrence (default: first)"),
      },
    },
    async ({ what, which }) => {
      await cursor.goTo(what, which)
      const w = what ?? "page"
      return `Action: Navigated to ${w} (${which ?? "first"})\nNext: word_get_cursor_info() or word_type_text({text:"..."})`
    },
  )

  regTool("word_select_all",
    {
      description: "WHEN: need to format, copy, or delete the entire document at once. WHAT: selects all content in the active document (Ctrl+A). CONSTRAINT: selection covers the full document range; subsequent actions (delete, set_font, copy) apply to every paragraph.",
    },
    async () => {
      await cursor.selectAll()
      return `Action: All content selected\nNext: word_set_font({size:12, name:"Arial"}) or word_clipboard({action:"copy"}) or word_delete() or word_type_text({text:"...", mode:"instant"})`
    },
  )

  regTool("word_select_text",
    {
      description: "WHEN: need to select a precise character range for targeted formatting or copying. WHAT: selects text from character position N with length L (0-based). CONSTRAINT: requires knowing start/length from word_get_cursor_info. For whole-document selection use word_select_all; for current word/paragraph use word_select_current.",
      inputSchema: {
        start: z.number().int().min(0).describe("Starting character position (0-based). Use word_get_cursor_info to find positions."),
        length: z.number().int().min(1).max(1000000).describe("Number of characters to select"),
      },
    },
    async ({ start, length }) => {
      await cursor.selectText(start, length)
      return `Action: Text selected (${start}–${start + length})\nDetail: ${length} characters\nNext: word_set_font({bold:true}) or word_clipboard({action:"copy"}) or word_delete()`
    },
  )

  regTool("word_select_current",
    {
      description: "WHEN: need to quickly format or copy the word or paragraph under the cursor without specifying character positions. WHAT: selects either the current word (cursor within a word) or the current paragraph (cursor within any paragraph). CONSTRAINT: scope=word selects one word; scope=paragraph selects the entire containing paragraph.",
      inputSchema: {
        scope: z.enum(["word", "paragraph"]).describe("Selection scope: 'word' for the single word under cursor, 'paragraph' for the entire current paragraph"),
      },
    },
    async ({ scope }) => {
      if (scope === "word") {
        await cursor.selectCurrentWord()
        return "Action: Current word selected\nNext: word_set_font({bold:true, italic:true}) or word_clipboard({action:\"copy\"})"
      }
      await cursor.selectCurrentParagraph()
      return "Action: Current paragraph selected\nNext: word_set_paragraph({alignment:\"center\"}) or word_clipboard({action:\"copy\"})"
    },
  )

  regTool("word_delete",
    {
      description: "WHEN: want to remove selected content from the document. WHAT: deletes the current selection (like pressing Delete key). CONSTRAINT: requires prior selection via word_select_text, word_select_all, or word_select_current. Does NOT work on empty selection; use word_backspace instead.",
    },
    async () => {
      await cursor.deleteSelection()
      return `Action: Selection deleted\nNext: word_undo_redo({action:"undo"}) or word_type_text({text:"...", mode:"instant"})`
    },
  )

  regTool("word_backspace",
    {
      description: "WHEN: need to delete the last N characters before the cursor (like pressing Backspace N times). WHAT: removes characters one at a time before the cursor position. CONSTRAINT: max 1000 characters at once. For large deletions, select text first then use word_delete.",
      inputSchema: {
        count: z.number().int().min(1).max(1000).optional().describe("Number of backspaces (default: 1). Each backspace deletes one character before the cursor."),
      },
    },
    async ({ count }) => {
      await cursor.backspace(count)
      return `Action: Backspace x${count ?? 1}\nNext: word_type_text({text:"..."}) or word_undo_redo({action:"undo"})`
    },
  )

  regTool("word_clipboard",
    {
      description: "WHEN: need to move or duplicate content within or between documents (cut/copy/paste). WHAT: action=copy copies selection to clipboard; action=cut removes and copies; action=paste inserts clipboard at cursor. CONSTRAINT: copy/cut require prior selection; paste requires clipboard content from a prior copy/cut.",
      inputSchema: {
        action: z.enum(["copy", "cut", "paste"]).describe("'copy' to copy selection to clipboard, 'cut' to remove and copy, 'paste' to insert clipboard content at cursor"),
      },
    },
    async ({ action }) => {
      if (action === "copy") {
        await cursor.copy()
        return "Action: Copied to clipboard\nNext: word_clipboard({action:\"paste\"}) or word_type_text({text:\"...\", mode:\"instant\"})"
      }
      if (action === "cut") {
        await cursor.cut()
        return "Action: Cut to clipboard\nNext: word_clipboard({action:\"paste\"}) or word_undo_redo({action:\"undo\"})"
      }
      await cursor.paste()
      return `Action: Pasted from clipboard\nNext: word_undo_redo({action:"undo"}) or word_set_font({size:12, name:"Arial"})`
    },
  )

  regTool("word_undo_redo",
    {
      description: "WHEN: need to revert a mistake or restore after undoing too far. WHAT: action=undo reverses the last N actions; action=redo re-applies previously undone N actions. CONSTRAINT: undo history is cleared after word_save. Cannot undo past a save boundary. Max 100 steps at once.",
      inputSchema: {
        action: z.enum(["undo", "redo"]).describe("'undo' to revert recent actions, 'redo' to restore previously undone actions"),
        count: z.number().int().min(1).max(100).optional().describe("Number of steps to undo/redo (default: 1). Each step reverses one prior action."),
      },
    },
    async ({ action, count }) => {
      if (action === "undo") {
        await cursor.undo(count)
        return `Action: Undo x${count ?? 1}\nNext: word_undo_redo({action:"redo"}) to restore if undone too far`
      }
      await cursor.redo(count)
      return `Action: Redo x${count ?? 1}\nNext: word_type_text({text:"..."}) or word_save()`
    },
  )

  regTool("word_get_cursor_info",
    {
      description: "WHEN: need to know the current cursor position and selection range before editing. WHAT: returns cursor character offset, selection status, and selected text preview. CONSTRAINT: read-only; does NOT modify cursor position or document. Use before word_select_text or word_insert_at to verify position.",
    },
    async () => {
      const info = await cursor.getCursorInfo()
      const parts = [`Action: Cursor info`, `Selection: ${info.hasSelection ? "YES" : "NO"}`, `Range: ${info.start}–${info.end}`]
      if (info.hasSelection) parts.push(`Selected: "${info.selectedText.slice(0, 100)}"`)
      parts.push("Next: word_select_text({start:0, length:10}) to refine or word_type_text({text:\"...\"}) or word_delete()")
      return parts.join("\n")
    },
  )
}
