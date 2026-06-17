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
      description: "Find text and position cursor at the match. WHEN: need to locate a keyword before editing. NOT: want to also replace? use word_find_replace.",
      inputSchema: {
        findText: z.string().min(1).max(5000).describe("Text to search for"),
        matchCase: z.boolean().optional().describe("Case sensitive search"),
        matchWholeWord: z.boolean().optional().describe("Match whole words only"),
        direction: z.enum(["forward", "backward"]).optional().describe("Search direction (default: forward)"),
        wrap: z.boolean().optional().describe("Wrap around to beginning/end (default: true)"),
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
      description: "Find and replace text in the document. WHEN: need to replace specific words or phrases across the document. NOT: just want to locate text without replacing? use word_find_text.",
      inputSchema: {
        findText: z.string().min(1).max(5000).describe("Text to find"),
        replaceWith: z.string().max(5000).describe("Replacement text"),
        matchCase: z.boolean().optional().describe("Case sensitive search"),
        matchWholeWord: z.boolean().optional().describe("Match whole words only"),
        replaceAll: z.boolean().optional().describe("Replace all occurrences (default: true)"),
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
      description: "Navigate to a specific paragraph by 1-based index. WHEN: you know the paragraph index from word_get_structure. NOT: want to navigate by heading name? use word_select_at({by:'heading', match:'...'}).",
      inputSchema: {
        index: z.number().int().min(1).max(1000000).describe("Paragraph index (1-based)"),
      },
    },
    async ({ index }) => {
      await cursor.goToParagraph(index)
      return `Action: Navigated to paragraph ${index}\nNext: word_get_cursor_info() or word_type_text({text:"..."})`
    },
  )

  regTool("word_go_to",
    {
      description: "Navigate to a specific location. WHEN: need to jump to a page/section/line/bookmark/end. NOT: know the paragraph index? use word_go_to_paragraph.",
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
      description: "Select all content in the document. WHEN: need to format/copy/delete the entire document. NOT: want a specific range? use word_select_text.",
    },
    async () => {
      await cursor.selectAll()
      return `Action: All content selected\nNext: word_set_font({size:12, name:"Arial"}) or word_clipboard({action:"copy"}) or word_delete() or word_type_text({text:"...", mode:"instant"})`
    },
  )

  regTool("word_select_text",
    {
      description: "Select a specific range of text by character position. WHEN: need to select a precise range for formatting or copying. NOT: want to select all content? use word_select_all. NOT: want to select the word/paragraph under cursor? use word_select_current.",
      inputSchema: {
        start: z.number().int().min(0).describe("Starting character position (0-based)"),
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
      description: "Select the word or paragraph at the current cursor position. WHEN: need to quickly format or copy the word/paragraph under cursor. NOT: want to select a specific character range? use word_select_text.",
      inputSchema: {
        scope: z.enum(["word", "paragraph"]).describe("Selection scope: 'word' or 'paragraph'"),
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
      description: "Delete the currently selected text. WHEN: want to remove selected content from the document. NOT: nothing selected? use word_select_text or word_select_all first to select content to delete.",
    },
    async () => {
      await cursor.deleteSelection()
      return `Action: Selection deleted\nNext: word_undo_redo({action:"undo"}) or word_type_text({text:"...", mode:"instant"})`
    },
  )

  regTool("word_backspace",
    {
      description: "Delete characters before the cursor. WHEN: need to remove a few characters one at a time like pressing Backspace. NOT: want to delete selected text? select first then use word_delete.",
      inputSchema: {
        count: z.number().int().min(1).max(1000).optional().describe("Number of backspaces (default: 1)"),
      },
    },
    async ({ count }) => {
      await cursor.backspace(count)
      return `Action: Backspace x${count ?? 1}\nNext: word_type_text({text:"..."}) or word_undo_redo({action:"undo"})`
    },
  )

  regTool("word_clipboard",
    {
      description: "Copy or cut selected content to/from clipboard, or paste clipboard content at cursor. WHEN: need to move or duplicate content within or between documents. NOT: want to type new text? use word_stream_block or word_insert_at.",
      inputSchema: {
        action: z.enum(["copy", "cut", "paste"]).describe("'copy' to copy selection, 'cut' to remove and copy, 'paste' to insert clipboard"),
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
      description: "Undo or redo the last action(s). WHEN: need to revert a mistake or restore after undo. NOT: undo history is cleared after word_save.",
      inputSchema: {
        action: z.enum(["undo", "redo"]).describe("'undo' to revert, 'redo' to restore"),
        count: z.number().int().min(1).max(100).optional().describe("Number of steps (default: 1)"),
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
      description: "Get current cursor position info. WHEN: need to know cursor location and selection range before editing. NOT: want full document stats? use word_get_info.",
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
