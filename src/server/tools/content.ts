import { z } from "zod"
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js"
import { WordTextEditor } from "../../word/word-text-editor.js"
import { SecurityManager } from "../../security/policy.js"
import { mcpCall } from "./helper.js"

export function registerContentTools(
  server: McpServer,
  content: WordTextEditor,
  security: SecurityManager,
): void {
  server.registerTool(
    "word_type_text",
    {
      description: "Type text at the current cursor position. WHEN: entering text after positioning cursor with word_go_to_paragraph or word_find_text. NOT: want formatted headings/lists/tables? use word_write_markdown.",
      inputSchema: {
        text: z.string().min(0).max(1000000).describe("Text to insert"),
        mode: z.enum(["smooth", "instant"]).optional().describe("'smooth' (default) splits into sentence chunks; 'instant' writes all at once"),
      },
    },
    mcpCall(security, "word_type_text", async ({ text, mode }) => {
      security.validateTextLength(text)
      await content.typeText(text, mode)
      return `Action: Text inserted (${text.length} chars)\nDetail: Mode: ${mode ?? "smooth"}\nNext: word_set_font({size:12, name:"Arial"}) or word_select_all()`
    }),
  )

  server.registerTool(
    "word_insert_paragraph",
    {
      description: "Insert paragraph breaks.",
      inputSchema: {
        count: z.number().int().min(1).max(100).optional().describe("Number of paragraph breaks (default: 1)"),
      },
    },
    mcpCall(security, "word_insert_paragraph", async ({ count }) => {
      await content.insertParagraph(count)
      const c = count ?? 1
      return `Action: Paragraph(s) inserted (${c})\nNext: word_type_text({text:"..."}) or word_set_paragraph({lineSpacingRule:"double"})`
    }),
  )

  server.registerTool(
    "word_insert_page_break",
    {
      description: "Insert a page break at the cursor position.",
    },
    mcpCall(security, "word_insert_page_break", async () => {
      await content.insertPageBreak()
      return "Action: Page break inserted\nNext: word_type_text({text:\"...\", mode:\"instant\"}) to continue on next page"
    }),
  )

  server.registerTool(
    "word_insert_horizontal_line",
    {
      description: "Insert a horizontal line at the cursor position.",
    },
    mcpCall(security, "word_insert_horizontal_line", async () => {
      await content.insertHorizontalLine()
      return "Action: Horizontal line inserted\nNext: word_type_text({text:\"...\"}) to continue below the line"
    }),
  )

  server.registerTool(
    "word_insert_list",
    {
      description: "Insert a bulleted or numbered list at the cursor position.",
      inputSchema: {
        type: z.enum(["bullet", "number"]).describe("List type: 'bullet' or 'number'"),
        items: z.array(z.string().max(100000)).min(1).max(500).describe("List items"),
      },
    },
    mcpCall(security, "word_insert_list", async ({ type, items }) => {
      await content.insertList(type, items)
      return `Action: ${type === "bullet" ? "Bullet" : "Numbered"} list inserted (${items.length} items)\nNext: word_type_text({text:"..."}) to continue after list`
    }),
  )

  server.registerTool(
    "word_add_hyperlink",
    {
      description: "Add a hyperlink at the current cursor position.",
      inputSchema: {
        text: z.string().min(1).max(1000).describe("Display text"),
        address: z.string().min(1).max(2083).describe("URL or file path"),
        subAddress: z.string().max(255).optional().describe("Anchor or bookmark within the document"),
        screenTip: z.string().max(500).optional().describe("Tooltip on hover"),
      },
    },
    mcpCall(security, "word_add_hyperlink", async ({ text, address, subAddress, screenTip }) => {
      await content.addHyperlink(text, address, subAddress, screenTip)
      return `Action: Hyperlink added "${text}"\nDetail: ${address}\nNext: word_type_text({text:"..."}) to continue after the link`
    }),
  )

  server.registerTool(
    "word_add_footnote",
    {
      description: "Add a footnote at the current cursor position.",
      inputSchema: {
        text: z.string().min(1).max(10000).describe("Footnote text"),
      },
    },
    mcpCall(security, "word_add_footnote", async ({ text }) => {
      await content.addFootnote(text)
      const preview = text.slice(0, 80) + (text.length > 80 ? "..." : "")
      return `Action: Footnote added\nDetail: "${preview}"\nNext: word_type_text({text:"..."}) to continue in body text`
    }),
  )

  server.registerTool(
    "word_insert_section_break",
    {
      description: "Insert a section break at the cursor position.",
      inputSchema: {
        type: z.enum(["nextPage", "continuous", "evenPage", "oddPage"]).optional().describe("Section break type (default: nextPage)"),
      },
    },
    mcpCall(security, "word_insert_section_break", async ({ type }) => {
      await content.insertSectionBreak(type)
      const t = type ?? "nextPage"
      const labels: Record<string, string> = {
        nextPage: "Starts new section on next page",
        continuous: "Continues on same page",
        evenPage: "Starts on next even-numbered page",
        oddPage: "Starts on next odd-numbered page",
      }
      return `Action: Section break inserted (${t})\nDetail: ${labels[t]}\nNext: word_set_header({text:"Section2 Header"}) or word_set_page_numbers({target:"footer"})`
    }),
  )

  server.registerTool(
    "word_set_columns",
    {
      description: "Set the number of text columns for the current section.",
      inputSchema: {
        count: z.number().int().min(1).max(4).describe("Number of columns (1-4)"),
        spacing: z.number().min(0).max(20).optional().describe("Space between columns in cm"),
      },
    },
    mcpCall(security, "word_set_columns", async ({ count, spacing }) => {
      await content.setColumns(count, spacing)
      const spaceText = spacing != null ? `${spacing}cm` : "default"
      return `Action: Columns set (${count})\nDetail: Spacing: ${spaceText}\nNext: word_type_text({text:"...", mode:"instant"}) or word_insert_section_break({type:"continuous"})`
    }),
  )

  server.registerTool(
    "word_find_text",
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
    mcpCall(security, "word_find_text", async ({ findText, matchCase, matchWholeWord, direction, wrap }) => {
      const result = await content.findText(findText, { matchCase, matchWholeWord, direction, wrap })
      if (!result) return "Action: Text not found"
      return `Action: Text found\nDetail: ${result}`
    }),
  )

  server.registerTool(
    "word_go_to_paragraph",
    {
      description: "Navigate to a specific paragraph by 1-based index.",
      inputSchema: {
        index: z.number().int().min(1).max(1000000).describe("Paragraph index (1-based)"),
      },
    },
    mcpCall(security, "word_go_to_paragraph", async ({ index }) => {
      await content.goToParagraph(index)
      return `Action: Navigated to paragraph ${index}\nNext: word_get_cursor_info() or word_type_text({text:"..."})`
    }),
  )

  server.registerTool(
    "word_find_replace",
    {
      description: "Find and replace text in the document.",
      inputSchema: {
        findText: z.string().min(1).max(5000).describe("Text to find"),
        replaceWith: z.string().max(5000).describe("Replacement text"),
        matchCase: z.boolean().optional().describe("Case sensitive search"),
        matchWholeWord: z.boolean().optional().describe("Match whole words only"),
        replaceAll: z.boolean().optional().describe("Replace all occurrences (default: true)"),
      },
    },
    mcpCall(security, "word_find_replace", async ({ findText, replaceWith, matchCase, matchWholeWord, replaceAll }) => {
      await content.findReplace(findText, replaceWith, { matchCase, matchWholeWord, replaceAll })
      const mode = replaceAll !== false ? "all" : "first"
      return `Action: Find & replace completed (${mode})\nDetail: "${findText}" → "${replaceWith}"\nNext: word_undo({count:1}) or word_save()`
    }),
  )

  server.registerTool(
    "word_go_to",
    {
      description: "Navigate to a specific location. WHEN: need to jump to a page/section/line/bookmark/end. NOT: know the paragraph index? use word_go_to_paragraph.",
      inputSchema: {
        what: z.enum(["page", "section", "line", "bookmark", "end"]).optional().describe("Target (default: page)"),
        which: z.enum(["first", "last", "next", "previous"]).optional().describe("Which occurrence (default: first)"),
      },
    },
    mcpCall(security, "word_go_to", async ({ what, which }) => {
      await content.goTo(what, which)
      const w = what ?? "page"
      return `Action: Navigated to ${w} (${which ?? "first"})\nNext: word_get_cursor_info() or word_type_text({text:"..."})`
    }),
  )

  server.registerTool(
    "word_select_all",
    {
      description: "Select all content in the document. WHEN: need to format/copy/delete the entire document. NOT: want a specific range? use word_select_text.",
    },
    mcpCall(security, "word_select_all", async () => {
      await content.selectAll()
      return `Action: All content selected\nNext: word_set_font({size:12, name:"Arial"}) or word_copy() or word_delete() or word_type_text({text:"...", mode:"instant"})`
    }),
  )

  server.registerTool(
    "word_select_text",
    {
      description: "Select a specific range of text by character position.",
      inputSchema: {
        start: z.number().int().min(0).describe("Starting character position (0-based)"),
        length: z.number().int().min(1).max(1000000).describe("Number of characters to select"),
      },
    },
    mcpCall(security, "word_select_text", async ({ start, length }) => {
      await content.selectText(start, length)
      return `Action: Text selected (${start}–${start + length})\nDetail: ${length} characters\nNext: word_set_font({bold:true}) or word_copy() or word_delete()`
    }),
  )

  server.registerTool(
    "word_select_current_word",
    {
      description: "Select the word at the current cursor position.",
    },
    mcpCall(security, "word_select_current_word", async () => {
      await content.selectCurrentWord()
      return "Action: Current word selected\nNext: word_set_font({bold:true, italic:true}) or word_copy() or word_type_text({text:\"...\"})"
    }),
  )

  server.registerTool(
    "word_select_current_paragraph",
    {
      description: "Select the paragraph at the cursor position.",
    },
    mcpCall(security, "word_select_current_paragraph", async () => {
      await content.selectCurrentParagraph()
      return `Action: Current paragraph selected\nNext: word_set_paragraph({alignment:"center"}) or word_copy() or word_type_text({text:"..."})`
    }),
  )

  server.registerTool(
    "word_delete",
    {
      description: "Delete the currently selected text.",
    },
    mcpCall(security, "word_delete", async () => {
      await content.deleteSelection()
      return `Action: Selection deleted\nNext: word_undo({count:1}) or word_type_text({text:"...", mode:"instant"})`
    }),
  )

  server.registerTool(
    "word_backspace",
    {
      description: "Delete characters before the cursor.",
      inputSchema: {
        count: z.number().int().min(1).max(1000).optional().describe("Number of backspaces (default: 1)"),
      },
    },
    mcpCall(security, "word_backspace", async ({ count }) => {
      await content.backspace(count)
      return `Action: Backspace x${count ?? 1}\nNext: word_type_text({text:"..."}) or word_undo({count:1})`
    }),
  )

  server.registerTool(
    "word_copy",
    {
      description: "Copy selected content to clipboard. WHEN: after selecting text with word_select_* or word_find_text. NOT: want to remove from document? use word_cut.",
    },
    mcpCall(security, "word_copy", async () => {
      await content.copy()
      return `Action: Copied to clipboard\nNext: word_paste() or word_type_text({text:"...", mode:"instant"})`
    }),
  )

  server.registerTool(
    "word_cut",
    {
      description: "Cut selected content to clipboard.",
    },
    mcpCall(security, "word_cut", async () => {
      await content.cut()
      return "Action: Cut to clipboard\nNext: word_paste() or word_undo({count:1})"
    }),
  )

  server.registerTool(
    "word_paste",
    {
      description: "Paste content from clipboard at cursor position. WHEN: after copying/cutting content with word_copy or word_cut. NOT: want to type new content? use word_type_text.",
    },
    mcpCall(security, "word_paste", async () => {
      await content.paste()
      return `Action: Pasted from clipboard\nNext: word_undo({count:1}) or word_set_font({size:12, name:"Arial"})`
    }),
  )

  server.registerTool(
    "word_undo",
    {
      description: "Undo the last action(s). WHEN: need to revert a mistake or unwanted change. NOT: undo history is cleared after word_save.",
      inputSchema: {
        count: z.number().int().min(1).max(100).optional().describe("Number of undo steps (default: 1)"),
      },
    },
    mcpCall(security, "word_undo", async ({ count }) => {
      await content.undo(count)
      return `Action: Undo x${count ?? 1}\nNext: word_redo({count:1}) to restore if undone too far`
    }),
  )

  server.registerTool(
    "word_redo",
    {
      description: "Redo the last undone action(s).",
      inputSchema: {
        count: z.number().int().min(1).max(100).optional().describe("Number of redo steps (default: 1)"),
      },
    },
    mcpCall(security, "word_redo", async ({ count }) => {
      await content.redo(count)
      return `Action: Redo x${count ?? 1}\nNext: word_type_text({text:"..."}) or word_save()`
    }),
  )

  server.registerTool(
    "word_get_cursor_info",
    {
      description: "Get current cursor position info. WHEN: need to know cursor location and selection range before editing. NOT: want full document stats? use word_get_info.",
    },
    mcpCall(security, "word_get_cursor_info", async () => {
      const info = await content.getCursorInfo()
      const parts = [`Action: Cursor info`, `Selection: ${info.hasSelection ? "YES" : "NO"}`, `Range: ${info.start}–${info.end}`]
      if (info.hasSelection) parts.push(`Selected: "${info.selectedText.slice(0, 100)}"`)
      parts.push("Next: word_select_text({start:0, length:10}) to refine or word_type_text({text:\"...\"}) or word_delete()")
      return parts.join("\n")
    }),
  )

  server.registerTool(
    "word_insert_file",
    {
      description: "Insert the content of another Word document at the current cursor position. WHEN: need to merge content from another document. NOT: want to open another file? use word_document.",
      inputSchema: {
        path: z.string().min(1).max(4096).describe("Full path to the .docx file to insert"),
      },
    },
    mcpCall(security, "word_insert_file", async ({ path }) => {
      const safePath = security.pathSanitizer.validateForRead(path)
      await content.insertFile(safePath)
      return `Action: File inserted\nDetail: ${safePath}\nNext: word_type_text() or word_save()`
    }),
  )
}
