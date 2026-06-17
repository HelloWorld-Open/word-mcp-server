import { z } from "zod"
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js"
import { WordContentWriter } from "../../word/word-content-writer.js"
import { SecurityManager } from "../../security/policy.js"
import type { ServerContext } from "../server-context.js"
import { createRegTool } from "./shared.js"

export function registerContentTools(
  server: McpServer,
  context: ServerContext,
  content: WordContentWriter,
  security: SecurityManager,
): void {
  const regTool = createRegTool(server, security, context)
  regTool("word_insert_paragraph",
    {
      description: "Insert one or more paragraph breaks at the cursor position. WHEN: need to add spacing or move to a new line. NOT: want to insert content instead of blank lines? use word_stream_block or word_insert_at.",
      inputSchema: {
        count: z.number().int().min(1).max(100).optional().describe("Number of paragraph breaks (default: 1)"),
      },
    },
    async ({ count }) => {
      await content.insertParagraph(count)
      const c = count ?? 1
      return `Action: Paragraph(s) inserted (${c})\nNext: word_type_text({text:"..."}) or word_set_paragraph({lineSpacingRule:"double"})`
    },
  )

  regTool("word_insert_break",
    {
      description: "Insert a page break or horizontal line at the cursor position. WHEN: need to force a page break or add a visual separator. NOT: want to insert a section break? use word_insert_section_break.",
      inputSchema: {
        type: z.enum(["page", "line"]).describe("'page' for page break, 'line' for horizontal rule"),
      },
    },
    async ({ type }) => {
      if (type === "page") {
        await content.insertPageBreak()
        return "Action: Page break inserted\nNext: word_type_text({text:\"...\", mode:\"instant\"}) to continue on next page"
      }
      await content.insertHorizontalLine()
      return "Action: Horizontal line inserted\nNext: word_type_text({text:\"...\"}) to continue below the line"
    },
  )

  regTool("word_insert_list",
    {
      description: "Insert a bulleted or numbered list at the cursor position. WHEN: need to create a structured list. NOT: want rich markdown content including lists with other elements? use word_stream_block instead.",
      inputSchema: {
        type: z.enum(["bullet", "number"]).describe("List type: 'bullet' or 'number'"),
        items: z.array(z.string().max(100000)).min(1).max(500).describe("List items"),
      },
    },
    async ({ type, items }) => {
      await content.insertList(type, items)
      return `Action: ${type === "bullet" ? "Bullet" : "Numbered"} list inserted (${items.length} items)\nNext: word_type_text({text:"..."}) to continue after list`
    },
  )

  regTool("word_add_hyperlink",
    {
      description: "Add a hyperlink at the current cursor position. WHEN: need to insert a clickable link to a URL or document location. NOT: want to insert a bookmark anchor for internal navigation? use word_add_bookmark.",
      inputSchema: {
        text: z.string().min(1).max(1000).describe("Display text"),
        address: z.string().min(1).max(2083).describe("URL or file path"),
        subAddress: z.string().max(255).optional().describe("Anchor or bookmark within the document"),
        screenTip: z.string().max(500).optional().describe("Tooltip on hover"),
        quiet: z.boolean().optional().describe("简洁输出模式"),
      },
    },
    async ({ text, address, subAddress, screenTip, quiet }) => {
      await content.addHyperlink(text, address, subAddress, screenTip)
      if (quiet) return "Hyperlink added"
      return `Action: Hyperlink added "${text}"\nDetail: ${address}\nNext: word_type_text({text:"..."}) to continue after the link`
    },
  )

  regTool("word_add_footnote",
    {
      description: "Add a footnote at the current cursor position. WHEN: need to add an explanatory note at the bottom of the page. NOT: want to add a comment (not printed in layout)? use word_add_comment.",
      inputSchema: {
        text: z.string().min(1).max(10000).describe("Footnote text"),
        quiet: z.boolean().optional().describe("简洁输出模式"),
      },
    },
    async ({ text, quiet }) => {
      await content.addFootnote(text)
      if (quiet) return "Footnote added"
      const preview = text.slice(0, 80) + (text.length > 80 ? "..." : "")
      return `Action: Footnote added\nDetail: "${preview}"\nNext: word_type_text({text:"..."}) to continue in body text`
    },
  )

  regTool("word_insert_section_break",
    {
      description: "Insert a section break at the cursor position. WHEN: need to change page layout (margins, orientation, columns) within the same document. NOT: just want a page break? use word_insert_break with type:'page'.",
      inputSchema: {
        type: z.enum(["nextPage", "continuous", "evenPage", "oddPage"]).optional().describe("Section break type (default: nextPage)"),
        quiet: z.boolean().optional().describe("简洁输出模式"),
      },
    },
    async ({ type, quiet }) => {
      await content.insertSectionBreak(type)
      if (quiet) return "Section break inserted"
      const t = type ?? "nextPage"
      const labels: Record<string, string> = {
        nextPage: "Starts new section on next page",
        continuous: "Continues on same page",
        evenPage: "Starts on next even-numbered page",
        oddPage: "Starts on next odd-numbered page",
      }
      return `Action: Section break inserted (${t})\nDetail: ${labels[t]}\nNext: word_set_page_region({target:"header", text:"Section2 Header"}) or word_set_page_numbers({target:"footer"})`
    },
  )

  regTool("word_set_columns",
    {
      description: "Set the number of text columns for the current section. WHEN: need newspaper-style multi-column layout. NOT: want to change page margins or orientation? use word_set_page_setup.",
      inputSchema: {
        count: z.number().int().min(1).max(4).describe("Number of columns (1-4)"),
        spacing: z.number().min(0).max(20).optional().describe("Space between columns in cm"),
      },
    },
    async ({ count, spacing }) => {
      await content.setColumns(count, spacing)
      const spaceText = spacing != null ? `${spacing}cm` : "default"
      return `Action: Columns set (${count})\nDetail: Spacing: ${spaceText}\nNext: word_type_text({text:"...", mode:"instant"}) or word_insert_section_break({type:"continuous"})`
    },
  )


}
