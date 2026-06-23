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
      description: "WHEN: need to add spacing or move to a new line. WHAT: inserts N paragraph breaks (blank lines) at the cursor position. CONSTRAINT: max 100 breaks. For rich content (headings, lists, tables) use word_stream_block or word_insert_at instead.",
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
      description: "WHEN: need to force a page break or add a horizontal line separator. WHAT: type=page inserts a page break (content continues on next page); type=line inserts a horizontal rule line. CONSTRAINT: page break changes pagination. For section breaks (different margins/orientation) use word_insert_section_break.",
      inputSchema: {
        type: z.enum(["page", "line"]).describe("'page' for a page break (new page), 'line' for a horizontal rule (thematic break)"),
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
      description: "WHEN: need to create a structured bulleted or numbered list. WHAT: inserts a multi-item list at the cursor position. CONSTRAINT: supports bullet or numbered types only (no checkboxes). For rich markdown content (lists with headings/tables), use word_stream_block instead.",
      inputSchema: {
        type: z.enum(["bullet", "number"]).describe("List type: 'bullet' for unordered (•), 'number' for ordered (1. 2. 3.)"),
        items: z.array(z.string().max(100000)).min(1).max(500).describe("List items as plain text strings. Each item becomes one list entry."),
      },
    },
    async ({ type, items }) => {
      await content.insertList(type, items)
      return `Action: ${type === "bullet" ? "Bullet" : "Numbered"} list inserted (${items.length} items)\nNext: word_type_text({text:"..."}) to continue after list`
    },
  )

  regTool("word_add_hyperlink",
    {
      description: "WHEN: need to insert a clickable link to a URL or document location. WHAT: inserts hyperlink at cursor with display text and target address. CONSTRAINT: address must be a valid URL or file path; subAddress links to a bookmark within a document (requires existing bookmark via word_add_bookmark).",
      inputSchema: {
        text: z.string().min(1).max(1000).describe("Display text (visible clickable text in the document)"),
        address: z.string().min(1).max(2083).describe("URL (e.g., 'https://example.com') or file path (e.g., 'C:\\docs\\file.docx')"),
        subAddress: z.string().max(255).optional().describe("Anchor or bookmark name within the target document to link to a specific location"),
        screenTip: z.string().max(500).optional().describe("Tooltip text shown on hover (default: the address)"),
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
      description: "WHEN: need to add an explanatory note at the bottom of the page. WHAT: inserts a footnote at the cursor position and places the cursor inside the footnote area to type content. CONSTRAINT: footnotes appear at page bottom in print layout; for review notes visible in the Review pane, use word_add_comment instead.",
      inputSchema: {
        text: z.string().min(1).max(10000).describe("Footnote text (appears at the bottom of the page)"),
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
      description: "WHEN: need to change page layout (margins, orientation, columns) within the same document. WHAT: inserts a section divider that allows different formatting per section. CONSTRAINT: type=nextPage starts new section on next page; type=continuous starts on same page. After insertion, page setup settings apply to the new section only.",
      inputSchema: {
        type: z.enum(["nextPage", "continuous", "evenPage", "oddPage"]).optional().describe("Section break type (default: nextPage). nextPage=new section on next page, continuous=same page, evenPage/oddPage=start on even/odd numbered page"),
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
      description: "WHEN: need newspaper-style multi-column text layout (like a newsletter). WHAT: sets the current section to N columns with optional spacing. CONSTRAINT: applies to current section only; use word_insert_section_break first to create separate sections with different column counts.",
      inputSchema: {
        count: z.number().int().min(1).max(4).describe("Number of columns (1-4). 1=single column (normal), 2=newspaper style, 3-4=narrower columns"),
        spacing: z.number().min(0).max(20).optional().describe("Space between columns in cm (default: document default, typically 1.27cm)"),
      },
    },
    async ({ count, spacing }) => {
      await content.setColumns(count, spacing)
      const spaceText = spacing != null ? `${spacing}cm` : "default"
      return `Action: Columns set (${count})\nDetail: Spacing: ${spaceText}\nNext: word_type_text({text:"...", mode:"instant"}) or word_insert_section_break({type:"continuous"})`
    },
  )


}
