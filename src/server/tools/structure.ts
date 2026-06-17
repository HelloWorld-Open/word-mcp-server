import { z } from "zod"
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js"
import { WordFormatter } from "../../word/word-formatter.js"
import { SecurityManager } from "../../security/policy.js"
import type { ServerContext } from "../server-context.js"
import { createRegTool, ColorSchema } from "./shared.js"

export function registerStructureTools(
  server: McpServer,
  context: ServerContext,
  formatter: WordFormatter,
  security: SecurityManager,
): void {
  const regTool = createRegTool(server, security, context)
  regTool("word_set_page_region",
    {
      description: "Set the header or footer text. WHEN: every page needs repeating top/bottom text. NOT: want page numbers only? use word_set_page_numbers.",
      inputSchema: {
        target: z.enum(["header", "footer"]).describe("'header' for top of page, 'footer' for bottom of page"),
        text: z.string().max(5000).describe("Header/footer text"),
        alignment: z.enum(["left", "center", "right"]).optional().describe("Text alignment"),
        sectionIndex: z.number().int().min(1).optional().describe("Section index (1-based, default: last section)"),
        quiet: z.boolean().optional().describe("简洁输出模式"),
      },
    },
    async ({ target, text, alignment, sectionIndex, quiet }) => {
      if (target === "header") {
        await formatter.setHeader(text, alignment, sectionIndex)
        if (quiet) return "Header set"
        let next = `word_set_page_region({target:"footer", text:"Page "}) or word_set_page_numbers({target:"footer"})`
        if (sectionIndex) next += `\nSection: ${sectionIndex}`
        return `Action: Header set\nDetail: Alignment: ${alignment ?? "left"}\nNext: ${next}`
      }
      await formatter.setFooter(text, alignment, sectionIndex)
      if (quiet) return "Footer set"
      let next = `word_set_page_region({target:"header", text:"..."}) or word_set_page_numbers({target:"footer"})`
      if (sectionIndex) next += `\nSection: ${sectionIndex}`
      return `Action: Footer set\nDetail: Alignment: ${alignment ?? "left"}\nNext: ${next}`
    },
  )

  regTool("word_set_page_numbers",
    {
      description: "Add page numbers to header or footer. WHEN: every page needs page numbering. NOT: want to set custom header/footer text? use word_set_page_region.",
      inputSchema: {
        target: z.enum(["header", "footer"]).describe("Where to place page numbers"),
        alignment: z.enum(["left", "center", "right"]).optional().describe("Page number alignment (default: center)"),
        sectionIndex: z.number().int().min(1).optional().describe("Section index (1-based, default: last section)"),
        quiet: z.boolean().optional().describe("简洁输出模式"),
      },
    },
    async ({ target, alignment, sectionIndex, quiet }) => {
      await formatter.setPageNumbers(target, alignment, sectionIndex)
      if (quiet) return "Page numbers added"
      let detail = `Action: Page numbers added to ${target} (${alignment ?? "center"})`
      if (sectionIndex) detail += `\nSection: ${sectionIndex}`
      return `${detail}\nNext: word_set_page_region({target:"header", text:"... - Page "}) for custom text alongside numbers`
    },
  )

  regTool("word_insert_toc",
    {
      description: "Insert a Table of Contents at the cursor position. WHEN: document has headings and needs an auto-generated table of contents. NOT: want to list items manually? create a list with word_insert_list instead.",
    },
    async () => {
      await formatter.insertToc()
      return "Action: Table of Contents inserted\nNext: Right-click the TOC > Update Field to refresh after adding headings"
    },
  )

  regTool("word_add_bookmark",
    {
      description: "Add a bookmark at the current cursor position. WHEN: need to create a navigation anchor for hyperlinks or cross-references. NOT: want to jump to an existing bookmark? use word_go_to({what:'bookmark'}).",
      inputSchema: {
        name: z.string().min(1).max(100).regex(/^[a-zA-Z_\u4e00-\u9fff][a-zA-Z0-9_\u4e00-\u9fff]*$/).describe("Bookmark name (alphanumeric + underscore, start with letter/Chinese)"),
        quiet: z.boolean().optional().describe("简洁输出模式"),
      },
    },
    async ({ name, quiet }) => {
      await formatter.addBookmark(name)
      if (quiet) return "Bookmark added"
      return `Action: Bookmark added "${name}"\nNext: word_add_hyperlink with subAddress="${name}" to link here`
    },
  )

  regTool("word_add_comment",
    {
      description: "Add a comment at the current selection. WHEN: need to add review notes or feedback visible in the Review pane. NOT: want to add a footnote (printed at page bottom)? use word_add_footnote.",
      inputSchema: {
        text: z.string().min(1).max(10000).describe("Comment text"),
        quiet: z.boolean().optional().describe("简洁输出模式"),
      },
    },
    async ({ text, quiet }) => {
      await formatter.addComment(text)
      if (quiet) return "Comment added"
      const preview = text.slice(0, 80) + (text.length > 80 ? "..." : "")
      return `Action: Comment added\nDetail: "${preview}"\nNext: word_type_text({text:"..."}) to continue editing`
    },
  )

  regTool("word_set_watermark",
    {
      description: "Add or remove a text watermark (e.g. 'DRAFT', 'CONFIDENTIAL'). WHEN: need to indicate document status (draft, confidential, sample). NOT: want to add text that repeats on every page? use word_set_page_region to set a header instead.",
      inputSchema: {
        text: z.string().min(1).max(100).describe("Watermark text"),
        remove: z.boolean().optional().describe("Set true to remove existing watermarks"),
        fontSize: z.number().int().min(12).max(200).optional().describe("Font size (default: 48)"),
        color: ColorSchema.optional().describe("Watermark color"),
        quiet: z.boolean().optional().describe("简洁输出模式"),
      },
    },
    async ({ text, remove, fontSize, color, quiet }) => {
      await formatter.setWatermark({ text, remove, fontSize, color })
      if (quiet) return remove ? "Watermark removed" : "Watermark set"
      if (remove) return "Action: Watermark removed\nNext: word_set_watermark({text:\"DRAFT\"}) to add a new watermark"
      return `Action: Watermark set "${text}"\nDetail: Font size: ${fontSize ?? 48}\nNext: word_set_watermark({text:"CONFIDENTIAL", color:"red"}) to change`
    },
  )
}
