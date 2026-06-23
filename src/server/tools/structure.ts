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
      description: "WHEN: every page needs repeating text at the top (header) or bottom (footer). WHAT: sets header or footer text with optional alignment for the current or specified section. CONSTRAINT: headers/footers repeat on every page of the section. For page numbers only, use word_set_page_numbers which is simpler.",
      inputSchema: {
        target: z.enum(["header", "footer"]).describe("'header' for top of page, 'footer' for bottom of page"),
        text: z.string().max(5000).describe("Header/footer text content (plain text)"),
        alignment: z.enum(["left", "center", "right"]).optional().describe("Text alignment within the header/footer region"),
        sectionIndex: z.number().int().min(1).optional().describe("Section index (1-based, default: last/current section)"),
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
      description: "WHEN: every page needs automatic page numbering. WHAT: inserts page number field in header or footer with optional alignment for the specified section. CONSTRAINT: page numbers are auto-generated (update when pages change). For custom text alongside page numbers, use word_set_page_region with page number field.",
      inputSchema: {
        target: z.enum(["header", "footer"]).describe("Where to place page numbers (header=top, footer=bottom)"),
        alignment: z.enum(["left", "center", "right"]).optional().describe("Page number alignment (default: center)"),
        sectionIndex: z.number().int().min(1).optional().describe("Section index (1-based, default: last/current section)"),
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
      description: "WHEN: document has headings and needs an auto-generated table of contents. WHAT: inserts a Table of Contents field at the cursor position that builds from document headings. CONSTRAINT: TOC is a field that requires manual update (right-click > Update Field) after headings change. For manual item lists, use word_insert_list instead.",
    },
    async () => {
      await formatter.insertToc()
      return "Action: Table of Contents inserted\nNext: Right-click the TOC > Update Field to refresh after adding headings"
    },
  )

  regTool("word_add_bookmark",
    {
      description: "WHEN: need to create a navigation anchor for hyperlinks or cross-references within the document. WHAT: adds a bookmark at the current cursor position with a unique name. CONSTRAINT: name is permanent; rename by deleting and re-adding. Use with word_add_hyperlink subAddress or word_go_to({what:'bookmark'}) to jump here.",
      inputSchema: {
        name: z.string().min(1).max(100).regex(/^[a-zA-Z_\u4e00-\u9fff][a-zA-Z0-9_\u4e00-\u9fff]*$/).describe("Bookmark name (alphanumeric + underscore, start with letter or Chinese character). Example: 'section2_intro'"),
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
      description: "WHEN: need to add review notes or feedback visible in the Review pane (like Google Docs comments). WHAT: adds a comment linked to the current text selection. CONSTRAINT: comments appear in the Review pane and are visible in print layout. For printed page-bottom notes, use word_add_footnote instead.",
      inputSchema: {
        text: z.string().min(1).max(10000).describe("Comment text (review note)"),
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
      description: "WHEN: need to indicate document status (DRAFT, CONFIDENTIAL, SAMPLE) with a background watermark. WHAT: adds or removes a diagonal text watermark across every page. CONSTRAINT: only one watermark at a time; setting a new one replaces the old. Watermarks appear behind text. For repeating text in a specific region, use headers/footers instead.",
      inputSchema: {
        text: z.string().min(1).max(100).describe("Watermark text (e.g., 'DRAFT', 'CONFIDENTIAL', 'SAMPLE', 'URGENT')"),
        remove: z.boolean().optional().describe("Set true to remove existing watermarks (text is ignored)"),
        fontSize: z.number().int().min(12).max(200).optional().describe("Font size in points (default: 48). Larger = more prominent."),
        color: ColorSchema.optional().describe("Watermark text color (default: gray)"),
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
