import { z } from "zod"
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js"
import { WordDocumentStructure } from "../../word/word-document-structure.js"
import { SecurityManager } from "../../security/policy.js"
import type { ServerContext } from "../server-context.js"
import { mcpCall } from "./helper.js"

export function registerStructureTools(
  server: McpServer,
  context: ServerContext,
  content: WordDocumentStructure,
  security: SecurityManager,
): void {
  server.registerTool(
    "word_set_header",
    {
      description: "Set the header text. WHEN: every page needs repeating top text. NOT: want page numbers only? use word_set_page_numbers.",
      inputSchema: {
        text: z.string().max(5000).describe("Header text"),
        alignment: z.enum(["left", "center", "right"]).optional().describe("Header alignment"),
      },
    },
    mcpCall(security, context, "word_set_header", async ({ text, alignment }) => {
      await content.setHeader(text, alignment)
      return `Action: Header set\nDetail: Alignment: ${alignment ?? "left"}\nNext: word_set_footer({text:"Page "}) or word_set_page_numbers({target:"footer"})`
    }),
  )

  server.registerTool(
    "word_set_footer",
    {
      description: "Set the footer text. WHEN: every page needs repeating bottom text. NOT: want page numbers only? use word_set_page_numbers.",
      inputSchema: {
        text: z.string().max(5000).describe("Footer text"),
        alignment: z.enum(["left", "center", "right"]).optional().describe("Footer alignment"),
      },
    },
    mcpCall(security, context, "word_set_footer", async ({ text, alignment }) => {
      await content.setFooter(text, alignment)
      return `Action: Footer set\nDetail: Alignment: ${alignment ?? "left"}\nNext: word_set_header({text:"..."}) or word_set_page_numbers({target:"footer"})`
    }),
  )

  server.registerTool(
    "word_set_page_numbers",
    {
      description: "Add page numbers to header or footer.",
      inputSchema: {
        target: z.enum(["header", "footer"]).describe("Where to place page numbers"),
        alignment: z.enum(["left", "center", "right"]).optional().describe("Page number alignment (default: center)"),
      },
    },
    mcpCall(security, context, "word_set_page_numbers", async ({ target, alignment }) => {
      await content.setPageNumbers(target, alignment)
      return `Action: Page numbers added to ${target} (${alignment ?? "center"})\nNext: word_set_header({text:"... - Page "}) for custom text alongside numbers`
    }),
  )

  server.registerTool(
    "word_insert_toc",
    {
      description: "Insert a Table of Contents at the cursor position.",
    },
    mcpCall(security, context, "word_insert_toc", async () => {
      await content.insertToc()
      return "Action: Table of Contents inserted\nNext: Right-click the TOC > Update Field to refresh after adding headings"
    }),
  )

  server.registerTool(
    "word_add_bookmark",
    {
      description: "Add a bookmark at the current cursor position.",
      inputSchema: {
        name: z.string().min(1).max(100).regex(/^[a-zA-Z_\u4e00-\u9fff][a-zA-Z0-9_\u4e00-\u9fff]*$/).describe("Bookmark name (alphanumeric + underscore, start with letter/Chinese)"),
      },
    },
    mcpCall(security, context, "word_add_bookmark", async ({ name }) => {
      await content.addBookmark(name)
      return `Action: Bookmark added "${name}"\nNext: word_add_hyperlink with subAddress="${name}" to link here`
    }),
  )

  server.registerTool(
    "word_add_comment",
    {
      description: "Add a comment at the current selection.",
      inputSchema: {
        text: z.string().min(1).max(10000).describe("Comment text"),
      },
    },
    mcpCall(security, context, "word_add_comment", async ({ text }) => {
      await content.addComment(text)
      const preview = text.slice(0, 80) + (text.length > 80 ? "..." : "")
      return `Action: Comment added\nDetail: "${preview}"\nNext: word_type_text({text:"..."}) to continue editing`
    }),
  )

  server.registerTool(
    "word_set_watermark",
    {
      description: "Add or remove a text watermark (e.g. 'DRAFT', 'CONFIDENTIAL').",
      inputSchema: {
        text: z.string().min(1).max(100).describe("Watermark text"),
        remove: z.boolean().optional().describe("Set true to remove existing watermarks"),
        fontSize: z.number().int().min(12).max(200).optional().describe("Font size (default: 48)"),
        color: z.enum(["auto", "black", "blue", "turquoise", "bright_green", "pink", "red", "yellow", "white", "dark_blue", "teal", "green", "violet", "dark_red", "dark_yellow", "gray_50", "gray_25"]).optional().describe("Watermark color"),
      },
    },
    mcpCall(security, context, "word_set_watermark", async ({ text, remove, fontSize, color }) => {
      await content.setWatermark({ text, remove, fontSize, color })
      if (remove) return "Action: Watermark removed\nNext: word_set_watermark({text:\"DRAFT\"}) to add a new watermark"
      return `Action: Watermark set "${text}"\nDetail: Font size: ${fontSize ?? 48}\nNext: word_set_watermark({text:"CONFIDENTIAL", color:"red"}) to change`
    }),
  )
}
