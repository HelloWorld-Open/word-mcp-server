import { z } from "zod"
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js"
import { SecurityManager } from "../../security/policy.js"
import type { ServerContext } from "../server-context.js"
import type { WordTableEditor } from "../../word/word-table-editor.js"
import type { WordMediaEditor } from "../../word/word-media-editor.js"
import type { WordDocumentStructure } from "../../word/word-document-structure.js"
import type { WordTextEditor } from "../../word/word-text-editor.js"
import { mcpCall } from "./helper.js"

export function registerManagerTools(
  server: McpServer,
  context: ServerContext,
  tableEditor: WordTableEditor,
  mediaEditor: WordMediaEditor,
  documentStructure: WordDocumentStructure,
  textEditor: WordTextEditor,
  security: SecurityManager,
): void {
  server.registerTool(
    "word_mgr_insert_table",
    {
      description: "Insert a table with auto cursor reset — Phase 4 ELEMENTS — Creates table at cursor then adds paragraph separator.",
      inputSchema: {
        rows: z.number().int().min(1).max(500).describe("Number of rows"),
        columns: z.number().int().min(1).max(100).describe("Number of columns"),
        data: z.array(z.array(z.string().max(100000)).max(100)).max(1000).optional().describe("Optional 2D cell data"),
        autoFitBehavior: z.enum(["fixed", "contents", "window"]).optional().describe("Auto-fit behavior"),
      },
    },
    mcpCall(security, context, "word_mgr_insert_table", async (args) => {
      const result = await tableEditor.insertTable(args)
      return `Table created: ${result.rows}x${result.columns}`
    }),
  )

  server.registerTool(
    "word_mgr_insert_chart",
    {
      description: "Insert a chart with auto cursor reset — Phase 4 ELEMENTS — Inserts chart then paragraph separator.",
      inputSchema: {
        type: z.enum(["column", "bar", "line", "pie", "area"]).describe("Chart type"),
        data: z.array(z.array(z.union([z.string(), z.number()])).min(1)).min(1).max(100).describe("2D data: first row = headers, first column = categories"),
        title: z.string().max(255).optional().describe("Chart title"),
        width: z.number().min(1).max(1000).optional().describe("Width in points"),
        height: z.number().min(1).max(1000).optional().describe("Height in points"),
      },
    },
    mcpCall(security, context, "word_mgr_insert_chart", async (args) => {
      const result = await mediaEditor.insertChart(args)
      return `Chart inserted (${result.type}, ${result.series} series)`
    }),
  )

  server.registerTool(
    "word_mgr_insert_image",
    {
      description: "Insert an image with auto cursor reset — Phase 4 ELEMENTS — Embeds image then paragraph separator.",
      inputSchema: {
        imagePath: z.string().min(1).max(4096).describe("Full path to image file (jpg, png, gif, bmp)"),
        width: z.number().min(1).max(1000).optional().describe("Width in points"),
        height: z.number().min(1).max(1000).optional().describe("Height in points"),
      },
    },
    mcpCall(security, context, "word_mgr_insert_image", async (args) => {
      await mediaEditor.insertImage({
        imagePath: security.pathSanitizer.resolveAndValidate(args.imagePath),
        width: args.width as number | undefined,
        height: args.height as number | undefined,
      })
      return "Image inserted"
    }),
  )

  server.registerTool(
    "word_mgr_insert_textbox",
    {
      description: "Insert a floating text box with auto cursor reset — Phase 4 ELEMENTS.",
      inputSchema: {
        text: z.string().min(1).max(10000).describe("Textbox content"),
        width: z.number().min(1).max(1000).optional().describe("Width in points (default: 200)"),
        height: z.number().min(1).max(1000).optional().describe("Height in points (default: 100)"),
        orientation: z.enum(["horizontal", "vertical"]).optional().describe("Text orientation"),
      },
    },
    mcpCall(security, context, "word_mgr_insert_textbox", async (args) => {
      const result = await mediaEditor.insertTextbox(args)
      return `Textbox inserted (${result.width}x${result.height})`
    }),
  )

  server.registerTool(
    "word_mgr_set_header",
    {
      description: "Set page header with auto context return — Phase 3 GLOBAL — Sets header then returns cursor to main document body.",
      inputSchema: {
        text: z.string().max(5000).describe("Header text"),
        alignment: z.enum(["left", "center", "right"]).optional().describe("Header alignment"),
      },
    },
    mcpCall(security, context, "word_mgr_set_header", async (args) => {
      await documentStructure.setHeader(args.text, args.alignment)
      return "Header set"
    }),
  )

  server.registerTool(
    "word_mgr_set_footer",
    {
      description: "Set page footer with auto context return — Phase 3 GLOBAL — Sets footer then returns cursor to main document body.",
      inputSchema: {
        text: z.string().max(5000).describe("Footer text"),
        alignment: z.enum(["left", "center", "right"]).optional().describe("Footer alignment"),
      },
    },
    mcpCall(security, context, "word_mgr_set_footer", async (args) => {
      await documentStructure.setFooter(args.text, args.alignment)
      return "Footer set"
    }),
  )

  server.registerTool(
    "word_mgr_set_page_numbers",
    {
      description: "Add page numbers with auto context return — Phase 3 GLOBAL — Adds page number field then returns cursor.",
      inputSchema: {
        target: z.enum(["header", "footer"]).describe("Where to place page numbers"),
        alignment: z.enum(["left", "center", "right"]).optional().describe("Page number alignment (default: center)"),
      },
    },
    mcpCall(security, context, "word_mgr_set_page_numbers", async ({ target, alignment }) => {
      await documentStructure.setPageNumbers(target, alignment)
      return "Page numbers added"
    }),
  )

  server.registerTool(
    "word_mgr_set_watermark",
    {
      description: "Add a text watermark with auto context return — Phase 6 FINISH — Adds watermark then returns cursor.",
      inputSchema: {
        text: z.string().min(1).max(100).describe("Watermark text (e.g. DRAFT, CONFIDENTIAL)"),
      },
    },
    mcpCall(security, context, "word_mgr_set_watermark", async ({ text }) => {
      await documentStructure.setWatermark({ text })
      return "Watermark set"
    }),
  )

  server.registerTool(
    "word_mgr_add_bookmark",
    {
      description: "Add a bookmark with auto cursor reset — Phase 5 MARKERS — Adds bookmark then go_to(end)+insert_paragraph.",
      inputSchema: {
        name: z.string().min(1).max(100).regex(/^[a-zA-Z_\u4e00-\u9fff][a-zA-Z0-9_\u4e00-\u9fff]*$/).describe("Bookmark name (alphanumeric + underscore, start with letter/Chinese)"),
      },
    },
    mcpCall(security, context, "word_mgr_add_bookmark", async (args) => {
      await documentStructure.addBookmark(args.name)
      return "Bookmark added"
    }),
  )

  server.registerTool(
    "word_mgr_add_comment",
    {
      description: "Add a comment with auto cursor reset — Phase 5 MARKERS — Adds comment then go_to(end)+insert_paragraph.",
      inputSchema: {
        text: z.string().min(1).max(10000).describe("Comment text"),
      },
    },
    mcpCall(security, context, "word_mgr_add_comment", async (args) => {
      await documentStructure.addComment(args.text)
      return "Comment added"
    }),
  )

  server.registerTool(
    "word_mgr_add_footnote",
    {
      description: "Add a footnote with auto cursor reset — Phase 5 MARKERS — Adds footnote then go_to(end)+insert_paragraph.",
      inputSchema: {
        text: z.string().min(1).max(10000).describe("Footnote text"),
      },
    },
    mcpCall(security, context, "word_mgr_add_footnote", async (args) => {
      await textEditor.addFootnote(args.text)
      return "Footnote added"
    }),
  )

  server.registerTool(
    "word_mgr_add_hyperlink",
    {
      description: "Add a hyperlink with auto cursor reset — Phase 5 MARKERS — Adds hyperlink then go_to(end)+insert_paragraph.",
      inputSchema: {
        text: z.string().min(1).max(1000).describe("Display text"),
        address: z.string().min(1).max(2083).describe("URL or file path"),
        subAddress: z.string().max(255).optional().describe("Anchor or bookmark within the document"),
        screenTip: z.string().max(500).optional().describe("Tooltip on hover"),
      },
    },
    mcpCall(security, context, "word_mgr_add_hyperlink", async (args) => {
      await textEditor.addHyperlink(args.text, args.address, args.subAddress, args.screenTip)
      return "Hyperlink added"
    }),
  )

  server.registerTool(
    "word_mgr_insert_section_break",
    {
      description: "Insert a section break with auto cursor reset — Phase 6.",
      inputSchema: {
        type: z.enum(["nextPage", "continuous", "evenPage", "oddPage"]).optional().describe("Section break type (default: nextPage)"),
      },
    },
    mcpCall(security, context, "word_mgr_insert_section_break", async (args) => {
      await textEditor.insertSectionBreak(args.type)
      return "Section break inserted"
    }),
  )

}
