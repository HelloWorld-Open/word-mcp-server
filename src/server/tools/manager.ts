import { z } from "zod"
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js"
import { SecurityManager } from "../../security/policy.js"
import { WordDocumentManager } from "../../manager/word-manager.js"
import { mcpCall } from "./helper.js"

export function registerManagerTools(
  server: McpServer,
  mgr: WordDocumentManager,
  security: SecurityManager,
): void {
  server.registerTool(
    "word_mgr_create_document",
    {
      description: "Create a new document with optional page setup in one call — Phase 1 INIT — Automatically closes any active document first.",
      inputSchema: {
        title: z.string().max(255).optional().describe("Document title"),
        author: z.string().max(255).optional().describe("Document author"),
        topMargin: z.number().min(0).max(100).optional().describe("Top margin in cm"),
        bottomMargin: z.number().min(0).max(100).optional().describe("Bottom margin in cm"),
        leftMargin: z.number().min(0).max(100).optional().describe("Left margin in cm"),
        rightMargin: z.number().min(0).max(100).optional().describe("Right margin in cm"),
        orientation: z.enum(["portrait", "landscape"]).optional().describe("Page orientation"),
      },
    },
    mcpCall(security, "word_mgr_create_document", async (args) => {
      return await mgr.createDocument(args)
    }, { preconditions: [] }),
  )

  server.registerTool(
    "word_mgr_write_content",
    {
      description: "Write formatted Markdown content and auto-reset cursor — Phase 2 MAIN — 80%+ body content in one call — Handles cursor reset automatically.",
      inputSchema: {
        text: z.string().min(1).max(100000).describe("Markdown content (headings, bold, italic, lists, tables, code)"),
      },
    },
    mcpCall(security, "word_mgr_write_content", async (args) => {
      return await mgr.writeContent(args)
    }),
  )

  server.registerTool(
    "word_mgr_apply_heading",
    {
      description: "Apply a heading at end of document with auto cursor reset — Phase 2 MAIN — Writes text, applies style, inserts paragraph.",
      inputSchema: {
        text: z.string().min(1).max(1000).describe("Heading text"),
        level: z.number().int().min(1).max(9).describe("Heading level (1-9)"),
      },
    },
    mcpCall(security, "word_mgr_apply_heading", async ({ text, level }) => {
      return await mgr.applyHeading(text, level)
    }),
  )

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
    mcpCall(security, "word_mgr_insert_table", async (args) => {
      return await mgr.insertTable(args)
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
    mcpCall(security, "word_mgr_insert_chart", async (args) => {
      return await mgr.insertChart(args)
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
    mcpCall(security, "word_mgr_insert_image", async (args) => {
      const safeArgs = {
        ...args,
        imagePath: security.pathSanitizer.resolveAndValidate(args.imagePath),
      }
      return await mgr.insertImage(safeArgs)
    }),
  )

  server.registerTool(
    "word_mgr_insert_list",
    {
      description: "Insert a bullet or numbered list with auto cursor reset — Phase 4 ELEMENTS.",
      inputSchema: {
        type: z.enum(["bullet", "number"]).describe("List type"),
        items: z.array(z.string().max(100000)).min(1).max(500).describe("List items"),
      },
    },
    mcpCall(security, "word_mgr_insert_list", async (args) => {
      return await mgr.insertList(args)
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
    mcpCall(security, "word_mgr_insert_textbox", async (args) => {
      return await mgr.insertTextbox(args)
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
    mcpCall(security, "word_mgr_set_header", async (args) => {
      return await mgr.setHeader(args)
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
    mcpCall(security, "word_mgr_set_footer", async (args) => {
      return await mgr.setFooter(args)
    }),
  )

  server.registerTool(
    "word_mgr_set_page_numbers",
    {
      description: "Add page numbers with auto context return — Phase 3 GLOBAL — Adds page number field then returns cursor.",
      inputSchema: {
        target: z.enum(["header", "footer"]).describe("Where to place page numbers"),
      },
    },
    mcpCall(security, "word_mgr_set_page_numbers", async ({ target }) => {
      return await mgr.setPageNumbers(target)
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
    mcpCall(security, "word_mgr_set_watermark", async ({ text }) => {
      return await mgr.setWatermark(text)
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
    mcpCall(security, "word_mgr_add_bookmark", async (args) => {
      return await mgr.addBookmark(args)
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
    mcpCall(security, "word_mgr_add_comment", async (args) => {
      return await mgr.addComment(args)
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
    mcpCall(security, "word_mgr_add_footnote", async (args) => {
      return await mgr.addFootnote(args)
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
    mcpCall(security, "word_mgr_add_hyperlink", async (args) => {
      return await mgr.addHyperlink(args)
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
    mcpCall(security, "word_mgr_insert_section_break", async (args) => {
      return await mgr.insertSectionBreak(args)
    }),
  )

  server.registerTool(
    "word_mgr_format_page",
    {
      description: "Format page layout (margins, orientation, size) — Phase 1 INIT.",
      inputSchema: {
        topMargin: z.number().min(0).max(100).optional().describe("Top margin in cm"),
        bottomMargin: z.number().min(0).max(100).optional().describe("Bottom margin in cm"),
        leftMargin: z.number().min(0).max(100).optional().describe("Left margin in cm"),
        rightMargin: z.number().min(0).max(100).optional().describe("Right margin in cm"),
        orientation: z.enum(["portrait", "landscape"]).optional().describe("Page orientation"),
        pageWidth: z.number().min(5).max(100).optional().describe("Page width in cm"),
        pageHeight: z.number().min(5).max(100).optional().describe("Page height in cm"),
      },
    },
    mcpCall(security, "word_mgr_format_page", async (args) => {
      return await mgr.formatPage(args)
    }),
  )

  server.registerTool(
    "word_mgr_save",
    {
      description: "Save the document (and optionally export to PDF) — Phase 6 FINISH.",
      inputSchema: {
        exportPath: z.string().max(4096).optional().describe("Optional PDF export path"),
      },
    },
    mcpCall(security, "word_mgr_save", async ({ exportPath }) => {
      const safePath = exportPath ? security.pathSanitizer.validateForWrite(exportPath) : undefined
      return await mgr.saveAndExport(safePath)
    }),
  )

  server.registerTool(
    "word_mgr_close",
    {
      description: "Close the current document with optional save — Phase 6 FINISH.",
      inputSchema: {
        save: z.boolean().optional().describe("Save before closing (default: false)"),
      },
    },
    mcpCall(security, "word_mgr_close", async ({ save }) => {
      return await mgr.closeDocument(save ?? false)
    }),
  )
}
