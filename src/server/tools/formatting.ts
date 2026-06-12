import { z } from "zod"
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js"
import { WordFormatting } from "../../word/formatting.js"
import { SecurityManager } from "../../security/policy.js"
import { mcpCall } from "./helper.js"

function trackChangeResult(enable: boolean): string {
  return `Action: Track changes ${enable ? "enabled" : "disabled"}\nDetail: All edits will ${enable ? "now be recorded as revisions (red markup)" : "no longer be tracked"}\nNext: word_type_text({text:"..."}) to make tracked edits, or word_set_track_changes({enable:false}) to stop`
}

export function registerFormattingTools(
  server: McpServer,
  formatting: WordFormatting,
  security: SecurityManager,
): void {
  server.registerTool(
    "word_set_font",
    {
      description: "Set font properties. WHEN: after selecting text, or to set forward-typing font. NOT: want to format only part of document? select first.",
      inputSchema: {
        name: z.string().max(100).optional().describe("Font name (e.g. 'Arial', 'SimSun')"),
        size: z.number().min(1).max(1638).optional().describe("Font size in points"),
        bold: z.boolean().optional().describe("Bold"),
        italic: z.boolean().optional().describe("Italic"),
        underline: z.enum(["none", "single", "double", "wavy"]).optional().describe("Underline style"),
        color: z.enum(["auto", "black", "blue", "turquoise", "bright_green", "pink", "red", "yellow", "white", "dark_blue", "teal", "green", "violet", "dark_red", "dark_yellow", "gray_50", "gray_25"]).optional().describe("Font color"),
        strikethrough: z.boolean().optional().describe("Strikethrough"),
        highlightColor: z.enum(["auto", "black", "blue", "turquoise", "bright_green", "pink", "red", "yellow", "white", "dark_blue", "teal", "green", "violet", "dark_red", "dark_yellow", "gray_50", "gray_25"]).optional().describe("Highlight color"),
        superscript: z.boolean().optional().describe("Superscript"),
        subscript: z.boolean().optional().describe("Subscript"),
      },
    },
    mcpCall(security, "word_set_font", async (args) => {
      await formatting.setFont(args)
      const props: string[] = []
      if (args.name) props.push(`font: ${args.name}`)
      if (args.size) props.push(`size: ${args.size}pt`)
      if (args.bold !== undefined) props.push(args.bold ? "bold" : "no bold")
      if (args.italic !== undefined) props.push(args.italic ? "italic" : "no italic")
      if (args.underline) props.push(`underline: ${args.underline}`)
      if (args.color) props.push(`color: ${args.color}`)
      if (args.highlightColor) props.push(`highlight: ${args.highlightColor}`)
      return `Action: Font formatting applied\nDetail: ${props.join(", ")}\nNext: word_set_paragraph({alignment:"center", lineSpacingRule:"double"}) or word_type_text({text:"..."})`
    }),
  )

  server.registerTool(
    "word_set_paragraph",
    {
      description: "Set paragraph formatting. WHEN: cursor is in the paragraph to format. NOT: apply to selected text only? use word_set_font instead.",
      inputSchema: {
        alignment: z.enum(["left", "center", "right", "justify"]).optional().describe("Paragraph alignment"),
        leftIndent: z.number().min(-100).max(100).optional().describe("Left indent in cm"),
        rightIndent: z.number().min(-100).max(100).optional().describe("Right indent in cm"),
        firstLineIndent: z.number().min(-100).max(100).optional().describe("First line indent in cm"),
        spaceBefore: z.number().min(0).max(1584).optional().describe("Space before paragraph in pts"),
        spaceAfter: z.number().min(0).max(1584).optional().describe("Space after paragraph in pts"),
        lineSpacing: z.number().min(0).max(1584).optional().describe("Line spacing in pts"),
        lineSpacingRule: z.enum(["single", "one_point_five", "double", "at_least", "exactly", "multiple"]).optional().describe("Line spacing rule"),
      },
    },
    mcpCall(security, "word_set_paragraph", async (args) => {
      await formatting.setParagraphFormat(args)
      const props: string[] = []
      if (args.alignment) props.push(`align: ${args.alignment}`)
      if (args.leftIndent !== undefined) props.push(`leftIndent: ${args.leftIndent}cm`)
      if (args.rightIndent !== undefined) props.push(`rightIndent: ${args.rightIndent}cm`)
      if (args.firstLineIndent !== undefined) props.push(`firstLineIndent: ${args.firstLineIndent}cm`)
      if (args.spaceBefore !== undefined) props.push(`spaceBefore: ${args.spaceBefore}pt`)
      if (args.spaceAfter !== undefined) props.push(`spaceAfter: ${args.spaceAfter}pt`)
      return `Action: Paragraph formatting applied\nDetail: ${props.join(", ")}\nNext: word_set_font({size:14, name:"Arial"}) or word_type_text({text:"..."})`
    }),
  )

  server.registerTool(
    "word_apply_style",
    {
      description: "Apply a named style to the current paragraph or selection.",
      inputSchema: {
        styleName: z.string().min(1).max(255).describe("Style name (e.g. 'Heading 1', 'Normal', 'Title')"),
      },
    },
    mcpCall(security, "word_apply_style", async ({ styleName }) => {
      await formatting.applyStyle(styleName)
      return `Action: Style '${styleName}' applied\nNext: word_type_text({text:"..."}) or word_set_font({size:12, bold:true}) for overrides`
    }),
  )

  server.registerTool(
    "word_set_page_setup",
    {
      description: "Set page layout options. WHEN: before typing content to ensure correct layout. NOT: already typed content? margins apply to current section only.",
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
    mcpCall(security, "word_set_page_setup", async (args) => {
      await formatting.setPageSetup(args)
      const props: string[] = []
      if (args.topMargin !== undefined) props.push(`top: ${args.topMargin}cm`)
      if (args.bottomMargin !== undefined) props.push(`bottom: ${args.bottomMargin}cm`)
      if (args.leftMargin !== undefined) props.push(`left: ${args.leftMargin}cm`)
      if (args.rightMargin !== undefined) props.push(`right: ${args.rightMargin}cm`)
      if (args.orientation) props.push(`orientation: ${args.orientation}`)
      return `Action: Page setup applied\nDetail: ${props.join(", ")}\nNext: word_set_header({text:"...", alignment:"center"}) or word_set_page_numbers({target:"footer"})`
    }),
  )

  server.registerTool(
    "word_set_properties",
    {
      description: "Set document metadata properties.",
      inputSchema: {
        title: z.string().max(255).optional().describe("Document title"),
        author: z.string().max(255).optional().describe("Author name"),
        subject: z.string().max(255).optional().describe("Subject"),
        keywords: z.string().max(1000).optional().describe("Keywords (comma separated)"),
        comments: z.string().max(5000).optional().describe("Comments/description"),
        category: z.string().max(255).optional().describe("Category"),
      },
    },
    mcpCall(security, "word_set_properties", async (args) => {
      await formatting.setDocumentProperties(args)
      const props: string[] = []
      if (args.title) props.push(`title: ${args.title}`)
      if (args.author) props.push(`author: ${args.author}`)
      return `Action: Document properties updated\nDetail: ${props.join(", ")}\nNext: word_save() or word_type_text({text:"...", mode:"instant"})`
    }),
  )

  server.registerTool(
    "word_list_styles",
    {
      description: "List all in-use styles in the current document.",
    },
    mcpCall(security, "word_list_styles", async () => {
      const styles = await formatting.listStyles()
      const lines = styles.map((s) => `- ${s.name} (${s.builtIn ? "built-in" : "custom"})`)
      return `Action: ${styles.length} style(s) available\n${lines.join("\n")}\nNext: word_apply_style({styleName:"Heading 1"})`
    }),
  )

  server.registerTool(
    "word_set_track_changes",
    {
      description: "Enable or disable Word's Track Changes (revision marking). WHEN: before making edits that need to be reviewed later. NOT: to accept/reject changes, use word_accept_changes or word_reject_changes.",
      inputSchema: {
        enable: z.boolean().describe("true to enable track changes, false to disable"),
      },
    },
    mcpCall(security, "word_set_track_changes", async ({ enable }) => {
      await formatting.setTrackChanges(enable)
      return trackChangeResult(enable)
    }),
  )

  server.registerTool(
    "word_accept_changes",
    {
      description: "Accept all tracked changes in the document.",
    },
    mcpCall(security, "word_accept_changes", async () => {
      const count = await formatting.acceptAllChanges()
      return `Action: ${count} change(s) accepted\nDetail: All revisions accepted into final text\nNext: word_save() or word_set_track_changes({enable:true}) for further edits`
    }),
  )

  server.registerTool(
    "word_reject_changes",
    {
      description: "Reject all tracked changes in the document.",
    },
    mcpCall(security, "word_reject_changes", async () => {
      const count = await formatting.rejectAllChanges()
      return `Action: ${count} change(s) rejected\nDetail: All revisions rejected, document reverted\nNext: word_undo({count:1}) or word_type_text({text:"...", mode:"instant"})`
    }),
  )
}
