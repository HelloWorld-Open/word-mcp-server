import { z } from "zod"
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js"
import { WordFormatter } from "../../word/word-formatter.js"
import { SecurityManager } from "../../security/policy.js"
import type { ServerContext } from "../server-context.js"
import { createRegTool, ColorSchema } from "./shared.js"

function trackChangeResult(enable: boolean): string {
  return `Action: Track changes ${enable ? "enabled" : "disabled"}\nDetail: All edits will ${enable ? "now be recorded as revisions (red markup)" : "no longer be tracked"}\nNext: word_type_text({text:"..."}) to make tracked edits, or word_set_track_changes({enable:false}) to stop`
}

export function registerFormattingTools(
  server: McpServer,
  context: ServerContext,
  formatter: WordFormatter,
  security: SecurityManager,
): void {
  const regTool = createRegTool(server, security, context)
  regTool("word_set_font",
    {
      description: "WHEN: need to change text appearance (font, size, bold, color, etc.) on selected text or set forward-typing style. WHAT: applies font formatting to current selection, or sets the default for subsequently typed text if nothing is selected. CONSTRAINT: must select text first for retroactive formatting; without selection only affects NEW text after cursor.",
      inputSchema: {
        name: z.string().max(100).optional().describe("Font name (e.g. 'Arial', 'SimSun', 'Times New Roman', 'Calibri')"),
        size: z.number().min(1).max(1638).optional().describe("Font size in points (e.g. 12 for body text, 16 for headings)"),
        bold: z.boolean().optional().describe("Bold formatting"),
        italic: z.boolean().optional().describe("Italic formatting"),
        underline: z.enum(["none", "single", "double", "wavy"]).optional().describe("Underline style"),
        color: ColorSchema.optional().describe("Font color (enumerated Word color)"),
        strikethrough: z.boolean().optional().describe("Strikethrough"),
        highlightColor: ColorSchema.optional().describe("Highlight (marker) color"),
        superscript: z.boolean().optional().describe("Superscript (e.g., for footnotes or exponents)"),
        subscript: z.boolean().optional().describe("Subscript (e.g., for chemical formulas)"),
      },
    },
    async (args) => {
      await formatter.setFont(args)
      const props: string[] = []
      if (args.name) props.push(`font: ${args.name}`)
      if (args.size) props.push(`size: ${args.size}pt`)
      if (args.bold !== undefined) props.push(args.bold ? "bold" : "no bold")
      if (args.italic !== undefined) props.push(args.italic ? "italic" : "no italic")
      if (args.underline) props.push(`underline: ${args.underline}`)
      if (args.color) props.push(`color: ${args.color}`)
      if (args.highlightColor) props.push(`highlight: ${args.highlightColor}`)
      return `Action: Font formatting applied\nDetail: ${props.join(", ")}\nNext: word_set_paragraph({alignment:"center", lineSpacingRule:"double"}) or word_type_text({text:"..."})`
    },
  )

  regTool("word_set_paragraph",
    {
      description: "WHEN: need to adjust paragraph layout (alignment, indentation, line spacing, spacing before/after). WHAT: sets formatting for the current paragraph(s). CONSTRAINT: applies to the paragraph containing the cursor. For document-wide defaults, configure Normal style via word_stream_start baseStyleProfile.",
      inputSchema: {
        alignment: z.enum(["left", "center", "right", "justify"]).optional().describe("Paragraph alignment"),
        leftIndent: z.number().min(-100).max(100).optional().describe("Left indent in cm (negative values for outdenting/hanging)"),
        rightIndent: z.number().min(-100).max(100).optional().describe("Right indent in cm"),
        firstLineIndent: z.number().min(-100).max(100).optional().describe("First line indent in cm (e.g., 0.74 ≈ 2 Chinese characters at 12pt)"),
        spaceBefore: z.number().min(0).max(1584).optional().describe("Space before paragraph in points (12pt ≈ one blank line)"),
        spaceAfter: z.number().min(0).max(1584).optional().describe("Space after paragraph in points"),
        lineSpacing: z.number().min(0).max(1584).optional().describe("Line spacing value. With 'multiple' rule: 1.5=1.5x, 2=double. With 'exactly': value in points."),
        lineSpacingRule: z.enum(["single", "one_point_five", "double", "at_least", "exactly", "multiple"]).optional().describe("Line spacing rule: 'single'=default, 'multiple'=multiplier in lineSpacing, 'exactly'=fixed pts"),
      },
    },
    async (args) => {
      await formatter.setParagraphFormat(args)
      const props: string[] = []
      if (args.alignment) props.push(`align: ${args.alignment}`)
      if (args.leftIndent !== undefined) props.push(`leftIndent: ${args.leftIndent}cm`)
      if (args.rightIndent !== undefined) props.push(`rightIndent: ${args.rightIndent}cm`)
      if (args.firstLineIndent !== undefined) props.push(`firstLineIndent: ${args.firstLineIndent}cm`)
      if (args.spaceBefore !== undefined) props.push(`spaceBefore: ${args.spaceBefore}pt`)
      if (args.spaceAfter !== undefined) props.push(`spaceAfter: ${args.spaceAfter}pt`)
      return `Action: Paragraph formatting applied\nDetail: ${props.join(", ")}\nNext: word_set_font({size:14, name:"Arial"}) or word_type_text({text:"..."})`
    },
  )

  regTool("word_apply_style",
    {
      description: "WHEN: need to apply Word's built-in paragraph styles (Heading 1, Title, Normal, etc.) to content. WHAT: sets the paragraph style of the current selection or cursor position. CONSTRAINT: only built-in styles work; custom styles created in Word are not listed. For individual font/paragraph properties, use word_set_font or word_set_paragraph.",
      inputSchema: {
        styleName: z.string().min(1).max(255).describe("Style name (e.g. 'Heading 1', 'Heading 2', 'Normal', 'Title', 'Subtitle', 'Quote')"),
      },
    },
    async ({ styleName }) => {
      await formatter.applyStyle(styleName)
      return `Action: Style '${styleName}' applied\nNext: word_type_text({text:"..."}) or word_set_font({size:12, bold:true}) for overrides`
    },
  )

  regTool("word_set_page_setup",
    {
      description: "WHEN: need to set page dimensions, margins, or orientation before/after writing content. WHAT: configures page layout for the current section (margins, paper size, orientation). CONSTRAINT: margins apply to current section only; use word_insert_section_break first for different layouts per section.",
      inputSchema: {
        topMargin: z.number().min(0).max(100).optional().describe("Top margin in cm (default: 2.54cm / 1 inch)"),
        bottomMargin: z.number().min(0).max(100).optional().describe("Bottom margin in cm"),
        leftMargin: z.number().min(0).max(100).optional().describe("Left margin in cm"),
        rightMargin: z.number().min(0).max(100).optional().describe("Right margin in cm"),
        orientation: z.enum(["portrait", "landscape"]).optional().describe("Page orientation: portrait (vertical) or landscape (horizontal)"),
        pageWidth: z.number().min(5).max(100).optional().describe("Page width in cm (default: 21cm for A4 portrait)"),
        pageHeight: z.number().min(5).max(100).optional().describe("Page height in cm (default: 29.7cm for A4)"),
      },
    },
    async (args) => {
      await formatter.setPageSetup(args)
      const props: string[] = []
      if (args.topMargin !== undefined) props.push(`top: ${args.topMargin}cm`)
      if (args.bottomMargin !== undefined) props.push(`bottom: ${args.bottomMargin}cm`)
      if (args.leftMargin !== undefined) props.push(`left: ${args.leftMargin}cm`)
      if (args.rightMargin !== undefined) props.push(`right: ${args.rightMargin}cm`)
      if (args.orientation) props.push(`orientation: ${args.orientation}`)
      return `Action: Page setup applied\nDetail: ${props.join(", ")}\nNext: word_set_page_region({target:"header", text:"...", alignment:"center"}) or word_set_page_numbers({target:"footer"})`
    },
  )

  regTool("word_set_properties",
    {
      description: "WHEN: need to fill in document metadata for search/filing/organization purposes. WHAT: sets document-level properties (title, author, subject, keywords, comments, category). CONSTRAINT: metadata is embedded in the .docx file; visible in File > Info. Does NOT affect visible document content.",
      inputSchema: {
        title: z.string().max(255).optional().describe("Document title (appears in file properties, search results)"),
        author: z.string().max(255).optional().describe("Author name"),
        subject: z.string().max(255).optional().describe("Subject or category description"),
        keywords: z.string().max(1000).optional().describe("Keywords for search (comma-separated, e.g., 'report, Q3, financial')"),
        comments: z.string().max(5000).optional().describe("Comments/description for the document"),
        category: z.string().max(255).optional().describe("Category (e.g., 'Report', 'Proposal', 'Invoice')"),
      },
    },
    async (args) => {
      await formatter.setDocumentProperties(args)
      const props: string[] = []
      if (args.title) props.push(`title: ${args.title}`)
      if (args.author) props.push(`author: ${args.author}`)
      return `Action: Document properties updated\nDetail: ${props.join(", ")}\nNext: word_save() or word_type_text({text:"...", mode:"instant"})`
    },
  )

  regTool("word_list_styles",
    {
      description: "WHEN: need to see what styles are available in the document before applying one. WHAT: lists all in-use styles (both built-in and custom) with their type. CONSTRAINT: read-only; does not modify document. Use before word_apply_style to discover available style names.",
    },
    async () => {
      const styles = await formatter.listStyles()
      const lines = styles.map((s) => `- ${s.name} (${s.builtIn ? "built-in" : "custom"})`)
      return `Action: ${styles.length} style(s) available\n${lines.join("\n")}\nNext: word_apply_style({styleName:"Heading 1"})`
    },
  )

  regTool("word_set_body_indent",
    {
      description: "WHEN: formatting a document (especially Chinese academic papers) where each body paragraph needs standard first-line indent. WHAT: applies first-line indent to all paragraphs using the 'Normal' style. CONSTRAINT: only affects 'Normal' style paragraphs; explicitly-formatted paragraphs are skipped. For individual paragraphs, use word_set_paragraph({firstLineIndent:...}).",
      inputSchema: {
        indent: z.number().min(0).max(10).describe("First line indent in cm. For Chinese 12pt font: 0.74cm ≈ 2 characters. For English 12pt: ~1.27cm ≈ 5 spaces."),
      },
    },
    async ({ indent }) => {
      const indentCm = indent ?? 0.74
      const count = await formatter.applyBodyIndent(indentCm)
      return `Action: Body indent applied (indent=${indentCm}cm)\nDetail: ${count} paragraph(s) processed\nNext: word_set_paragraph({firstLineIndent:0.74}) for individual paragraphs`
    },
  )

  regTool("word_set_track_changes",
    {
      description: "WHEN: before making edits that need to be reviewed later in a collaborative workflow. WHAT: enables or disables Word's Track Changes (revision marking). CONSTRAINT: when enabled, all insertions/deletions are marked in red/underline for review. To finalize (accept/reject all changes), use word_track_changes_apply.",
      inputSchema: {
        enable: z.boolean().describe("true to enable track changes (revisions recorded), false to disable (edits apply directly)"),
      },
    },
    async ({ enable }) => {
      await formatter.setTrackChanges(enable)
      return trackChangeResult(enable)
    },
  )

  regTool("word_track_changes_apply",
    {
      description: "WHEN: finished reviewing tracked changes and want to finalize the document. WHAT: action=accept applies all revisions; action=reject discards all revisions and reverts to original text. CONSTRAINT: affects ALL tracked changes at once; there is no per-change selection. Cannot be undone after word_save.",
      inputSchema: {
        action: z.enum(["accept", "reject"]).describe("'accept' to apply all tracked changes into final text, 'reject' to discard all changes and revert to original"),
      },
    },
    async ({ action }) => {
      if (action === "accept") {
        const count = await formatter.acceptAllChanges()
        return `Action: ${count} change(s) accepted\nDetail: All revisions accepted into final text\nNext: word_save() or word_set_track_changes({enable:true}) for further edits`
      }
      const count = await formatter.rejectAllChanges()
      return `Action: ${count} change(s) rejected\nDetail: All revisions rejected, document reverted\nNext: word_undo_redo({action:"undo"}) or word_type_text({text:"...", mode:"instant"})`
    },
  )
}
