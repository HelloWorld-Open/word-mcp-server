import { z } from "zod"
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js"
import { SecurityManager } from "../../security/policy.js"
import type { ServerContext } from "../server-context.js"
import type { StreamingMarkdownWriter } from "../../word/word-stream-writer.js"
import { createRegTool, ColorSchema } from "./shared.js"

export function registerStreamTools(
  server: McpServer,
  context: ServerContext,
  streamWriter: StreamingMarkdownWriter,
  security: SecurityManager,
): void {
  const regTool = createRegTool(server, security, context)
  regTool("word_stream_start",
    {
      description: "WHEN: creating a new document from scratch. WHAT: starts a streaming session — creates a new document with optional title, author, page setup, and base style configuration. Content is written incrementally with word_stream_block. CONSTRAINT: auto-closes any existing active document. For editing existing files, use word_document instead.",
      inputSchema: {
        title: z.string().max(255).optional().describe("Document title (set in document properties and title bar)"),
        author: z.string().max(255).optional().describe("Document author (set in document properties)"),
        templatePath: z.string().max(4096).optional().describe("Full path to .dotx or .dotm template file. Styles and content from the template are inherited."),
        orientation: z.enum(["portrait", "landscape"]).optional().describe("Page orientation (default: portrait)"),
        topMargin: z.number().min(0).max(100).optional().describe("Top margin in cm (default: 2.54)"),
        bottomMargin: z.number().min(0).max(100).optional().describe("Bottom margin in cm (default: 2.54)"),
        leftMargin: z.number().min(0).max(100).optional().describe("Left margin in cm (default: 3.17 for binding)"),
        rightMargin: z.number().min(0).max(100).optional().describe("Right margin in cm (default: 3.17)"),
        baseStyleProfile: z.record(
          z.string().max(100),
          z.object({
            font: z.object({
              name: z.string().max(100).optional().describe("Font name (e.g., 'SimSun' for Chinese body, 'Calibri' for English)"),
              size: z.number().min(1).max(1638).optional().describe("Font size in points (e.g., 12 for body, 16 for heading)"),
              bold: z.boolean().optional().describe("Bold"),
              italic: z.boolean().optional().describe("Italic"),
              color: ColorSchema.optional().describe("Font color"),
              underline: z.enum(["none", "single", "double", "wavy"]).optional().describe("Underline style"),
              strikethrough: z.boolean().optional().describe("Strikethrough"),
              highlight: z.string().max(20).optional().describe("Highlight color: enum name (e.g., 'yellow') or hex (e.g., '#FFF0E0')"),
            }).optional(),
            paragraph: z.object({
              alignment: z.enum(["left", "center", "right", "justify"]).optional().describe("Paragraph alignment"),
              firstLineIndent: z.number().min(-100).max(100).optional().describe("First line indent in cm (0.74 ≈ 2 Chinese chars at 12pt)"),
              spaceBefore: z.number().min(0).max(1584).optional().describe("Space before paragraph in points"),
              spaceAfter: z.number().min(0).max(1584).optional().describe("Space after paragraph in points"),
              lineSpacing: z.number().min(0).max(1584).optional().describe("Line spacing value. With 'multiple' rule: 1.5=1.5x spacing."),
              lineSpacingRule: z.enum(["single", "one_point_five", "double", "at_least", "exactly", "multiple"]).optional().describe("Line spacing rule (default: 'multiple' with lineSpacing=1.15)"),
              keepWithNext: z.boolean().optional().describe("Keep this paragraph with the next (prevents page break between them)"),
              pageBreakBefore: z.boolean().optional().describe("Always start this paragraph on a new page"),
              borders: z.object({
                style: z.enum(["none", "single", "dot", "dash", "double"]).describe("Border line style"),
                color: ColorSchema.optional().describe("Border color"),
                size: z.number().min(1).max(48).optional().describe("Line width in quarter-points (8 = 1pt)"),
                sides: z.array(z.enum(["top", "bottom", "left", "right"])).optional().describe("Which sides to apply border (default: all four)"),
              }).optional(),
            }).optional(),
          }),
        ).optional().describe("Pre-configure built-in styles (Normal, Heading 1-9, Title, etc.). Font and paragraph settings defined here are inherited by all content using those styles — zero per-block COM overhead."),
      },
    },
    async (args) => {
      const safeArgs = args.templatePath
        ? { ...args, templatePath: security.pathSanitizer.resolveAndValidate(args.templatePath) }
        : args
      return await streamWriter.start(safeArgs)
    },
    { preconditions: [] },
  )

  regTool("word_stream_block",
    {
      description: "WHEN: after word_stream_start, to write document content incrementally in chapters/sections. WHAT: writes a Markdown content block into the active streaming session — content appears in Word in real time. CONSTRAINT: requires an active streaming session created by word_stream_start. Supports headings, bold, italic, lists, tables, code blocks, blockquotes.",
      inputSchema: {
        text: z.string().min(1).max(100000).describe("Markdown content to write. Send chapters/sections one at a time. Supports: # H1, ## H2, **bold**, *italic*, - lists, 1. numbered, |table|, ```code```, > quote, --- hr."),
      },
    },
    async (args) => {
      const result = await streamWriter.writeBlock(args.text)
      return `Written ${result.chars} chars (${result.blockType}), total ${result.blockIndex} blocks`
    },
  )

  regTool("word_stream_end",
    {
      description: "WHEN: finished writing all content with word_stream_block calls. WHAT: ends the streaming session, saves the document, and optionally exports to PDF. CONSTRAINT: after this call, no more word_stream_block calls are allowed. Use word_document to reopen the saved file for further editing.",
      inputSchema: {
        save: z.boolean().optional().describe("Whether to save the document (default: true). Set false to discard the streaming session."),
        exportPath: z.string().max(4096).optional().describe("Optional PDF export path (e.g., 'C:\\output\\report.pdf'). Only valid when save=true."),
      },
    },
    async (args) => {
      const safePath = args.exportPath
        ? security.pathSanitizer.validateForWrite(args.exportPath)
        : undefined
      const result = await streamWriter.end({ save: args.save, exportPath: safePath })
      const lines: string[] = [`Stream session ended`]
      lines.push(`Total ${result.blockCount} blocks, ${result.charCount} chars, ${result.elapsed}ms`)
      if (result.savedPath) lines.push(`Saved to: ${result.savedPath}`)
      if (result.pdfPath) lines.push(`PDF exported to: ${result.pdfPath}`)
      return lines.join("\n")
    },
  )
}
