import { z } from "zod"
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js"
import { WordMarkdown } from "../../word/word-markdown.js"
import { WordTextEditor } from "../../word/word-text-editor.js"
import { PositionMap, type Locator } from "../../word/position-map.js"
import { SecurityManager } from "../../security/policy.js"
import { mcpCall } from "./helper.js"

const locatorFields = {
  by: z.enum(["heading", "paragraph", "table", "bookmark"]).optional().describe("Target type (default: heading)"),
  match: z.string().max(5000).optional().describe("Text to match (for heading/paragraph)"),
  matchMode: z.enum(["exact", "contains", "startsWith", "regex"]).optional().describe("Matching mode (default: exact)"),
  occurrence: z.number().int().min(1).max(1000).optional().describe("Which occurrence to target (1-based, default: 1)"),
  offsetDirection: z.enum(["before", "after"]).optional().describe("Offset direction from the matched element"),
  offsetCount: z.number().int().min(1).max(1000).optional().describe("Number of paragraphs to offset (default: 1)"),
  name: z.string().min(1).max(255).optional().describe("Bookmark name (required when by='bookmark')"),
} as const

function formatPosition(pos: { found: boolean; paragraphIndex?: number; error?: string; headingContext?: string | null; tableIndex?: number }): string {
  if (!pos.found) return `Location not found: ${pos.error ?? "unknown reason"}`
  const parts = [`Paragraph ${pos.paragraphIndex}`]
  if (pos.headingContext) parts.push(`Under: ${pos.headingContext}`)
  if (pos.tableIndex != null) parts.push(`Table #${pos.tableIndex}`)
  return parts.join(" | ")
}

function splitMarkdown(text: string, maxCharLen: number): string[] {
  if (text.length <= maxCharLen) return [text]
  const parts = text.split(/\n(?=#{1,6}\s)/)
  const chunks: string[] = []
  let current = ""
  for (const part of parts) {
    const candidate = current ? current + "\n" + part : part
    if (candidate.length > maxCharLen && current.length > 0) {
      chunks.push(current)
      current = part
    } else {
      current = candidate
    }
  }
  if (current) chunks.push(current)
  return chunks
}

function sleep(ms: number): Promise<void> {
  return new Promise(resolve => setTimeout(resolve, ms))
}

function locatorFromArgs(args: Record<string, unknown>): Locator {
  const by = (args.by as string) ?? "heading"
  const base: Record<string, unknown> = { by }
  if (by === "heading" || by === "paragraph") {
    if (args.match != null) base.match = args.match
    if (args.matchMode != null) base.matchMode = args.matchMode
    if (args.occurrence != null) base.occurrence = args.occurrence
  }
  if (by === "table") {
    if (args.occurrence != null) base.occurrence = args.occurrence
  }
  if (by === "bookmark") {
    base.name = args.name
  }
  if (args.offsetDirection != null || args.offsetCount != null) {
    base.offset = {
      direction: (args.offsetDirection as string) ?? "after",
      count: (args.offsetCount as number) ?? 1,
    }
  }
  return base as unknown as Locator
}

export function registerMarkdownTool(
  server: McpServer,
  markdown: WordMarkdown,
  textEditor: WordTextEditor,
  positionMap: PositionMap,
  security: SecurityManager,
): void {
  server.registerTool(
    "word_write_markdown",
    {
      description: "Write formatted content using Markdown syntax. Supports: headings (# H1, ## H2, etc.), **bold**, *italic*, `code`, bullet lists (- item), numbered lists (1. item), tables (| pipes |), > blockquotes, and --- horizontal rules. Each block is styled and formatted automatically in Word.",
      inputSchema: {
        text: z.string().min(1).max(100000).describe("Markdown content to convert to formatted Word document"),
      },
    },
    mcpCall(security, "word_write_markdown", async ({ text }) => {
      security.validateTextLength(text)
      const chunks = splitMarkdown(text, 500)
      let totalBlocks = 0, totalChars = 0
      for (let i = 0; i < chunks.length; i++) {
        const result = await markdown.write(chunks[i])
        totalBlocks += result.blocks
        totalChars += result.chars
        if (i < chunks.length - 1) await sleep(100)
      }
      if (totalBlocks === 0) {
        return "Action: No content to write"
      }
      let next = "Next: word_type_text({text:\"...\", mode:\"instant\"}) or word_save()"
      if (totalBlocks <= 3 && totalChars > 0) {
        next = "Next: word_get_info() to verify or word_save_as({path:\"C:\\output.docx\"})"
      }
      return `Action: Markdown written (${totalBlocks} blocks, ${totalChars} chars)\nDetail: Blocks: ${totalBlocks}, Chars: ${totalChars}\n${next}`
    }, { timeoutMs: 300000 }),
  )

  server.registerTool(
    "word_write_markdown_at",
    {
      description: "Write formatted Markdown content at a specific semantic location (heading/paragraph/table/bookmark). Supports same Markdown syntax as word_write_markdown. Moves cursor to the resolved location, then writes.",
      inputSchema: {
        ...locatorFields,
        text: z.string().min(1).max(100000).describe("Markdown content to write at the resolved location"),
      },
    },
    mcpCall(security, "word_write_markdown_at", async (args) => {
      const locator = locatorFromArgs(args)
      const pos = await positionMap.resolve(locator)
      if (!pos.found) {
        return `Action: Write markdown failed\nDetail: ${pos.error}`
      }
      await textEditor.goToParagraph(pos.paragraphIndex)
      const result = await markdown.write(args.text as string)
      return `Action: Markdown written at ${formatPosition(pos)}\nDetail: Blocks: ${result.blocks}, Chars: ${result.chars}`
    }),
  )
}
