import { z } from "zod"
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js"
import { WordTextEditor } from "../../word/word-text-editor.js"
import { WordTableEditor } from "../../word/word-table-editor.js"
import { WordMarkdown } from "../../word/word-markdown.js"
import { PositionMap } from "../../word/position-map.js"
import { locatorFromArgs } from "./semantic.js"
import { SecurityManager } from "../../security/policy.js"
import type { ServerContext } from "../server-context.js"
import { mcpCall } from "./helper.js"

const baseLocatorFields = {
  by: z.enum(["heading", "paragraph", "table", "bookmark", "cursor"]).optional().describe("Target type (default: heading)"),
  match: z.string().max(5000).optional().describe("Text to match (for heading/paragraph)"),
  matchMode: z.enum(["exact", "contains", "startsWith", "regex"]).optional().describe("Matching mode (default: exact)"),
  occurrence: z.number().int().min(1).max(1000).optional().describe("Which occurrence to target (1-based, default: 1)"),
  offsetDirection: z.enum(["before", "after"]).optional().describe("Offset direction from the matched element"),
  offsetCount: z.number().int().min(1).max(1000).optional().describe("Number of paragraphs to offset (default: 1)"),
  name: z.string().min(1).max(255).optional().describe("Bookmark name (required when by='bookmark')"),
  level: z.number().int().min(1).max(9).optional().describe("Heading level filter (1-9). Applied BEFORE match and occurrence."),
}

const insertOp = z.object({
  ...baseLocatorFields,
  action: z.literal("insert").describe("Insert markdown content at the located position"),
  text: z.string().min(1).max(100000).describe("Markdown content to insert"),
})

const batchOp = z.discriminatedUnion("action", [insertOp])

export function registerBatchTools(
  server: McpServer,
  context: ServerContext,
  textEditor: WordTextEditor,
  tableEditor: WordTableEditor,
  markdown: WordMarkdown,
  positionMap: PositionMap,
  security: SecurityManager,
): void {
  server.registerTool(
    "word_batch_ops",
    {
      description: "批量执行多个文档写入操作。单次定位扫描 + 连续写入，比逐条调用 word_insert_at 快 5-10 倍。操作间自动 yield 并校验 DOM 一致性。支持按标题/段落/表格/书签定位后插入 Markdown 内容。",
      inputSchema: {
        ops: z.array(batchOp).min(1).max(50).describe("操作列表，按文档顺序排列"),
      },
    },
    mcpCall(security, context, "word_batch_ops", async (args) => {
      const ops = args.ops as Array<Record<string, unknown>>
      const results: string[] = []

      await positionMap.ensureFresh()
      let expectedCount = positionMap.cachedParaCount

      for (let i = 0; i < ops.length; i++) {
        const op = ops[i]

        if (i > 0) {
          await new Promise(resolve => setImmediate(resolve))
          const match = await positionMap.paraCountMatches(expectedCount)
          if (!match) {
            results.push(`[${i + 1}] 锚点校验: DOM 段落数已变化，重新定位扫描`)
            await positionMap.ensureFresh()
          }
        }

        if (op.by && op.by !== "cursor") {
          const locator = locatorFromArgs(op)
          const pos = await positionMap.resolve(locator, true)
          if (!pos.found) {
            results.push(`[${i + 1}] 跳过: ${pos.error}`)
            continue
          }
          await textEditor.goToParagraph(pos.paragraphIndex)
        }

        const text = op.text as string
        const preview = text.slice(0, 60) + (text.length > 60 ? "..." : "")
        const result = await markdown.insertAtCursor(text)
        expectedCount = await positionMap.fetchActualParaCount()
        results.push(`[${i + 1}] 完成: "${preview}" (${result.blocks} blocks, ${result.chars} chars)`)
      }

      positionMap.markDirty()
      return results.join("\n")
    }),
  )
}
