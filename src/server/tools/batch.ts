import { z } from "zod"
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js"
import { WordCursor } from "../../word/word-cursor.js"
import { WordContentWriter } from "../../word/word-content-writer.js"
import { PositionMap } from "../../word/position-map.js"
import { locatorFromArgs } from "./semantic.js"
import { SecurityManager } from "../../security/policy.js"
import type { ServerContext } from "../server-context.js"
import { createRegTool, locatorFields } from "./shared.js"

const insertOp = z.object({
  ...locatorFields,
  action: z.literal("insert").describe("Insert markdown content at the located position"),
  text: z.string().min(1).max(100000).describe("Markdown content to insert"),
})

const batchOp = z.discriminatedUnion("action", [insertOp])

export function registerBatchTools(
  server: McpServer,
  context: ServerContext,
  cursor: WordCursor,
  contentWriter: WordContentWriter,
  positionMap: PositionMap,
  security: SecurityManager,
): void {
  const regTool = createRegTool(server, security, context)
  regTool("word_batch_ops",
    {
      description: "WHEN: need to perform multiple document insert operations at different locations in one pass for 5-10x speedup. WHAT: batch-executes markdown insert operations (by heading/paragraph/table/bookmark/cursor) with automatic DOM consistency checking between operations. CONSTRAINT: max 50 operations per call. For a single operation, use word_insert_at which has simpler error reporting.",
      inputSchema: {
        ops: z.array(batchOp).min(1).max(50).describe("操作列表，按文档顺序排列"),
      },
    },
    async (args) => {
      const ops = args.ops as Array<Record<string, unknown>>
      const results: string[] = []

      // Fast path: when all ops are cursor-mode, skip PositionMap entirely —
      // no heading/paragraph/table/bookmark resolution needed, saving O(N) COM
      // calls per operation (table iteration + heading search per refresh).
      const allCursorMode = ops.every(op => !op.by || op.by === "cursor")

      if (!allCursorMode) {
        await positionMap.ensureFresh()
      }
      let expectedCount = allCursorMode ? 0 : positionMap.cachedParaCount

      for (let i = 0; i < ops.length; i++) {
        const op = ops[i]

        if (i > 0 && !allCursorMode) {
          await new Promise(resolve => setImmediate(resolve))
          const match = await positionMap.paraCountMatches(expectedCount)
          if (!match) {
            results.push(`[${i + 1}] Anchor check: DOM paragraph count changed, re-scanning`)
            await positionMap.ensureFresh()
          }
        }

        if (op.by && op.by !== "cursor") {
          const locator = locatorFromArgs(op)
          const pos = await positionMap.resolve(locator, true)
          if (!pos.found) {
            results.push(`[${i + 1}] Skipped: ${pos.error}`)
            continue
          }
          await cursor.goToParagraph(pos.paragraphIndex)
        }

        const text = op.text as string
        const preview = text.slice(0, 60) + (text.length > 60 ? "..." : "")
        const result = await contentWriter.insertAtCursor(text)
        if (!allCursorMode) {
          expectedCount = await positionMap.fetchActualParaCount()
        }
        results.push(`[${i + 1}] Done: "${preview}" (${result.blocks} blocks, ${result.chars} chars)`)
      }

      if (!allCursorMode) {
        positionMap.markDirty()
      }
      return results.join("\n")
    },
  )
}
