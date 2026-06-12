import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js"
import { mcpCall, setBatchActive } from "./helper.js"
import type { IWordSession } from "../../word/session.js"
import { SecurityManager } from "../../security/policy.js"

export function registerBatchTools(
  server: McpServer,
  session: IWordSession,
  security: SecurityManager,
): void {
  server.registerTool(
    "word_batch_start",
    {
      description: "Start a batch operation mode — disables Word screen updating to accelerate multiple consecutive operations 3-10x. WHEN: about to perform 3+ operations in sequence (typing, formatting, tables). NOT: only doing 1-2 quick edits? skip this, the gain is negligible.",
    },
    mcpCall(security, "word_batch_start", async () => {
      session.setScreenUpdating(false)
      setBatchActive(true)
      return "Action: Batch mode started\nDetail: Screen updating disabled — operations will run 3-10x faster\nNext: Perform your edits, then call word_batch_end() to restore live updates"
    }),
  )

  server.registerTool(
    "word_batch_end",
    {
      description: "End batch operation mode — re-enables Word screen updating and refreshes the window. WHEN: after word_batch_start, once all operations are done. NOT: batch not started? this has no effect.",
    },
    mcpCall(security, "word_batch_end", async () => {
      session.setScreenUpdating(true)
      setBatchActive(false)
      try {
        const doc = session.activeDoc as Record<string, unknown> | null
        if (doc) {
          const w = doc.ActiveWindow as Record<string, unknown> | undefined
          if (w) {
            ;(w.Refresh as (() => void) | undefined)?.()
          }
        }
      } catch { /* ignore */ }
      return "Action: Batch mode ended\nDetail: Screen updating restored, window refreshed"
    }),
  )

}