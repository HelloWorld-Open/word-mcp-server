import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js"
import type { ServerContext } from "../server-context.js"
import { getSessionLogs, type LogEntry } from "../../logger.js"

export function registerDiagnosticResource(
  server: McpServer,
  _context: ServerContext,
): void {
  server.registerResource(
    "session-diagnostics",
    "word-mcp://diagnostics/logs",
    {
      description: "Session-level diagnostic log buffer. Contains the last 500 log entries from all tool calls in the current session. Supports query parameters: count (1-500, default 50) and level (error/warn/info/debug/trace, optional filter). Use for post-hoc investigation of errors and operation history.",
      mimeType: "application/json",
    },
    async (uri: URL) => {
      const countParam = uri.searchParams.get("count")
      const levelParam = uri.searchParams.get("level")
      const count = countParam ? Math.min(Math.max(1, parseInt(countParam, 10) || 50), 500) : 50
      const level = levelParam && ["fatal", "error", "warn", "info", "debug", "trace"].includes(levelParam) ? levelParam : undefined

      const logs = getSessionLogs(count, level)

      return {
        contents: [{
          uri: uri.href,
          mimeType: "application/json",
          text: JSON.stringify({
            total: logs.length,
            remaining: 0,
            logs: logs.map((l: LogEntry) => ({
              level: l.level,
              msg: l.msg,
              time: new Date(l.time).toISOString(),
            })),
          }),
        }],
      }
    },
  )
}
