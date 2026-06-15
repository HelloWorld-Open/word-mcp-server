import type { ServerContext } from "../server-context.js"
import { SecurityManager } from "../../security/policy.js"
import { logAudit } from "../../security/audit.js"
import { ComError } from "../../word/com-errors.js"
import { WordMcpError, WordEngineTimeoutError, toMcpContent, sanitizeErrorMessage } from "../../security/errors.js"
import { SessionDirector } from "../session-director.js"

type Precondition = "DOC" | "NO_DOC"

interface McpCallOptions {
  timeoutMs?: number
  preconditions?: Precondition[]
}

export function mcpCall<T extends Record<string, unknown>>(
  security: SecurityManager,
  context: ServerContext,
  toolName: string,
  handler: (args: T) => Promise<string>,
  options?: McpCallOptions,
): (args: T) => Promise<{ content: Array<{ type: "text"; text: string }> }> {
  return async (args: T) => {
    const start = Date.now()
    security.checkRateLimit(toolName)

    const director = context.director
    if (director) {
      const precheck = await director.precheck(toolName, options?.preconditions)
      if (!precheck.ok) {
        return {
          content: [{ type: "text" as const, text: precheck.error + director.captureStatusSuffix() }],
        }
      }
    }

    const effectiveTimeout = options?.timeoutMs ?? 30000
    const timeoutPromise = new Promise<never>((_, reject) => {
      setTimeout(() => reject(new WordEngineTimeoutError(toolName)), effectiveTimeout)
    })

    try {
      const text = await Promise.race([handler(args), timeoutPromise])
      director?.markHealthy()
      logAudit({ tool: toolName, durationMs: Date.now() - start })
      director?.markDirtyIfNeeded(toolName)
      director?.schedulePositionRefresh()
      return { content: [{ type: "text" as const, text: text + (director?.captureStatusSuffix() ?? "") }] }
    } catch (err) {
      logAudit({ tool: toolName, durationMs: Date.now() - start, error: true })
      const suffix = director?.captureStatusSuffix() ?? ""

      if (SessionDirector.isEngineError(err)) {
        try {
          await director?.recoverSession()
        } catch {
        }
        if (SessionDirector.isReadOnlyTool(toolName)) {
          try {
            const retryTimeout = new Promise<never>((_, reject) => {
              setTimeout(() => reject(new WordEngineTimeoutError(toolName)), effectiveTimeout)
            })
            const retryText = await Promise.race([handler(args), retryTimeout])
            director?.markHealthy()
            logAudit({ tool: toolName, durationMs: Date.now() - start, retry: true })
            return { content: [{ type: "text" as const, text: retryText + suffix }] }
          } catch {
          }
        }
      }

      if (director && !SessionDirector.isEngineError(err)) {
        director.markHealthy()
      }

      if (err instanceof WordMcpError) {
        const content = toMcpContent(err)
        content[0].text += suffix
        return { content }
      }
      if (err instanceof ComError) {
        const tag = err.recoverable ? "COM_TRANSIENT" : "COM_FATAL"
        const hint = err.recoverable
          ? "\n>> Recovery: The operation will be retried automatically. If the problem persists, close Word and try again."
          : "\n>> Recovery: Word connection lost. Please restart the document session."
        return { content: [{ type: "text" as const, text: `[${tag}] ${err.message}${hint}${suffix}` }] }
      }
      return { content: [{ type: "text" as const, text: sanitizeErrorMessage(err) + suffix }] }
    }
  }
}
