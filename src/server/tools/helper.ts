import type { ServerContext, ReadyServerContext } from "../server-context.js"
import { SecurityManager } from "../../security/policy.js"
import { logAudit } from "../../security/audit.js"
import { ComError } from "../../word/com-errors.js"
import { WordMcpError, WordEngineTimeoutError, ServerNotReadyError, toMcpContent, sanitizeErrorMessage } from "../../security/errors.js"
import { isEngineError, isReadOnlyTool } from "./shared.js"
import type { SessionDirector } from "../session-director.js"
import type { Precondition } from "../session-director.js"

export function ensureReady(ctx: ServerContext): asserts ctx is ReadyServerContext {
  if (!ctx.session || !ctx.positionMap || !ctx.director) {
    throw new ServerNotReadyError()
  }
}

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

    ensureReady(context)

    const effectiveTimeout = options?.timeoutMs ?? 30000
    const timeoutPromise = effectiveTimeout === 0
      ? undefined
      : new Promise<never>((_, reject) => {
          setTimeout(() => reject(new WordEngineTimeoutError(toolName)), effectiveTimeout)
        })

    const abortController = effectiveTimeout > 0 ? new AbortController() : undefined
    const signal = abortController?.signal

    try {
      const text = abortController
        ? await Promise.race([handler(args), new Promise<never>((_, reject) => {
            setTimeout(() => {
              abortController.abort()
              reject(new WordEngineTimeoutError(toolName))
            }, effectiveTimeout)
          })])
        : await handler(args)
      director?.circuitBreaker.onSuccess()
      director?.markHealthy()
      logAudit({ tool: toolName, durationMs: Date.now() - start })
      director?.markDirtyIfNeeded(toolName)
      director?.schedulePositionRefresh()
      const contextSuffix = !isReadOnlyTool(toolName) ? (director?.captureContextSuffix() ?? "") : ""
      return { content: [{ type: "text" as const, text: text + contextSuffix + (director?.captureStatusSuffix() ?? "") }] }
    } catch (err) {
      logAudit({ tool: toolName, durationMs: Date.now() - start, error: true })
      const suffix = director?.captureStatusSuffix() ?? ""

      if (err instanceof WordEngineTimeoutError) {
        context.session?.setBusy(true)
      }

      if (isEngineError(err)) {
        director?.circuitBreaker.onFailure()
        try {
          await director?.recoverSession()
        } catch (recoveryErr) {
          console.error("[mcpCall] Session recovery failed:", recoveryErr)
        }
        if (isReadOnlyTool(toolName)) {
          try {
            const retryTimeout = new Promise<never>((_, reject) => {
              setTimeout(() => reject(new WordEngineTimeoutError(toolName)), effectiveTimeout)
            })
            const retryText = await Promise.race([handler(args), retryTimeout])
            director?.circuitBreaker.onSuccess()
            director?.markHealthy()
            logAudit({ tool: toolName, durationMs: Date.now() - start, retry: true })
            return { content: [{ type: "text" as const, text: retryText + suffix }] }
          } catch {
            director?.circuitBreaker.onFailure()
          }
        }
      }

      if (director && !isEngineError(err)) {
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
