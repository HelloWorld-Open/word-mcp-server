import type { ServerContext, ReadyServerContext } from "../server-context.js"
import { SecurityManager } from "../../security/policy.js"
import { logAudit } from "../../security/audit.js"
import { ComError } from "../../word/com-errors.js"
import { WordMcpError, WordEngineTimeoutError, ServerNotReadyError, toMcpContent, sanitizeErrorMessage } from "../../security/errors.js"
import { isEngineError, isReadOnlyTool } from "./shared.js"
import type { SessionDirector } from "../session-director.js"
import type { Precondition } from "../session-director.js"
import { generateTraceId, runWithTraceId, getLogBuffer } from "../../logger.js"
import { isToolResponse, type ToolResponse, type SystemMeta, type RecoveryAction, type DiagnosticContext } from "../tool-result.js"

const MAX_RECOVERY_ATTEMPTS = 1

const RECOVERY_MAP: Record<string, RecoveryAction> = {
  "[NO_DOCUMENT]": {
    tool: "word_stream_start",
    args: { title: "新文档" },
    hint: "当前无活动文档。请创建新文档或打开已有文件。",
    maxAttempts: 1,
  },
  "[STREAMING]": {
    tool: "word_stream_end",
    args: {},
    hint: "当前有流式会话活跃。请先结束流式会话再执行此操作。",
    maxAttempts: 1,
  },
  "[EDIT_MODE]": {
    tool: "word_document",
    args: {},
    hint: "当前在编辑模式。请关闭当前文档或切换到其他文档。",
    maxAttempts: 1,
  },
}

export function ensureReady(ctx: ServerContext): asserts ctx is ReadyServerContext {
  if (!ctx.session || !ctx.positionMap || !ctx.director) {
    throw new ServerNotReadyError()
  }
}

interface McpCallOptions {
  timeoutMs?: number
  preconditions?: Precondition[]
}

type ContentBlock = { type: "text"; text: string; annotations?: { audience: ("user" | "assistant")[] } }

function buildErrorContent(text: string, diagnostic?: DiagnosticContext | null): ContentBlock[] {
  const blocks: ContentBlock[] = [{ type: "text", text }]
  if (diagnostic) {
    blocks.push({
      type: "text",
      text: JSON.stringify(diagnostic, null, 2),
      annotations: { audience: ["assistant"] },
    })
  }
  return blocks
}

function buildSuccessContent(text: string, meta: SystemMeta | null, toolData?: Record<string, unknown>): ContentBlock[] {
  const blocks: ContentBlock[] = [{ type: "text", text }]
  if (meta) {
    const merged = toolData ? { ...meta, ...toolData } : meta
    blocks.push({
      type: "text",
      text: JSON.stringify(merged, null, 2),
      annotations: { audience: ["assistant"] },
    })
  }
  return blocks
}

function assembleDiagnostic(
  ctx: { traceId: string; tool: string; paramCount: number; durationMs: number },
  error?: { code: string; type: "precheck" | "WordMcpError" | "ComError" | "engine_timeout" | "unknown"; message: string; recoverable: boolean },
  recovery?: RecoveryAction & { attempt: number },
): DiagnosticContext | null {
  const logs = getLogBuffer()
  if (!logs.length && !error && !recovery) return null

  return {
    request: { ...ctx },
    ...(error ? { error } : {}),
    ...(recovery ? { recovery } : {}),
    logs,
  }
}

function resolveRecovery(err: unknown, attempt: number): (RecoveryAction & { attempt: number }) | undefined {
  if (attempt >= MAX_RECOVERY_ATTEMPTS) return undefined

  let action: RecoveryAction | undefined
  if (err instanceof WordMcpError) {
    action = RECOVERY_MAP[`[${err.code}]`]
  } else if (err instanceof ComError && !err.recoverable) {
    action = { tool: "word_get_status", args: {}, hint: "Word 连接已断开。请检查 Word 是否仍在运行，或重启 MCP 服务器。", maxAttempts: 1 }
  }

  return action ? { ...action, attempt } : undefined
}

function errorDiagnostic(err: unknown, errMeta: { traceId: string; tool: string; paramCount: number; durationMs: number }, attempt: number): DiagnosticContext | null {
  if (err instanceof WordMcpError) {
    return assembleDiagnostic(
      errMeta,
      { code: err.code, type: "WordMcpError", message: err.message, recoverable: err.recoverable },
      resolveRecovery(err, attempt),
    )
  }
  if (err instanceof ComError) {
    return assembleDiagnostic(
      errMeta,
      { code: err.recoverable ? "COM_TRANSIENT" : "COM_FATAL", type: "ComError", message: err.message, recoverable: err.recoverable },
      resolveRecovery(err, attempt),
    )
  }
  if (err instanceof WordEngineTimeoutError) {
    return assembleDiagnostic(
      errMeta,
      { code: "ENGINE_TIMEOUT", type: "engine_timeout", message: err.message, recoverable: false },
    )
  }
  return assembleDiagnostic(errMeta)
}

export function mcpCall<T extends Record<string, unknown>>(
  security: SecurityManager,
  context: ServerContext,
  toolName: string,
  handler: (args: T) => Promise<string | ToolResponse>,
  options?: McpCallOptions,
): (args: T) => Promise<{ content: ContentBlock[] }> {
  const recoveryCounters = new Map<string, number>()

  return async (args: T) => {
    const start = Date.now()
    const traceId = generateTraceId()
    context.traceId = traceId

    const requestLogger = context.logger?.child({ traceId, tool: toolName })

    requestLogger?.info({ args }, "tool call")
    security.checkRateLimit(toolName)
    const paramCount = Object.keys(args).length

    const durationMs = () => Date.now() - start
    const errMeta = () => ({ traceId, tool: toolName, paramCount, durationMs: durationMs() })

    const director = context.director
    if (director) {
      const precheck = await director.precheck(toolName, options?.preconditions)
      if (!precheck.ok) {
        requestLogger?.warn(`precheck rejected: ${precheck.error}`)
        context.traceId = undefined
        const attempts = recoveryCounters.get(traceId) ?? 0
        recoveryCounters.set(traceId, attempts + 1)
        const diagnostic = assembleDiagnostic(
          errMeta(),
          { code: precheck.error.match(/^\[(\w+)\]/)?.[1] ?? "PRECHECK_FAILED", type: "precheck", message: precheck.error, recoverable: true },
          resolveRecovery(new WordMcpError(precheck.error, "PRECHECK_FAILED", true, ""), attempts),
        )
        return { content: buildErrorContent(precheck.error, diagnostic) }
      }
    }

    ensureReady(context)

    const effectiveTimeout = options?.timeoutMs ?? 30000

    const abortController = effectiveTimeout > 0 ? new AbortController() : undefined
    const signal = abortController?.signal

    try {
      const raw = abortController
        ? await Promise.race([runWithTraceId(traceId, () => handler(args)), new Promise<never>((_, reject) => {
            setTimeout(() => {
              abortController.abort()
              reject(new WordEngineTimeoutError(toolName))
            }, effectiveTimeout)
          })])
        : await runWithTraceId(traceId, () => handler(args))

      const toolData = isToolResponse(raw) ? raw.data : undefined
      const text = isToolResponse(raw) ? raw.text : raw
      director?.circuitBreaker.onSuccess()
      director?.markHealthy()
      logAudit({ tool: toolName, durationMs: durationMs(), traceId, paramCount })
      director?.markDirtyIfNeeded(toolName)
      director?.schedulePositionRefresh()

      const meta = director?.buildToolMeta() ?? null
      requestLogger?.info({ durationMs: durationMs() }, "tool ok")
      context.traceId = undefined
      return { content: buildSuccessContent(text, meta, toolData) }
    } catch (err) {
      logAudit({ tool: toolName, durationMs: durationMs(), error: true, traceId, paramCount })

      if (err instanceof WordEngineTimeoutError) {
        context.session?.setBusy(true)
      }

      if (isEngineError(err)) {
        director?.circuitBreaker.onFailure()
        try {
          await director?.recoverSession()
        } catch (recoveryErr) {
          requestLogger?.error({ err: recoveryErr }, "Session recovery failed")
        }
        if (isReadOnlyTool(toolName)) {
          try {
            const retryTimeout = new Promise<never>((_, reject) => {
              setTimeout(() => reject(new WordEngineTimeoutError(toolName)), effectiveTimeout)
            })
            const retryRaw = await Promise.race([runWithTraceId(traceId, () => handler(args)), retryTimeout])
            const retryToolData = isToolResponse(retryRaw) ? retryRaw.data : undefined
            const retryText = isToolResponse(retryRaw) ? retryRaw.text : retryRaw
            director?.circuitBreaker.onSuccess()
            director?.markHealthy()
            logAudit({ tool: toolName, durationMs: durationMs(), retry: true, traceId })
            context.traceId = undefined
            const meta = director?.buildToolMeta() ?? null
            return { content: buildSuccessContent(retryText, meta, retryToolData) }
          } catch {
            director?.circuitBreaker.onFailure()
          }
        }
      }

      if (director && !isEngineError(err)) {
        director.markHealthy()
      }

      requestLogger?.error({ err, durationMs: durationMs() }, "tool error")
      context.traceId = undefined

      const attempts = recoveryCounters.get(traceId) ?? 0
      recoveryCounters.set(traceId, attempts + 1)

      if (err instanceof WordMcpError) {
        const diagnostic = errorDiagnostic(err, errMeta(), attempts)
        return { content: buildErrorContent(toMcpContent(err)[0].text, diagnostic) }
      }
      if (err instanceof ComError) {
        const tag = err.recoverable ? "COM_TRANSIENT" : "COM_FATAL"
        const hint = err.recoverable
          ? "\n>> Recovery: The operation will be retried automatically. If the problem persists, close Word and try again."
          : "\n>> Recovery: Word connection lost. Please restart the document session."
        const text = `[${tag}] ${err.message}${hint}`
        const diagnostic = errorDiagnostic(err, errMeta(), attempts)
        return { content: buildErrorContent(text, diagnostic) }
      }
      return { content: buildErrorContent(sanitizeErrorMessage(err), errorDiagnostic(err, errMeta(), attempts)) }
    }
  }
}
