import { SecurityManager } from "../../security/policy.js"
import { logAudit } from "../../security/audit.js"
import { WordMcpError, WordEngineTimeoutError, toMcpContent, sanitizeErrorMessage } from "../../security/errors.js"
import type { IWordSession } from "../../word/session.js"
import type { PositionMap } from "../../word/position-map.js"

type Precondition = "DOC" | "NO_DOC"

interface McpCallOptions {
  timeoutMs?: number
  preconditions?: Precondition[]
}

function checkPreconditions(session: IWordSession | null, preconditions: Precondition[] | undefined): string | null {
  if (!session) return null
  const checks = preconditions ?? ["DOC"]
  if (checks.length === 0) return null
  for (const p of checks) {
    if (p === "DOC" && !session.activeDoc) {
      return "[NO_DOCUMENT] 当前没有打开的文档。\n>> Recovery: 请先使用 word_create({title:'...'}) 创建新文档，或 word_document({path:'...'}) 打开已有文档。"
    }
    if (p === "NO_DOC" && session.activeDoc) {
      return "[DOC_ACTIVE] 当前已有文档打开。\n>> Recovery: 请先使用 word_close() 关闭当前文档，再创建新文档。"
    }
  }
  return null
}

const ENGINE_ERROR_KEYWORDS = [
  "automation", "rpc", "server", "call was rejected",
  "0x800", "0x800706ba", "0x80010108",
  "class not registered", "failed due to",
  "object has been disconnected",
]

let _statusSession: IWordSession | null = null
let _positionMap: PositionMap | null = null
let _batchActive = false

export function setStatusSession(session: IWordSession): void {
  _statusSession = session
}

export function setPositionMap(map: PositionMap): void {
  _positionMap = map
}

export function setBatchActive(active: boolean): void {
  _batchActive = active
}

function isEngineError(err: unknown): boolean {
  if (err instanceof WordEngineTimeoutError) return true
  if (err instanceof WordMcpError) return false
  if (err instanceof Error) {
    const msg = err.message.toLowerCase()
    return ENGINE_ERROR_KEYWORDS.some(k => msg.includes(k))
  }
  return false
}

function captureStatusSuffix(): string {
  const s = _statusSession
  if (!s) return ""
  try {
    if (!s.activeDoc) return "\n---\ndoc: none"
    const path = s.activeDocPath
    if (!path) return "\n---\ndoc: untitled"
    const name = path.split(/[\\/]/).pop() ?? "?"
    return `\n---\ndoc: "${name}"`
  } catch {
    return ""
  }
}

function getPositionMap(): PositionMap | null {
  return _positionMap
}

export function mcpCall<T extends Record<string, unknown>>(
  security: SecurityManager,
  toolName: string,
  handler: (args: T) => Promise<string>,
  options?: McpCallOptions,
): (args: T) => Promise<{ content: Array<{ type: "text"; text: string }> }> {
  return async (args: T) => {
    const start = Date.now()
    security.checkRateLimit(toolName)
    const effectiveTimeout = options?.timeoutMs ?? 30000

    const session = _statusSession
    if (session) {
      if (session.isUnhealthy()) {
      } else if (!session.isAlive()) {
        try { await session.recover() } catch { /* best effort */ }
      }
    }

    const preErr = checkPreconditions(session, options?.preconditions)
    if (preErr) {
      return { content: [{ type: "text" as const, text: preErr + captureStatusSuffix() }] }
    }

    const timeoutPromise = new Promise<never>((_, reject) => {
      setTimeout(() => reject(new WordEngineTimeoutError(toolName)), effectiveTimeout)
    })

    try {
      const text = await Promise.race([handler(args), timeoutPromise])
      if (session) session.markHealthy()
      logAudit({ tool: toolName, durationMs: Date.now() - start })
      getPositionMap()?.markDirty()
      return { content: [{ type: "text" as const, text: text + captureStatusSuffix() }] }
    } catch (err) {
      logAudit({ tool: toolName, durationMs: Date.now() - start, error: true })
      const suffix = captureStatusSuffix()

      if (session && isEngineError(err)) {
        try {
          await session.recover()
          const retryText = await Promise.race([handler(args), timeoutPromise])
          session.markHealthy()
          logAudit({ tool: toolName, durationMs: Date.now() - start, retry: true })
          getPositionMap()?.markDirty()
          return { content: [{ type: "text" as const, text: retryText + captureStatusSuffix() }] }
        } catch {
        }
      }

      if (session && !isEngineError(err)) {
        session.markHealthy()
      }

      if (err instanceof WordMcpError) {
        const content = toMcpContent(err)
        content[0].text += suffix
        return { content }
      }
      return { content: [{ type: "text" as const, text: sanitizeErrorMessage(err) + suffix }] }
    }
  }
}
