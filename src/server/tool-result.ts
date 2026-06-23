import type { LogEntry } from "../logger.js"

export interface RecoveryAction {
  tool: string
  args: Record<string, unknown>
  hint: string
  maxAttempts: number
}

export interface DiagnosticContext {
  request: {
    traceId: string
    tool: string
    paramCount: number
    durationMs: number
  }
  error?: {
    code: string
    type: "precheck" | "WordMcpError" | "ComError" | "engine_timeout" | "unknown"
    message: string
    recoverable: boolean
  }
  recovery?: RecoveryAction & { attempt: number }
  logs: LogEntry[]
}

export interface SystemMeta {
  struct?: { v: number; p: number; h: number; t: number }
  doc: { name: string | null; state: "none" | "untitled" | "named" }
  diagnostic?: DiagnosticContext
}

export interface ToolResponse {
  text: string
  data?: Record<string, unknown>
}

export function isToolResponse(v: unknown): v is ToolResponse {
  return typeof v === "object" && v !== null && "text" in v && typeof (v as ToolResponse).text === "string"
}
