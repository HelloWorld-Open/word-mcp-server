interface AuditEntry {
  tool: string
  durationMs: number
  error?: boolean
  retry?: boolean
  args?: Record<string, unknown>
  traceId?: string
  paramCount?: number
  recoveryTriggered?: boolean
}

export function logAudit(entry: AuditEntry): void {
  const { traceId } = entry
  const status = entry.error ? "REJECTED" : "OK"
  let extra = ""
  if (entry.args) {
    extra = ` args=${JSON.stringify(entry.args, redactReplacer)}`
  }
  const ts = new Date().toISOString()
  const parts = [`[audit]`, ts, status, entry.tool, `${entry.durationMs}ms`]
  if (traceId) parts.push(`traceId=${traceId}`)
  if (entry.paramCount != null) parts.push(`params=${entry.paramCount}`)
  if (entry.recoveryTriggered) parts.push(`recovery=1`)
  console.error(parts.join(" ") + extra)
}

const SENSITIVE_KEYS = new Set(["password", "token", "apiKey", "secret", "api_key", "api-key"])

function redactReplacer(_key: string, val: unknown): unknown {
  if (SENSITIVE_KEYS.has(_key)) return "***REDACTED***"
  if (typeof val === "string" && val.length > 100) return val.slice(0, 100) + "..."
  return val
}
