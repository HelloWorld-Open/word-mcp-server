interface AuditEntry {
  tool: string
  durationMs: number
  error?: boolean
  retry?: boolean
  args?: Record<string, unknown>
}

const SENSITIVE_KEYS = new Set(["password", "token", "apiKey", "secret", "api_key", "api-key"])

function redactArg(k: string, v: unknown): unknown {
  if (SENSITIVE_KEYS.has(k)) return "***REDACTED***"
  if (typeof v === "string" && v.length > 100) return v.slice(0, 100) + "..."
  return v
}

export function logAudit(entry: AuditEntry): void {
  const timestamp = new Date().toISOString()
  const status = entry.error ? "REJECTED" : "OK"
  let extra = ""
  if (entry.args) {
    const s: Record<string, unknown> = {}
    for (const [k, v] of Object.entries(entry.args)) {
      s[k] = redactArg(k, v)
    }
    extra = ` args=${JSON.stringify(s)}`
  }
  console.error(`[audit] ${timestamp} ${status} ${entry.tool} ${entry.durationMs}ms${extra}`)
}
