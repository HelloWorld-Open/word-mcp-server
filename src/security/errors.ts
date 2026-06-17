export class WordMcpError extends Error {
  constructor(
    message: string,
    public readonly code: string,
    public readonly recoverable: boolean = false,
    public readonly recoveryHint?: string,
  ) {
    super(message)
    this.name = "WordMcpError"
  }
}

export class WordEngineTimeoutError extends WordMcpError {
  constructor(method: string) {
    super(
      `Operation "${method}" timed out`,
      "ENGINE_TIMEOUT",
      true,
      "The operation took too long. Try a smaller scope (fewer rows in table, shorter text). If Word dialog is open, close it and retry.",
    )
    this.name = "WordEngineTimeoutError"
  }
}

export class ServerNotReadyError extends WordMcpError {
  constructor() {
    super(
      "Server components not fully initialized. The session director, position map, or Word session is null.",
      "SERVER_NOT_READY",
      false,
      "Please restart the MCP server. If the problem persists, ensure Microsoft Word is installed and accessible.",
    )
    this.name = "ServerNotReadyError"
  }
}

export class PathSecurityError extends WordMcpError {
  constructor(message: string) {
    super(message, "PATH_SECURITY", false)
    this.name = "PathSecurityError"
  }
}

export function toMcpContent(error: WordMcpError): Array<{ type: "text"; text: string }> {
  return [{ type: "text" as const, text: formatWordMcpError(error) }]
}

function formatWordMcpError(error: WordMcpError): string {
  const lines: string[] = [`[${error.code}] ${error.message}`]
  if (error.recoveryHint) {
    lines.push(`>> Recovery: ${error.recoveryHint}`)
  }
  return lines.join("\n")
}

export function sanitizeErrorMessage(raw: unknown): string {
  if (raw instanceof Error) {
    const msg = raw.message
    if (raw instanceof WordMcpError) {
      return formatWordMcpError(raw)
    }
    const sanitized = msg
      .replace(/\b[A-Za-z]:\\(?:[^\\:*?"<>|\r\n]{1,255}\\){0,10}[^\\:*?"<>|\r\n]{0,255}\b/g, "[path]")
      .replace(/\\{2}[^\\\s]{1,255}(?:\\[^\\\s]{1,255}){0,5}/g, "[network-path]")
    return sanitized.length > 200 ? sanitized.slice(0, 200) + "..." : sanitized
  }
  return String(raw).slice(0, 200)
}
