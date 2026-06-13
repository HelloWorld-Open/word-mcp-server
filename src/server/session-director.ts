import type { IWordSession } from "../word/session.js"
import type { PositionMap } from "../word/position-map.js"
import { WordMcpError, WordEngineTimeoutError } from "../security/errors.js"

type Precondition = "DOC" | "NO_DOC"

type SessionPath = "idle" | "streaming" | "editing"

const ENGINE_ERROR_KEYWORDS = [
  "automation", "rpc", "server", "call was rejected",
  "0x800", "0x800706ba", "0x80010108",
  "class not registered", "failed due to",
  "object has been disconnected",
]

const READ_ONLY_TOOLS = new Set([
  "word_get_text", "word_get_paragraph", "word_get_structure", "word_get_info",
  "word_get_status", "word_get_table_data", "word_get_comments", "word_get_bookmarks",
  "word_get_lists", "word_get_sections", "word_get_cursor_info", "word_locate",
  "word_list_styles",
])

const STREAM_BLOCKED_TOOLS = new Set([
  "word_document", "word_open", "word_close", "word_quit",
])

const EDIT_BLOCKED_TOOLS = new Set([
  "word_stream_start",
])

const STREAMING_WATCHDOG_MS = 600_000

export class SessionDirector {
  private _currentPath: SessionPath = "idle"
  private _streamingWatchdog: ReturnType<typeof setTimeout> | null = null
  private _session: IWordSession | null
  private _positionMap: PositionMap | null

  constructor(session: IWordSession | null, positionMap: PositionMap | null) {
    this._session = session
    this._positionMap = positionMap
  }

  get isStreamingActive(): boolean {
    return this._currentPath === "streaming"
  }

  get currentPath(): SessionPath {
    return this._currentPath
  }

  get session(): IWordSession | null {
    return this._session
  }

  get positionMap(): PositionMap | null {
    return this._positionMap
  }

  acquireStreamLock(toolName: string): string | null {
    if (this._currentPath === "editing") {
      return "[编辑模式] 当前处于文档编辑模式。\n>> Recovery: 请先使用 word_close() 关闭当前文档，再使用 word_stream_start 创建新文档。"
    }
    if (this._currentPath === "streaming") {
      return `[STREAMING] 当前处于流式会话中。\n>> Recovery: 请先用 word_stream_end 结束流式会话，再使用 ${toolName}。`
    }
    this._currentPath = "streaming"
    this._startWatchdog()
    return null
  }

  releaseStreamLock(): void {
    this._currentPath = "idle"
    this._stopWatchdog()
  }

  enterEditMode(): void {
    if (this._currentPath === "idle") {
      this._currentPath = "editing"
    }
  }

  exitEditMode(): void {
    if (this._currentPath === "editing") {
      this._currentPath = "idle"
    }
  }

  refreshWatchdog(): void {
    if (this._currentPath === "streaming") this._startWatchdog()
  }

  private _startWatchdog(): void {
    this._stopWatchdog()
    this._streamingWatchdog = setTimeout(() => {
      this._currentPath = "idle"
    }, STREAMING_WATCHDOG_MS)
  }

  private _stopWatchdog(): void {
    if (this._streamingWatchdog !== null) {
      clearTimeout(this._streamingWatchdog)
      this._streamingWatchdog = null
    }
  }

  async precheck(
    toolName: string,
    precondition?: Precondition | Precondition[],
  ): Promise<{ ok: true } | { ok: false; error: string }> {
    if (this._currentPath === "streaming" && STREAM_BLOCKED_TOOLS.has(toolName)) {
      return {
        ok: false,
        error: `[STREAMING] 当前处于流式会话中。\n>> Recovery: 请先用 word_stream_end 结束流式会话，再使用 ${toolName}。`,
      }
    }
    if (this._currentPath === "editing" && EDIT_BLOCKED_TOOLS.has(toolName)) {
      return {
        ok: false,
        error: "[编辑模式] 当前处于文档编辑模式。\n>> Recovery: 请先使用 word_close() 关闭当前文档，再使用 word_stream_start 创建新文档。",
      }
    }

    const session = this._session
    if (session) {
      if (session.isUnhealthy()) {
      } else if (!session.isAlive()) {
        try { await session.recover() } catch { /* best effort */ }
      }
    }

    const preErr = this._checkPreconditions(precondition)
    if (preErr) return { ok: false, error: preErr }

    return { ok: true }
  }

  private _checkPreconditions(precondition?: Precondition | Precondition[]): string | null {
    const session = this._session
    if (!session) return null

    let p: Precondition | undefined
    if (Array.isArray(precondition)) {
      if (precondition.length === 0) return null
      p = precondition[0]
    } else {
      p = precondition ?? "DOC"
    }
    if (p === "DOC" && !session.activeDoc) {
      try {
        const app = session.application as Record<string, unknown>
        const doc = app.ActiveDocument as Record<string, unknown> | undefined
        if (doc) {
          session.setActiveDoc(doc)
          try { session.setActiveDocPath(doc.FullName as string) } catch { /* ignore */ }
        }
      } catch { /* ignore */ }
      if (!session.activeDoc) {
        return "[NO_DOCUMENT] 当前没有打开的文档。\n>> Recovery: 请先使用 word_stream_start({title:'...'}) 创建新文档，或 word_document({path:'...'}) 打开已有文档。"
      }
    }
    if (p === "NO_DOC" && session.activeDoc) {
      return "[DOC_ACTIVE] 当前已有文档打开。\n>> Recovery: 请先使用 word_close() 关闭当前文档，再创建新文档。"
    }
    return null
  }

  captureStatusSuffix(): string {
    const session = this._session
    if (!session) return ""
    try {
      const pathLabel = this._currentPath === "streaming" ? " [流式会话活跃]" : this._currentPath === "editing" ? " [编辑模式]" : ""
      if (!session.activeDoc) return `\n---\ndoc: none${pathLabel}`
      const path = session.activeDocPath
      if (!path) return `\n---\ndoc: untitled${pathLabel}`
      const name = path.split(/[\\/]/).pop() ?? "?"
      return `\n---\ndoc: "${name}"${pathLabel}`
    } catch {
      return ""
    }
  }

  markHealthy(): void {
    this._session?.markHealthy()
  }

  markDirtyIfNeeded(toolName: string): void {
    if (!SessionDirector.isReadOnlyTool(toolName)) this._positionMap?.markDirty()
  }

  async recoverSession(): Promise<void> {
    await this._session?.recover()
  }

  static isEngineError(err: unknown): boolean {
    if (err instanceof WordEngineTimeoutError) return true
    if (err instanceof WordMcpError) return false
    if (err instanceof Error) {
      const msg = err.message.toLowerCase()
      return ENGINE_ERROR_KEYWORDS.some(k => msg.includes(k))
    }
    return false
  }

  static isReadOnlyTool(toolName: string): boolean {
    return READ_ONLY_TOOLS.has(toolName)
  }
}
