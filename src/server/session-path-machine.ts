export type SessionPath = "idle" | "streaming" | "editing"

const STREAMING_WATCHDOG_MS = 600_000

export class SessionPathMachine {
  private _currentPath: SessionPath = "idle"
  private _streamingWatchdog: ReturnType<typeof setTimeout> | null = null
  private _onLog: ((level: string, message: string) => void) | null = null

  setOnLog(handler: (level: string, message: string) => void): void {
    this._onLog = handler
  }

  private _log(level: string, message: string): void {
    this._onLog?.(level, message)
  }

  get isStreamingActive(): boolean {
    return this._currentPath === "streaming"
  }

  get currentPath(): SessionPath {
    return this._currentPath
  }

  acquireStreamLock(toolName: string): string | null {
    if (this._currentPath === "editing") {
      this._log("warn", `Stream lock denied for ${toolName} — currently in edit mode`)
      return "[EDIT_MODE] Currently in document editing mode.\n>> Recovery: Close the current document with word_close() first, then use word_stream_start to create a new document."
    }
    if (this._currentPath === "streaming") {
      this._log("warn", `Stream lock denied for ${toolName} — already streaming`)
      return `[STREAMING] Currently in a stream session.\n>> Recovery: End the stream session with word_stream_end first, then use ${toolName}.`
    }
    this._currentPath = "streaming"
    this._startWatchdog()
    this._log("info", `Stream lock acquired by ${toolName} → streaming`)
    return null
  }

  releaseStreamLock(): void {
    this._currentPath = "idle"
    this._stopWatchdog()
    this._log("info", "Stream lock released → idle")
  }

  enterEditMode(): void {
    if (this._currentPath === "idle") {
      this._currentPath = "editing"
      this._log("info", "Edit mode entered → editing")
    }
  }

  exitEditMode(): void {
    if (this._currentPath === "editing") {
      this._currentPath = "idle"
      this._log("info", "Edit mode exited → idle")
    }
  }

  refreshWatchdog(): void {
    if (this._currentPath === "streaming") {
      this._startWatchdog()
      this._log("debug", "Streaming watchdog refreshed")
    }
  }

  private _startWatchdog(): void {
    this._stopWatchdog()
    this._streamingWatchdog = setTimeout(() => {
      this._currentPath = "idle"
      this._log("warn", "Streaming watchdog timed out → idle")
    }, STREAMING_WATCHDOG_MS)
  }

  private _stopWatchdog(): void {
    if (this._streamingWatchdog !== null) {
      clearTimeout(this._streamingWatchdog)
      this._streamingWatchdog = null
    }
  }
}
