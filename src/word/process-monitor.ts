import { exec } from "node:child_process"

export interface ProcessCheckFn {
  (callback: (err: Error | null, found: boolean) => void): void
}

interface ProcessMonitorOptions {
  checkFn?: ProcessCheckFn
}

export class ProcessMonitor {
  private _alive = false
  private _lastCheck = 0
  private _intervalMs: number
  private _timer: ReturnType<typeof setInterval> | null = null
  private _pendingCheck: boolean = false
  private _destroyed = false
  private _pid: string | undefined
  private _options?: ProcessMonitorOptions
  private _onStatusChange: ((alive: boolean) => void) | null = null

  constructor(intervalMs = 30000, options?: ProcessMonitorOptions) {
    this._intervalMs = intervalMs
    this._options = options
  }

  start(): void {
    if (this._timer || this._destroyed) return
    this._alive = true
    this._check()
    this._timer = setInterval(() => this._check(), this._intervalMs)
  }

  stop(): void {
    if (this._timer) {
      clearInterval(this._timer)
      this._timer = null
    }
    this._alive = false
  }

  destroy(): void {
    this.stop()
    this._destroyed = true
  }

  isAlive(): boolean {
    return this._alive
  }

  lastCheckTime(): number {
    return this._lastCheck
  }

  getPid(): string | undefined {
    return this._pid
  }

  onAliveChange(cb: (alive: boolean) => void): void {
    this._onStatusChange = cb
  }

  async waitForExit(timeoutMs = 30000): Promise<boolean> {
    const deadline = Date.now() + timeoutMs
    while (Date.now() < deadline) {
      if (!this._alive) return true
      await new Promise<void>(r => setTimeout(r, 200))
    }
    return !this._alive
  }

  /** After a recovery-initiated restart, optimistically mark the process as alive.
   *  The next periodic tasklist poll will confirm or correct this state. */
  markAlive(): void {
    this._alive = true
  }

  private _check(): void {
    if (this._pendingCheck || this._destroyed) return
    this._pendingCheck = true
    if (this._options?.checkFn) {
      this._options.checkFn((err, found) => {
        this._pendingCheck = false
        this._lastCheck = Date.now()
        if (err) return
        this._alive = found
        if (!found) this._pid = undefined
        if (this._onStatusChange) this._onStatusChange(found)
      })
      return
    }
    exec(
      'tasklist /FO CSV /NH',
      { timeout: 5000, encoding: "utf-8" },
      (err, stdout) => {
        this._pendingCheck = false
        this._lastCheck = Date.now()
        const found = !err && stdout.toUpperCase().includes("WINWORD.EXE")
        this._alive = found
        if (found) {
          const lines = stdout.split('\n')
          for (const line of lines) {
            const m = line.match(/"WINWORD\.EXE","(\d+)"/i)
            if (m) {
              this._pid = m[1]
              break
            }
          }
        } else {
          this._pid = undefined
        }
      },
    )
  }
}
