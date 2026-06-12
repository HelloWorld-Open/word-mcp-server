import { exec } from "node:child_process"

export class ProcessMonitor {
  private _alive = false
  private _lastCheck = 0
  private _intervalMs: number
  private _timer: ReturnType<typeof setInterval> | null = null
  private _pendingCheck: boolean = false
  private _destroyed = false

  constructor(intervalMs = 3000) {
    this._intervalMs = intervalMs
  }

  start(): void {
    if (this._timer || this._destroyed) return
    this._check()
    this._timer = setInterval(() => this._check(), this._intervalMs)
  }

  stop(): void {
    if (this._timer) {
      clearInterval(this._timer)
      this._timer = null
    }
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

  async waitForExit(timeoutMs = 30000): Promise<boolean> {
    const deadline = Date.now() + timeoutMs
    while (Date.now() < deadline) {
      if (!this._alive) return true
      await new Promise<void>(r => setTimeout(r, 200))
    }
    return !this._alive
  }

  private _check(): void {
    if (this._pendingCheck || this._destroyed) return
    this._pendingCheck = true
    exec(
      'tasklist /FO CSV /NH',
      { timeout: 5000, encoding: "utf-8" },
      (err, stdout) => {
        this._pendingCheck = false
        this._lastCheck = Date.now()
        this._alive = !err && stdout.toUpperCase().includes("WINWORD.EXE")
      },
    )
  }
}
