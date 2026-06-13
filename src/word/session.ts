import { createRequire } from "node:module"
import { ProcessMonitor } from "./process-monitor.js"

const require = createRequire(import.meta.url)

interface WinaxModule {
  Object: new (progid: string) => Record<string, unknown>
}

export interface IWordSession {
  readonly application: Record<string, unknown>
  readonly activeDoc: Record<string, unknown> | null
  readonly activeDocPath: string | null
  setActiveDoc(doc: Record<string, unknown> | null): void
  setActiveDocPath(path: string | null): void
  ensureAlive(): void
  isAlive(): boolean
  start(): void
  quit(): void
  setOnLog(handler: (level: string, message: string) => void): void
  setScreenUpdating(on: boolean): void
  healthCheck(): boolean
  recover(): Promise<void>
  comCall<T>(fn: () => T): T
  markHealthy(): void
  markUnhealthy(): void
  isUnhealthy(): boolean
}

function defaultWinaxLoader(): WinaxModule {
  return require("winax") as WinaxModule
}

export class WordSession implements IWordSession {
  private word: Record<string, unknown> | null = null
  private _activeDoc: Record<string, unknown> | null = null
  private _activeDocPath: string | null = null
  private onLog: ((level: string, message: string) => void) | null = null
  private winaxMod: WinaxModule
  private _unhealthy = false
  private _recovering = false
  private monitor: ProcessMonitor

  constructor(winaxLoader?: () => WinaxModule) {
    this.winaxMod = (winaxLoader ?? defaultWinaxLoader)()
    this.monitor = new ProcessMonitor(3000)
  }

  setOnLog(handler: (level: string, message: string) => void): void {
    this.onLog = handler
  }

  start(): void {
    if (this.word) return
    this.monitor.start()
    this.onLog?.("info", "Creating Word.Application via winax")
    const app = new this.winaxMod.Object("Word.Application") as Record<string, unknown>
    app.Visible = true
    app.DisplayAlerts = 0
    app.AutomationSecurity = 3
    this.word = app
    this._unhealthy = false
    this.onLog?.("info", "Word.Application created successfully")
  }

  ensureAlive(): void {
    if (!this.isAlive()) this.start()
  }

  get application(): Record<string, unknown> {
    if (this._unhealthy && !this._recovering) {
      throw new Error("Session is unhealthy - recovery must be triggered via health check pipeline")
    }
    this.ensureAlive()
    return this.word!
  }

  get activeDoc(): Record<string, unknown> | null {
    return this._activeDoc
  }

  get activeDocPath(): string | null {
    return this._activeDocPath
  }

  setActiveDoc(doc: Record<string, unknown> | null): void {
    this._activeDoc = doc
  }

  setActiveDocPath(path: string | null): void {
    this._activeDocPath = path
  }

  markHealthy(): void {
    this._unhealthy = false
  }

  markUnhealthy(): void {
    this._unhealthy = true
  }

  isUnhealthy(): boolean {
    return this._unhealthy
  }

  quit(): void {
    this.monitor.stop()
    // 先检查 Word 进程是否还活着，不做 COM 调用
    const processAlive = this.healthCheck()

    if (!processAlive) {
      // Word 进程已死，只需清理内存状态
      this._activeDoc = null
      this._activeDocPath = null
      this.word = null
      this._unhealthy = false
      return
    }

    // Word 进程活着，尝试正常 COM 退出
    if (this.word) {
      try {
        const docs = this.word.Documents as { Count: number }
        if (docs.Count > 0) {
          for (let i = docs.Count; i >= 1; i--) {
            try {
              const doc = (this.word.Documents as { Item: (i: number) => Record<string, unknown> }).Item(i)
              const isSaved = doc.Saved as boolean
              if (!isSaved) {
                this.onLog?.("warn", `Closing unsaved document: ${(doc.Name as string) ?? "unknown"}`)
              }
              ;(doc.Close as (s: boolean) => void)(false)
            } catch {
              // ignore per-doc cleanup errors
            }
          }
        }
        ;(this.word.Quit as () => void)()
      } catch {
        // ignore quit errors
      }
      try {
        ;(this.word.Release as () => void)()
      } catch (e) {
        this.onLog?.("warn", `Release failed (non-critical): ${e}`)
      }
      this.word = null
      this._activeDoc = null
      this._activeDocPath = null
    }

    // 强制清理 COM 退出后的残留 WINWORD.EXE
    try {
      const { execSync } = require("node:child_process") as { execSync: (cmd: string, opts: { timeout: number; stdio: string }) => Buffer }
      execSync("taskkill /F /IM WINWORD.EXE", { timeout: 5000, stdio: "ignore" })
    } catch (e) {
      this.onLog?.("warn", `Force-kill WINWORD.EXE failed (normal if already exited): ${e}`)
    }
  }

  healthCheck(): boolean {
    return this.monitor.isAlive()
  }

  isAlive(): boolean {
    if (!this.word) return false

    // COM 存活检查: Version 属性（~5ms，不比 tasklist 慢且不依赖外部进程）
    try {
      const v = (this.word as Record<string, unknown>).Version
      if (v === undefined) {
        this.onLog?.("warn", "COM Version returned undefined — stale COM state")
        throw new Error("stale com")
      }
      this._unhealthy = false
      return true
    } catch {
      this.word = null
      this._unhealthy = true
      return false
    }
  }

  async recover(): Promise<void> {
    if (this._recovering) return
    this._recovering = true
    this.onLog?.("info", "Starting Word session recovery...")

    // 1. 清理内存状态（不做 COM 操作）
    const oldWord = this.word
    this.word = null
    this._activeDoc = null
    this._activeDocPath = null

    // 2. 尝试通知旧对象退出（不阻塞）
    if (oldWord) {
      try { ;(oldWord.Quit as () => void)() } catch { /* ignore */ }
      try { ;(oldWord.Release as () => void)() } catch { /* ignore */ }
    }

    // 3. 等待 COM 释放完成（通过 monitor 异步轮询，≤3s）
    await this.monitor.waitForExit(10000)

    // 4. 新建 COM 会话
    this._unhealthy = false
    this._recovering = false
    this.start()
    this.onLog?.("info", "Word session recovery complete")
  }

  comCall<T>(fn: () => T): T {
    try {
      const result = fn()
      this._unhealthy = false
      return result
    } catch (err) {
      this._unhealthy = true
      this.onLog?.("warn", `COM call failed: ${err instanceof Error ? err.message : String(err)}`)
      throw err
    }
  }

  setScreenUpdating(on: boolean): void {
    if (!this.word) return
    try {
      ;(this.word as Record<string, unknown>).ScreenUpdating = on
    } catch { /* ignore if Word not ready */ }
  }
}
