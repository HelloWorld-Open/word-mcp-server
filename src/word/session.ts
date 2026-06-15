import { exec } from "node:child_process"
import { readdirSync, statSync, unlinkSync } from "node:fs"
import { join } from "node:path"
import { createRequire } from "node:module"
import { ProcessMonitor } from "./process-monitor.js"
import { TransientComError, FatalComError } from "./com-errors.js"

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
  withScreenOff<T>(fn: () => Promise<T>): Promise<T>
  healthCheck(): boolean
  recover(): Promise<void>
  comCall<T>(fn: () => T): T
  markHealthy(): void
  markUnhealthy(): void
  isUnhealthy(): boolean
}

const TRANSIENT_HRESULTS = new Set([
  0x80010005, // RPC_E_CALL_REJECTED
  0x80010108, // RPC_E_SERVER_DIED
  0x800AC472, // CO_E_OBJNOTCONNECTED
  0x800706BA, // RPC_S_SERVER_UNAVAILABLE
  0x80010001, // RPC_E_SERVERFAULT
])

function isTransientComError(err: unknown): boolean {
  const code = (err as Record<string, unknown>)?.number ?? (err as Record<string, unknown>)?.code
  return typeof code === "number" && TRANSIENT_HRESULTS.has(code)
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

  private _cleanupRecoveryFiles(): void {
    const MAX_AGE_MS = 120 * 60 * 1000
    const now = Date.now()
    const dirs = [join(process.env.APPDATA ?? "", "Microsoft", "Word")]
    for (const dir of dirs) {
      let files: string[]
      try { files = readdirSync(dir) } catch { continue }
      for (const f of files) {
        if (!f.endsWith(".asd")) continue
        try {
          const fp = join(dir, f)
          const stat = statSync(fp)
          if (now - stat.mtimeMs > MAX_AGE_MS) unlinkSync(fp)
        } catch { /* skip locked or inaccessible */ }
      }
    }
  }

  start(): void {
    if (this.word) return
    this._cleanupRecoveryFiles()
    this.monitor.start()
    this.onLog?.("info", "Creating Word.Application via winax")
    const app = new this.winaxMod.Object("Word.Application") as Record<string, unknown>
    app.AutomationSecurity = 3
    app.DisplayAlerts = 0
    try { ;(app as any).ShowStartupDialog = false } catch { /* Word 2013+ only */ }
    try { ;(app as any).Visible = true } catch (e) { this.onLog?.("warn", `Visible=true failed: ${e}`) }
    this.word = app
    this._unhealthy = false
    this.onLog?.("info", "Word.Application created successfully")
  }

  ensureAlive(): void {
    if (!this.isAlive()) this.start()
  }

  get application(): Record<string, unknown> {
    if (this._unhealthy && !this._recovering) {
      throw new FatalComError("Session is unhealthy - recovery must be triggered via health check pipeline")
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

    // Always try COM Quit
    if (this.word) {
      try {
        const docs = this.word.Documents as { Count: number; Item: (i: number) => Record<string, unknown> }
        while (docs.Count > 0) {
          try {
            const doc = docs.Item(1) as Record<string, unknown>
            const isSaved = (doc.Saved as boolean) ?? true
            if (!isSaved) {
              this.onLog?.("warn", `Closing unsaved document: ${(doc.Name as string) ?? "unknown"}`)
            }
            ;(doc.Close as (s: boolean) => void)(false)
          } catch {
            break
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

    // Fallback: kill by PID if process still running
    const pid = this.monitor.getPid()
    if (pid) {
      try {
        exec(`taskkill /PID ${pid} /F /T`, { timeout: 5000 })
        this.onLog?.("info", `Killed WINWORD.EXE (PID: ${pid})`)
      } catch {
        // process already dead — ignore
      }
    }
  }

  healthCheck(): boolean {
    return this.monitor.isAlive()
  }

  isAlive(): boolean {
    if (!this.word) return false

    // 进程级检查（无 COM 调用，避免死进程 COM 挂起 30-60s）
    if (!this.monitor.isAlive()) {
      this.onLog?.("warn", "Process-level check: WINWORD.EXE not found — stale COM proxy")
      this.word = null
      this._unhealthy = true
      return false
    }

    this._unhealthy = false
    return true
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

    // 3. 等待 COM 释放完成（仅当 monitor 至少完成过一次检查时）
    //    跳过从未检查过的场景（测试环境或从未成功启动过），避免乐观 _alive 导致死等
    if (this.monitor.lastCheckTime() > 0) {
      await this.monitor.waitForExit(10000)
    }

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
      const msg = err instanceof Error ? err.message : String(err)
      this.onLog?.("warn", `COM call failed: ${msg}`)
      if (!isTransientComError(err) || !this.monitor.isAlive()) {
        this._unhealthy = true
        throw new FatalComError(msg)
      }
      throw new TransientComError(msg)
    }
  }

  setScreenUpdating(on: boolean): void {
    if (!this.word) return
    try {
      ;(this.word as Record<string, unknown>).ScreenUpdating = on
    } catch { /* ignore if Word not ready */ }
  }

  async withScreenOff<T>(fn: () => Promise<T>): Promise<T> {
    if (!this.word) return fn()
    try {
      try { ;(this.word as Record<string, unknown>).ScreenUpdating = false } catch { /* ignore */ }
      return await fn()
    } finally {
      try { ;(this.word as Record<string, unknown>).ScreenUpdating = true } catch { /* ignore */ }
    }
  }
}
