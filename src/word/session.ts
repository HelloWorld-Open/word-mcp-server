import { exec } from "node:child_process"
import { readdirSync, statSync, unlinkSync } from "node:fs"
import { join } from "node:path"
import { createRequire } from "node:module"
import { ProcessMonitor } from "./process-monitor.js"
import { TransientComError, FatalComError } from "./com-errors.js"
import type { IDocumentProxy, ISelectionProxy, IRangeProxy } from "./com-proxy/types.js"
import { DocumentProxy } from "./com-proxy/document-proxy.js"
import { SelectionProxy } from "./com-proxy/selection-proxy.js"
import { RangeProxy } from "./com-proxy/range-proxy.js"
import type { ILogger } from "../logger.js"

const require = createRequire(import.meta.url)

interface WinaxModule {
  Object: new (progid: string) => Record<string, unknown>
}

export interface IWordSession {
  readonly application: Record<string, unknown>
  readonly activeDoc: Record<string, unknown> | null
  readonly activeDocPath: string | null
  readonly logger: ILogger | null
  setActiveDoc(doc: Record<string, unknown> | null): void
  setActiveDocPath(path: string | null): void
  ensureAlive(): void
  isAlive(): boolean
  start(): void
  quit(): void
  setLogger(logger: ILogger): void
  setScreenUpdating(on: boolean): void
  withScreenOff<T>(fn: () => Promise<T>): Promise<T>
  healthCheck(): boolean
  recover(): Promise<void>
  comCall<T>(fn: () => T, signal?: AbortSignal): T
  markHealthy(): void
  markUnhealthy(): void
  isUnhealthy(): boolean
  setBusy(busy: boolean): void
  isBusy(): boolean
  getLastCallStart(): number
  getDocProxy(): IDocumentProxy
  getSelectionProxy(): ISelectionProxy
  wrapRange(raw: Record<string, unknown>): IRangeProxy
  log(level: string, msg: string): void
  lockPrintView(): void
  /** Try to detect and adopt an already-open ActiveDocument from COM. Returns true if adopted. */
  tryAdoptActiveDoc(): boolean
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

export class WordSession implements IWordSession {
  private word: Record<string, unknown> | null = null
  private _activeDoc: Record<string, unknown> | null = null
  private _activeDocPath: string | null = null
  private _logger: ILogger | null = null
  private _winaxMod: WinaxModule | null = null
  private _unhealthy = false
  private _recovering = false
  private _busy = false
  private _lastCallStart = 0
  private monitor: ProcessMonitor
  private _docProxy: IDocumentProxy | null = null
  private _selProxy: ISelectionProxy | null = null
  private _rangeProxyCache = new WeakMap<Record<string, unknown>, IRangeProxy>()

  constructor(winaxLoader?: () => WinaxModule) {
    this.monitor = new ProcessMonitor(3000)
    if (winaxLoader) this._winaxMod = winaxLoader()
  }

  private getWinax(): WinaxModule {
    if (!this._winaxMod) {
      this._winaxMod = require("winax") as WinaxModule
    }
    return this._winaxMod
  }

  getDocProxy(): IDocumentProxy {
    const doc = this._activeDoc
    if (!doc) throw new Error("No active document — cannot create DocumentProxy")
    if (!this._docProxy) {
      this._docProxy = new DocumentProxy(doc, this)
    }
    return this._docProxy
  }

  getSelectionProxy(): ISelectionProxy {
    if (!this._selProxy) {
      const sel = (() => {
        try {
          return this.comCall(() =>
            (this.application?.Selection as Record<string, unknown>) as Record<string, unknown>
          )
        } catch {
          return {} as Record<string, unknown>
        }
      })()
      this._selProxy = new SelectionProxy(sel, this)
    }
    return this._selProxy
  }

  wrapRange(raw: Record<string, unknown>): IRangeProxy {
    let p = this._rangeProxyCache.get(raw)
    if (!p) {
      p = new RangeProxy(raw, this)
      this._rangeProxyCache.set(raw, p)
    }
    return p
  }

  get logger(): ILogger | null {
    return this._logger
  }

  setLogger(logger: ILogger): void {
    this._logger = logger
  }

  log(level: string, msg: string): void {
    if (level === "error") this._logger?.error(msg)
    else if (level === "warn") this._logger?.warn(msg)
    else if (level === "debug") this._logger?.debug(msg)
    else this._logger?.info(msg)
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
    this._logger?.info("Creating Word.Application via winax")
    const winaxMod = this.getWinax()
    const app = new winaxMod.Object("Word.Application") as Record<string, unknown>
    app.AutomationSecurity = 3
    app.DisplayAlerts = 0
    try { ;(app as any).ShowStartupDialog = false } catch { /* Word 2013+ only */ }
    try { ;(app as any).Visible = true } catch (e) { this._logger?.warn(`Visible=true failed: ${e}`) }
    this.word = app
    this.lockPrintView()
    this.monitor.start()
    this._unhealthy = false
    this._logger?.info("Word.Application created successfully")
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
    this._docProxy = null
    this._selProxy = null
    this._rangeProxyCache = new WeakMap()
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

  setBusy(busy: boolean): void {
    this._busy = busy
    if (!busy) this._lastCallStart = 0
  }

  isBusy(): boolean {
    return this._busy
  }

  getLastCallStart(): number {
    return this._lastCallStart
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
              this._logger?.warn(`Closing unsaved document: ${(doc.Name as string) ?? "unknown"}`)
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
        this._logger?.warn(`Release failed (non-critical): ${e}`)
      }
      this.word = null
      this._activeDoc = null
      this._activeDocPath = null
      this._docProxy = null
      this._selProxy = null
      this._rangeProxyCache = new WeakMap()
    }

    // Fallback: kill by PID if process still running
    const pid = this.monitor.getPid()
    if (pid) {
      try {
        exec(`taskkill /PID ${pid} /F /T`, { timeout: 5000 })
        this._logger?.info(`Killed WINWORD.EXE (PID: ${pid})`)
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
      this._logger?.warn("Process-level check: WINWORD.EXE not found — stale COM proxy")
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
    this._busy = true
    this._logger?.info("Starting Word session recovery...")

    // 1. 先记录 PID，用于强制杀进程（必须在清空 word 之前获取）
    const pid = this.monitor.getPid()

    // 2. 清理内存状态（不做 COM 操作）
    const oldWord = this.word
    this.word = null
    this._activeDoc = null
    this._activeDocPath = null

    // 3. 释放 COM 引用（尽力，不阻塞）
    if (oldWord) {
      try { ;(oldWord.Release as () => void)() } catch { /* ignore */ }
    }

    // 4. 强制杀进程 — taskkill 比 oldWord.Quit() 更可靠（弹窗也能杀掉）
    if (pid) {
      try {
        await new Promise<void>((resolve, reject) => {
          exec(`taskkill /PID ${pid} /F /T`, { timeout: 5000 }, (err) => {
            if (err) reject(err)
            else resolve()
          })
        })
        this._logger?.info(`Killed WINWORD.EXE (PID: ${pid}) during recovery`)
      } catch {
        this._logger?.warn(`taskkill for PID ${pid} failed (may already be dead)`)
      }
    }

    // 5. 等待进程完全退出
    if (pid && this.monitor.lastCheckTime() > 0) {
      await this.monitor.waitForExit(10000)
    }

    // 6. 清理残留恢复文件
    this._cleanupRecoveryFiles()

    // 7. 新建 COM 会话
    this._busy = false
    this._unhealthy = false
    this._recovering = false
    this.start()
    this.monitor.markAlive()
    this._logger?.info("Word session recovery complete")
  }

  comCall<T>(fn: () => T, signal?: AbortSignal): T {
    if (signal?.aborted) {
      this._logger?.warn("COM call blocked: signal aborted (previous timeout or recovery)")
      throw new TransientComError("Call aborted — previous operation timed out or session is recovering")
    }
    if (this._busy) {
      this._logger?.warn("COM call blocked: session marked busy (request stacking barrier)")
      throw new TransientComError("Session busy — a prior operation is still in progress or timed out")
    }
    if (this._recovering) {
      this._logger?.warn("COM call blocked: session is recovering")
      throw new TransientComError("Session is recovering from a failure, try again shortly")
    }
    try {
      this._lastCallStart = Date.now()
      const result = fn()
      this._unhealthy = false
      return result
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err)
      this._logger?.warn(`COM call failed: ${msg}`)
      if (!isTransientComError(err) || !this.monitor.isAlive()) {
        this._unhealthy = true
        throw new FatalComError(msg)
      }
      throw new TransientComError(msg)
    } finally {
      this._lastCallStart = 0
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

  lockPrintView(): void {
    try {
      const win = (this.word as Record<string, unknown>).ActiveWindow as Record<string, unknown> | undefined
      if (!win) return
      const view = win.View as Record<string, unknown> | undefined
      if (view) {
        view.Type = 3 // wdPrintView
        this._logger?.debug("View locked to Print Layout")
      }
    } catch { /* ActiveWindow may not exist before a document is opened */ }
  }

  tryAdoptActiveDoc(): boolean {
    try {
      const app = this.application
      const doc = app.ActiveDocument as Record<string, unknown> | undefined
      if (!doc) return false
      this.setActiveDoc(doc)
      try {
        const fullName = doc.FullName as string
        this.setActiveDocPath(fullName)
      } catch {
        // FullName may fail for unsaved/untitled documents
      }
      return true
    } catch {
      return false
    }
  }
}
