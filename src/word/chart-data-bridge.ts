import { fork, type ChildProcess } from "node:child_process"
import { resolve, dirname } from "node:path"
import { fileURLToPath } from "node:url"

const __filename = fileURLToPath(import.meta.url)
const __dirname = dirname(__filename)

const IDLE_TIMEOUT_MS = parseInt(process.env.CHART_WORKER_IDLE_TIMEOUT ?? "60000", 10)
const CHART_TIMEOUT_MS = parseInt(process.env.CHART_OP_TIMEOUT ?? "15000", 10)

export interface IChartDataBridge {
  setChartData(docName: string, inlineIndex: number, data: (string | number)[][]): Promise<{ ok: boolean; series: number }>
  dispose(): void
}

let nextTaskId = 1

export class ChartDataBridge implements IChartDataBridge {
  constructor(private options?: { forkFn?: typeof fork }) {}

  private worker: ChildProcess | null = null
  private pending: Map<number, { resolve: (v: { ok: boolean; series: number }) => void; timer: ReturnType<typeof setTimeout> }> = new Map()
  private idleTimer: ReturnType<typeof setTimeout> | null = null
  private restartBlockedUntil = 0

  setChartData(docName: string, inlineIndex: number, data: (string | number)[][]): Promise<{ ok: boolean; series: number }> {
    const id = nextTaskId++
    return new Promise((resolvePromise) => {
      let child: ChildProcess
      try {
        child = this.ensureWorker()
      } catch {
        resolvePromise({ ok: false, series: 1 })
        return
      }
      const timer = setTimeout(() => {
        this.pending.delete(id)
        resolvePromise({ ok: false, series: 1 })
        if (this.pending.size === 0) this.terminateWorker()
      }, CHART_TIMEOUT_MS)
      this.pending.set(id, { resolve: resolvePromise, timer })
      try {
        child.send({ id, params: { docName, inlineIndex, data } })
      } catch {
        clearTimeout(timer)
        this.pending.delete(id)
        resolvePromise({ ok: false, series: 1 })
      }
    })
  }

  dispose(): void {
    this.terminateWorker()
  }

  private ensureWorker(): ChildProcess {
    if (this.worker && !this.worker.killed) return this.worker
    const wait = this.restartBlockedUntil - Date.now()
    if (wait > 0) throw new Error(`Worker restart throttled (${wait}ms remaining)`)
    const workerPath = resolve(__dirname, "chart-data-worker.js")
    const forkFn = this.options?.forkFn ?? fork
    const child = forkFn(workerPath, [], {
      stdio: ["pipe", "pipe", "inherit", "ipc"],
      env: { ...process.env },
    })
    child.on("message", (msg: unknown) => {
      const response = msg as { id: number; result?: { ok: boolean; series: number }; error?: string }
      const pending = this.pending.get(response.id)
      if (pending) {
        clearTimeout(pending.timer)
        this.pending.delete(response.id)
        if (response.error) {
          pending.resolve({ ok: false, series: 1 })
        } else {
          pending.resolve(response.result ?? { ok: false, series: 1 })
        }
      }
      this.resetIdleTimer()
    })
    child.on("exit", () => {
      this.worker = null
      this.restartBlockedUntil = Date.now() + 1000
      for (const [, pending] of this.pending) {
        clearTimeout(pending.timer)
        pending.resolve({ ok: false, series: 1 })
      }
      this.pending.clear()
    })
    child.on("error", () => {
      this.worker = null
      this.restartBlockedUntil = Date.now() + 1000
      for (const [, pending] of this.pending) {
        clearTimeout(pending.timer)
        pending.resolve({ ok: false, series: 1 })
      }
      this.pending.clear()
    })
    child.stdout?.on("data", () => { })
    child.stdout?.on("error", () => { })
    this.worker = child
    this.resetIdleTimer()
    return child
  }

  private resetIdleTimer(): void {
    if (this.idleTimer) {
      clearTimeout(this.idleTimer)
    }
    if (this.worker && !this.worker.killed) {
      this.idleTimer = setTimeout(() => {
        this.terminateWorker()
      }, IDLE_TIMEOUT_MS)
    }
  }

  private terminateWorker(): void {
    if (this.worker && !this.worker.killed) {
      if (this.worker.stdout) {
        this.worker.stdout.removeAllListeners("data")
        this.worker.stdout.removeAllListeners("error")
      }
      try {
        this.worker.send({ id: -1, params: null }) // tell worker to quit Word gracefully
      } catch { /* worker may already be gone — force-kill will handle it */ }
      const w = this.worker
      setTimeout(() => { if (!w.killed) w.kill() }, 2000)
    }
    this.worker = null
    if (this.idleTimer) {
      clearTimeout(this.idleTimer)
      this.idleTimer = null
    }
  }
}
