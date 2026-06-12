import { fork, type ChildProcess } from "node:child_process"
import { fileURLToPath } from "node:url"
import { resolve, dirname } from "node:path"
import { WordBase } from "./word-base.js"

const __filename = fileURLToPath(import.meta.url)
const __dirname = dirname(__filename)

interface InsertImageParams {
  imagePath: string; width?: number; height?: number
}

const IDLE_TIMEOUT_MS = parseInt(process.env.CHART_WORKER_IDLE_TIMEOUT ?? "60000", 10)
const CHART_TIMEOUT_MS = parseInt(process.env.CHART_OP_TIMEOUT ?? "15000", 10)

let nextTaskId = 1

export class WordMediaEditor extends WordBase {
  private static readonly TEXTBOX_ORIENTATION: Record<string, number> = {
    horizontal: 1, vertical: 5,
  }

  private _chartWorker: ChildProcess | null = null
  private _chartWorkerIdleTimer: ReturnType<typeof setTimeout> | null = null
  private _chartPending: Map<number, { resolve: (v: { ok: boolean; series: number }) => void; timer: ReturnType<typeof setTimeout> }> = new Map()

  private ensureChartWorker(): ChildProcess {
    if (this._chartWorker && !this._chartWorker.killed) return this._chartWorker
    const workerPath = resolve(__dirname, "chart-data-worker.js")
    const child = fork(workerPath, [], {
      stdio: ["pipe", "pipe", "inherit", "ipc"],
      env: { ...process.env },
    })
    child.on("message", (msg: unknown) => {
      const response = msg as { id: number; result?: { ok: boolean; series: number }; error?: string }
      const pending = this._chartPending.get(response.id)
      if (pending) {
        clearTimeout(pending.timer)
        this._chartPending.delete(response.id)
        if (response.error) {
          pending.resolve({ ok: false, series: 1 })
        } else {
          pending.resolve(response.result ?? { ok: false, series: 1 })
        }
      }
      this.resetChartIdleTimer()
    })
    child.on("exit", () => {
      this._chartWorker = null
      for (const [, pending] of this._chartPending) {
        clearTimeout(pending.timer)
        pending.resolve({ ok: false, series: 1 })
      }
      this._chartPending.clear()
    })
    child.on("error", () => {
      this._chartWorker = null
      for (const [, pending] of this._chartPending) {
        clearTimeout(pending.timer)
        pending.resolve({ ok: false, series: 1 })
      }
      this._chartPending.clear()
    })
    child.stdout?.on("data", () => { /* drain, not used */ })
    this._chartWorker = child
    this.resetChartIdleTimer()
    return child
  }

  private resetChartIdleTimer(): void {
    if (this._chartWorkerIdleTimer) {
      clearTimeout(this._chartWorkerIdleTimer)
    }
    if (this._chartWorker && !this._chartWorker.killed) {
      this._chartWorkerIdleTimer = setTimeout(() => {
        this.terminateChartWorker()
      }, IDLE_TIMEOUT_MS)
    }
  }

  private terminateChartWorker(): void {
    if (this._chartWorker && !this._chartWorker.killed) {
      try {
        this._chartWorker.send({ id: -1, params: null })
      } catch { }
      const w = this._chartWorker
      setTimeout(() => { if (!w.killed) w.kill() }, 2000)
    }
    this._chartWorker = null
    if (this._chartWorkerIdleTimer) {
      clearTimeout(this._chartWorkerIdleTimer)
      this._chartWorkerIdleTimer = null
    }
  }

  async insertImage(params: InsertImageParams): Promise<void> {
    this.collapseSelection()
    const doc = this.requireDoc()
    const inlineShapes = doc.InlineShapes as { AddPicture: (p: string, l?: boolean, s?: boolean, a?: unknown) => Record<string, unknown> }
    const selPre = this.getSelection()
    ;(selPre.TypeParagraph as () => void)()
    this.collapseSelection()
    const shape = inlineShapes.AddPicture(params.imagePath)
    if (params.width != null) shape.Width = params.width
    if (params.height != null) shape.Height = params.height
    try {
      const sel = this.getSelection()
      ;(sel.EndKey as (u: number) => void)(6)
      ;(sel.Collapse as (d: number) => void)(0)
      ;(sel.TypeParagraph as () => void)()
    } catch { }
  }

  async insertChart(params: {
    type: string; data: (string | number)[][]; title?: string; width?: number; height?: number
  }): Promise<{ type: string; series: number }> {
    this.collapseSelection()
    const doc = this.requireDoc()
    const typeMap: Record<string, number> = { column: 51, bar: 57, line: 4, pie: 5, area: 1 }
    const chartType = typeMap[params.type] ?? 51
    const inlineShapes = doc.InlineShapes as { Count: number; AddChart2: (style: number, type: number, range: unknown) => Record<string, unknown> }
    const inlineIndexBefore = inlineShapes.Count
    const selPre = this.getSelection()
    ;(selPre.TypeParagraph as () => void)()
    this.collapseSelection()
    const shape = inlineShapes.AddChart2(-1, chartType, this.getSelection().Range)
    if (params.width != null) shape.Width = params.width
    if (params.height != null) shape.Height = params.height
    const chart = shape.Chart as Record<string, unknown>
    if (params.title) {
      ;(chart.HasTitle as boolean) = true
      ;((chart.ChartTitle as Record<string, unknown>).Text as string) = params.title
    }
    const result = await this.setChartDataViaWorker({
      docName: (doc as Record<string, unknown>).Name as string,
      inlineIndex: inlineIndexBefore + 1,
      data: params.data,
    })
    try {
      const inlineShape = (doc.InlineShapes as { Item: (i: number) => Record<string, unknown> }).Item(inlineIndexBefore + 1)
      const chartData = (inlineShape.Chart as Record<string, unknown>).ChartData as Record<string, unknown> | undefined
      if (chartData) {
        const wb = chartData.Workbook as Record<string, unknown>
        try { ;(wb.Close as (s: boolean) => void)(false) } catch { }
      }
    } catch { /* data sheet may not be open */ }

    try {
      const sel = this.getSelection()
      ;(sel.EndKey as (u: number) => void)(6)
      ;(sel.Collapse as (d: number) => void)(0)
      ;(sel.TypeParagraph as () => void)()
    } catch { }
    return { type: params.type, series: result.series }
  }

  async insertTextbox(params: {
    text: string; width?: number; height?: number; orientation?: unknown;
    positionLeft?: number; positionTop?: number
  }): Promise<{ width: number; height: number }> {
    this.collapseSelection()
    const doc = this.requireDoc()
    const shapes = doc.Shapes as { AddTextbox: (o: number, l: number, t: number, w: number, h: number) => Record<string, unknown> }
    const shape = shapes.AddTextbox(
      this.numOrEnum(params.orientation ?? "horizontal", WordMediaEditor.TEXTBOX_ORIENTATION),
      params.positionLeft ?? 50,
      params.positionTop ?? 50,
      params.width ?? 200,
      params.height ?? 100,
    )
    ;((shape.TextFrame as Record<string, unknown>).TextRange as Record<string, unknown>).Text = params.text
    try {
      const sel = this.getSelection()
      ;(sel.EndKey as (u: number) => void)(6)
      ;(sel.Collapse as (d: number) => void)(0)
      ;(sel.TypeParagraph as () => void)()
    } catch { }
    return { width: params.width ?? 200, height: params.height ?? 100 }
  }

  private setChartDataViaWorker(params: {
    docName: string; inlineIndex: number; data: (string | number)[][]
  }): Promise<{ ok: boolean; series: number }> {
    const id = nextTaskId++
    return new Promise((resolvePromise) => {
      let child: ChildProcess
      try {
        child = this.ensureChartWorker()
      } catch {
        resolvePromise({ ok: false, series: 1 })
        return
      }
      const timer = setTimeout(() => {
        this._chartPending.delete(id)
        resolvePromise({ ok: false, series: 1 })
        if (this._chartPending.size === 0) this.terminateChartWorker()
      }, CHART_TIMEOUT_MS)
      this._chartPending.set(id, { resolve: resolvePromise, timer })
      try {
        child.send({ id, params })
      } catch {
        clearTimeout(timer)
        this._chartPending.delete(id)
        resolvePromise({ ok: false, series: 1 })
      }
    })
  }
}
