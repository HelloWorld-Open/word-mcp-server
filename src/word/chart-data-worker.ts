/* eslint-disable */
import { createRequire } from "node:module"

const require = createRequire(import.meta.url)
const winax = require("winax") as {
  Object: new (progid: string) => Record<string, unknown>
}

interface ChartTask {
  id: number
  params: {
    docName: string
    inlineIndex: number
    data: (string | number)[][]
  }
}

let word: Record<string, unknown> | null = null

function ensureWord(): Record<string, unknown> {
  if (word) {
    try {
      const _ = (word.Version as string)
      return word
    } catch {
      try { ;(word.Quit as () => void)() } catch { /* already gone */ }
      word = null
    }
  }
  word = new winax.Object("Word.Application") as Record<string, unknown>
  return word
}

function setChartData(params: ChartTask["params"]): { ok: boolean; series: number } {
  const w = ensureWord()
  const docs = w.Documents as { Count: number; Item: (i: number) => Record<string, unknown> }

  let doc: Record<string, unknown> | null = null
  for (let i = 1; i <= docs.Count; i++) {
    const d = docs.Item(i)
    const fullName = (d.FullName as string) || ""
    const name = (d.Name as string) || ""
    if (fullName && fullName.endsWith(params.docName)) {
      doc = d
      break
    }
    if (name === params.docName) {
      doc = d
    }
  }
  if (!doc) throw new Error(`Document not found: ${params.docName}`)

  const inlineShapes = doc.InlineShapes as { Item: (i: number) => Record<string, unknown> }
  const shape = inlineShapes.Item(params.inlineIndex)
  const chart = shape.Chart as Record<string, unknown>

  const sc = chart.SeriesCollection as {
    Count: number; Item: (i: number) => { Delete: () => void }; NewSeries: () => Record<string, unknown>
  }
  while (sc.Count > 0) { sc.Item(1).Delete() }

  const headers = params.data[0] as (string | number)[]
  for (let c = 1; c < headers.length; c++) {
    const series = sc.NewSeries()
    series.Name = headers[c]
    const values: number[] = []
    for (let r = 1; r < params.data.length; r++) {
      values.push(Number(params.data[r][c]))
    }
    series.Values = values
    if (c === 1) {
      const cats: string[] = []
      for (let r = 1; r < params.data.length; r++) {
        cats.push(String(params.data[r][0]))
      }
      series.XValues = cats
    }
  }
  const seriesCount = params.data[0] ? params.data[0].length - 1 : 0
  return { ok: true, series: seriesCount }
}

process.on("message", (msg: unknown) => {
  const task = msg as ChartTask
  if (task.params == null) {
    if (task.id === -1) {
      if (word) {
        try { ;(word.Quit as () => void)() } catch (e) { console.warn("[chart-data-worker] Word.Quit failed on shutdown:", e) }
        word = null
      }
      process.exit(0)
    }
    return
  }
  try {
    const result = setChartData(task.params)
    process.send!({ id: task.id, result })
  } catch (err: unknown) {
    process.send!({ id: task.id, error: (err as Error)?.message ?? String(err) })
  }
})

process.on("uncaughtException", (err) => {
  try { process.send!({ id: -2, error: err.message }) } catch { /* parent may be gone */ }
  if (word) {
    try { ;(word.Quit as () => void)() } catch { /* Word may be gone */ }
    word = null
  }
  process.exit(1)
})

/* Best-effort: quit Word on any exit path (process.exit / normal exit).
 * Not called on SIGKILL where zombie is unavoidable. */
process.on("exit", () => {
  if (word) {
    try { ;(word.Quit as () => void)() } catch { /* ignore */ }
    word = null
  }
})
