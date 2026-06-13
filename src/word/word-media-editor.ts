import type { IWordSession } from "./session.js"
import { WordBase } from "./word-base.js"
import type { IChartDataBridge } from "./chart-data-bridge.js"

interface InsertImageParams {
  imagePath: string; width?: number; height?: number
}

export class WordMediaEditor extends WordBase {
  private static readonly TEXTBOX_ORIENTATION: Record<string, number> = {
    horizontal: 1, vertical: 5,
  }

  constructor(session: IWordSession, private chartBridge: IChartDataBridge) {
    super(session)
  }

  async insertImage(params: InsertImageParams): Promise<void> {
    this.collapseSelection()
    const doc = this.requireDoc()
    const inlineShapes = doc.InlineShapes as { AddPicture: (p: string, l?: boolean, s?: boolean, a?: unknown) => Record<string, unknown> }
    const selPre = this.getSelection()
    ;(selPre.TypeParagraph as () => void)()
    this.collapseSelection()
    const shape = inlineShapes.AddPicture(params.imagePath)
    if (params.width == null && params.height == null) {
      const maxWidth = 460
      if ((shape.Width as number) > maxWidth) {
        const ratio = maxWidth / (shape.Width as number)
        shape.Width = maxWidth
        shape.Height = (shape.Height as number) * ratio
      }
    } else {
      if (params.width != null) shape.Width = params.width
      if (params.height != null) shape.Height = params.height
    }
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
    const result = await this.chartBridge.setChartData(
      (doc as Record<string, unknown>).Name as string,
      inlineIndexBefore + 1,
      params.data,
    )
    try {
      const inlineShape = (doc.InlineShapes as { Item: (i: number) => Record<string, unknown> }).Item(inlineIndexBefore + 1)
      const chartData = (inlineShape.Chart as Record<string, unknown>).ChartData as Record<string, unknown> | undefined
      if (chartData) {
        const wb = chartData.Workbook as Record<string, unknown>
        try { ;(wb.Close as (s: boolean) => void)(false) } catch { }
      }
    } catch { }

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
}
