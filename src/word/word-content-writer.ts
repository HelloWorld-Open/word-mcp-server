import type { IWordSession } from "./session.js"
import type { IChartDataBridge } from "./chart-data-bridge.js"
import { WordBase } from "./word-base.js"
import { parseBlocks } from "./markdown-parser.js"
import { MarkdownRenderer } from "./markdown-renderer.js"
import { WordFormatter } from "./word-formatter.js"


const AUTO_FIT: Record<string, number> = {
  fixed: 0, contents: 1, window: 2,
}
const COLOR_INDEX: Record<string, number> = {
  auto: 0, black: 1, blue: 2, turquoise: 3, bright_green: 4, pink: 5,
  red: 6, yellow: 7, white: 8, dark_blue: 9, teal: 10, green: 11,
  violet: 12, dark_red: 13, dark_yellow: 14, gray_50: 15, gray_25: 16,
}

export class WordContentWriter extends WordBase {
  private renderer: MarkdownRenderer

  constructor(session: IWordSession, private chartBridge: IChartDataBridge) {
    super(session)
    this.renderer = new MarkdownRenderer({
      getSelection: () => this.getSelProxy(),
      requireDoc: () => this.getDocProxy(),
      goToEnd: () => this.goToEnd(),
      withScreenOff: (fn) => this.session.withScreenOff(fn),
    })
  }

  async writeBlocks(markdown: string): Promise<{ blocks: number; chars: number }> {
    const blocks = parseBlocks(markdown)
    if (blocks.length === 0) return { blocks: 0, chars: 0 }
    let totalChars = 0
    this.collapseSelection()

    const BATCH_MS = 200
    let bi = 0
    let needsEnd = true

    while (bi < blocks.length) {
      const batchStart = Date.now()
      do {
        totalChars += await this.renderer.renderBlock(blocks[bi], bi, blocks.length, "end", !needsEnd)
        needsEnd = blocks[bi].type === "table"
        bi++
      } while (bi < blocks.length && Date.now() - batchStart < BATCH_MS)

      await new Promise(resolve => setImmediate(resolve))
    }

    try {
      this.getSelProxy().typeParagraph()
    } catch { /* ignore */ }
    return { blocks: blocks.length, chars: totalChars }
  }

  async insertAtCursor(markdown: string): Promise<{ blocks: number; chars: number }> {
    const blocks = parseBlocks(markdown)
    if (blocks.length === 0) return { blocks: 0, chars: 0 }
    let totalChars = 0
    this.collapseSelection()

    const BATCH_MS = 200
    let bi = 0

    while (bi < blocks.length) {
      const batchStart = Date.now()
      do {
        totalChars += await this.renderer.renderBlock(blocks[bi], bi, blocks.length, "cursor")
        bi++
      } while (bi < blocks.length && Date.now() - batchStart < BATCH_MS)

      await new Promise(resolve => setImmediate(resolve))
    }

    return { blocks: blocks.length, chars: totalChars }
  }

  async replaceVariables(variables: Record<string, string>): Promise<{ key: string; count: number }[]> {
    const doc = this.getDocProxy()
    const results: { key: string; count: number }[] = []

    const sel = this.getSelProxy()
    const savedStart = sel.getStart()
    const savedEnd = sel.getEnd()

    try {
      for (const [key, value] of Object.entries(variables)) {
        const findText = `{{${key}}}`
        let count = 0

        const docEnd = doc.getContent().getEnd()
        const searchRange = doc.getRange(0, docEnd)
        const find = searchRange.getFind()
        ;(find.ClearFormatting as () => void)()
        find.Text = findText
        find.Forward = true
        find.Wrap = 0
        find.Format = false
        find.MatchCase = true
        find.MatchWholeWord = false
        find.MatchWildcards = false

        const breakCount = 100000
        while (count < breakCount) {
          const found = (find.Execute as (...args: unknown[]) => boolean)(findText, false, false, false, false, false, true, 0, false, "", 0)
          if (!found) break
          count++
            searchRange.setStart(searchRange.getEnd())
            searchRange.setEnd(docEnd)
        }

        if (count > 0) {
          const replaceRange = doc.getRange(0, docEnd)
          const replaceFind = replaceRange.getFind()
          ;(replaceFind.ClearFormatting as () => void)()
          ;((replaceFind.Replacement as Record<string, unknown>).ClearFormatting as () => void)()
          replaceFind.Text = findText
          ;(replaceFind.Replacement as Record<string, unknown>).Text = value
          replaceFind.Forward = true
          replaceFind.Wrap = 0
          replaceFind.Format = false
          replaceFind.MatchCase = true
          replaceFind.MatchWholeWord = false
          replaceFind.MatchWildcards = false
          ;(replaceFind.Execute as (...args: unknown[]) => unknown)(findText, false, false, false, false, false, true, 0, false, value, 2)
        }

        results.push({ key, count })
      }
    } finally {
      try {
        const restoreRange = doc.getRange(savedStart, savedEnd)
        restoreRange.select()
      } catch { /* ignore */ }
    }

    return results
  }

  async insertParagraph(count?: number): Promise<void> {
    this.collapseSelection()
    const sel = this.getSelProxy()
    const n = Math.max(1, count ?? 1)
    for (let i = 0; i < n; i++) {
      sel.typeParagraph()
    }
  }

  async insertPageBreak(): Promise<void> {
    this.collapseSelection()
    this.getSelProxy().insertBreak(7)
  }

  async insertHorizontalLine(): Promise<void> {
    this.collapseSelection()
    this.getSelProxy().addHorizontalLine()
  }

  async insertList(type: "bullet" | "number", items: string[]): Promise<void> {
    this.collapseSelection()
    const sel = this.getSelProxy()
    const word = this.getWord()
    const applyList = () => {
      const lf = sel.getRange().getListFormat()
      if (type === "bullet") {
        ;(lf.ApplyBulletDefault as () => void)()
      } else {
        ;(lf.ApplyNumberDefault as () => void)()
      }
    }
    const removeList = () => {
      const lf = sel.getRange().getListFormat()
      ;(lf.RemoveNumbers as () => void)()
    }
    try {
      try { word.ScreenUpdating = false } catch { /* ignore */ }
      applyList()
      const TIME_BUDGET = 50
      let batchStart = Date.now()
      for (let i = 0; i < items.length; i++) {
        sel.typeText(items[i])
        if (i < items.length - 1) {
          sel.typeParagraph()
        }
        if (i < items.length - 1 && Date.now() - batchStart >= TIME_BUDGET) {
          try { word.ScreenUpdating = true } catch { /* ignore */ }
          try { ;(word.ScreenRefresh as () => void)() } catch { /* ignore */ }
          await new Promise(resolve => setImmediate(resolve))
          try { word.ScreenUpdating = false } catch { /* ignore */ }
          batchStart = Date.now()
        }
      }
      sel.typeParagraph()
      removeList()
    } finally {
      try { word.ScreenUpdating = true } catch { /* ignore */ }
    }
  }

  async addHyperlink(text: string, address: string, subAddress?: string, screenTip?: string): Promise<void> {
    const sel = this.getSelProxy()
    const doc = this.getDocProxy()
    const hyperlinks = doc.getHyperlinks()

    const hasSelection = sel.getType() !== 1
    if (hasSelection) {
      hyperlinks.add(sel.getRange().raw, address, subAddress, screenTip, text)
    } else {
      this.collapseSelection()
      const range = this.getSelProxy().getRange().raw
      hyperlinks.add(range, address, subAddress, screenTip, text)
    }
    this.goToEnd()
    sel.collapse(0)
  }

  async addFootnote(text: string): Promise<void> {
    try {
      this.collapseSelection()
      const doc = this.getDocProxy()
      const footnotes = doc.getFootnotes()
      await this.session.withScreenOff(async () => {
        try {
          footnotes.add(this.getSelProxy().getRange().raw, text)
        } catch {
          try {
            try { this.getSelProxy().setStyle("Normal") } catch { /* ignore */ }
            this.cursor.goToEnd()
            const sel = this.getSelProxy()
            sel.typeText(`[脚注: ${text}]`)
          } catch { /* ignore */ }
        }
      })
      const end = doc.getContent().getEnd()
      const endRange = doc.getRange(end, end)
      endRange.select()
      this.getSelProxy().collapse(0)
    } catch { /* addFootnote ultimate fallback — never throw */ }
  }

  async insertFile(path: string): Promise<void> {
    this.collapseSelection()
    const range = this.getSelProxy().getRange()
    range.insertFile(path)
  }

  async insertSectionBreak(type?: string): Promise<void> {
    this.collapseSelection()
    const map: Record<string, number> = {
      nextPage: 8, continuous: 9, evenPage: 10, oddPage: 11,
    }
    this.getSelProxy().collapse(1)
    this.getSelProxy().insertBreak(map[type ?? "nextPage"] ?? 8)
  }

  async setColumns(count: number, spacing?: number): Promise<void> {
    const sections = this.getDocProxy().getSections()
    const ps = sections.item(sections.count).getPageSetup()
    const textColumns = ps.TextColumns as Record<string, unknown>
    ;(textColumns.SetCount as (c: number) => void)(count)
    if (spacing != null) {
      const spacingPoints = Math.round(spacing * 28.3465)
      ;(textColumns.Spacing as number) = spacingPoints
    }
  }

  async insertImage(params: { imagePath: string; width?: number; height?: number }): Promise<void> {
    this.collapseSelection()
    const doc = this.getDocProxy()
    const inlineShapes = doc.getInlineShapes()
    const selPre = this.getSelProxy()
    selPre.typeParagraph()
    this.collapseSelection()
    const shape = inlineShapes.addPicture(params.imagePath)
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
      this.goToEnd()
      this.getSelProxy().collapse(0)
      this.getSelProxy().typeParagraph()
    } catch { /* ignore */ }
  }

  async insertChart(params: {
    type: string; data: (string | number)[][]; title?: string; width?: number; height?: number
  }): Promise<{ type: string; series: number }> {
    this.collapseSelection()
    const doc = this.getDocProxy()
    const typeMap: Record<string, number> = { column: 51, bar: 57, line: 4, pie: 5, area: 1 }
    const chartType = typeMap[params.type] ?? 51
    const rawInlineShapes = doc.getInlineShapes()
    const inlineIndexBefore = rawInlineShapes.count
    const selPre = this.getSelProxy()
    selPre.typeParagraph()
    this.collapseSelection()
    const shape = rawInlineShapes.addChart2(-1, chartType, this.getSelProxy().getRange().raw)
    if (params.width != null) shape.Width = params.width
    if (params.height != null) shape.Height = params.height
    const chart = shape.Chart as Record<string, unknown>
    if (params.title) {
      ;(chart.HasTitle as boolean) = true
      ;((chart.ChartTitle as Record<string, unknown>).Text as string) = params.title
    }
    const result = await this.chartBridge.setChartData(
      doc.getName(),
      inlineIndexBefore + 1,
      params.data,
    )
    try {
      const inlineShape = doc.getInlineShapes().item(inlineIndexBefore + 1)
      const chartData = (inlineShape.Chart as Record<string, unknown>).ChartData as Record<string, unknown> | undefined
      if (chartData) {
        const wb = chartData.Workbook as Record<string, unknown>
        try { ;(wb.Close as (s: boolean) => void)(false) } catch { /* ignore */ }
      }
    } catch { /* ignore */ }

    try {
      this.goToEnd()
      this.getSelProxy().collapse(0)
      this.getSelProxy().typeParagraph()
    } catch { /* ignore */ }
    return { type: params.type, series: result.series }
  }

  async insertTextbox(params: {
    text: string; width?: number; height?: number; orientation?: unknown;
    positionLeft?: number; positionTop?: number
  }): Promise<{ width: number; height: number }> {
    this.collapseSelection()
    const doc = this.getDocProxy()
    const TEXTBOX_ORIENTATION: Record<string, number> = {
      horizontal: 1, vertical: 5,
    }
    const shapes = doc.getShapes() as { AddTextbox: (o: number, l: number, t: number, w: number, h: number) => Record<string, unknown> }
    const shape = shapes.AddTextbox(
      this.numOrEnum(params.orientation ?? "horizontal", TEXTBOX_ORIENTATION),
      params.positionLeft ?? 50,
      params.positionTop ?? 50,
      params.width ?? 200,
      params.height ?? 100,
    )
    ;((shape.TextFrame as Record<string, unknown>).TextRange as Record<string, unknown>).Text = params.text
    try {
      this.goToEnd()
      this.getSelProxy().collapse(0)
      this.getSelProxy().typeParagraph()
    } catch { /* ignore */ }
    return { width: params.width ?? 200, height: params.height ?? 100 }
  }

  async insertTable(params: {
    rows: number; columns: number; data?: string[][]; autoFitBehavior?: unknown
  }): Promise<{ rows: number; columns: number; failCount: number }> {
    this.collapseSelection()
    const doc = this.getDocProxy()
    const sel = this.getSelProxy()
    try { sel.setStyle("Normal") } catch { /* ignore */ }
    const tables = doc.getTables()
    const table = tables.add(sel.getRange().raw, params.rows, params.columns)
    let failCount = 0
    // Wrap cell-fill + styling in withScreenOff to avoid per-call UI repaints
    await this.session.withScreenOff(async () => {
      if (params.data) {
        for (let r = 0; r < params.data.length && r < params.rows; r++) {
          const row = params.data[r]
          for (let c = 0; c < row.length && c < params.columns; c++) {
            try {
              const cellObj = (table.Cell as (r: number, c: number) => Record<string, unknown>)(r + 1, c + 1)
              const cellRange = cellObj.Range as Record<string, unknown>
              cellRange.Text = row[c]
            } catch { failCount++ }
          }
        }
      }
      WordFormatter.applyDefaultBorders(table)
      if (params.data && params.data.length > 0) {
        let styled = false
        try {
          const styles = doc.getStyles()
          const candidates = ["Grid Table 4 - Accent 1", "网格表4 - 着色1", "Grid Table 4"]
          for (const name of candidates) {
            for (let i = 1; i <= styles.count; i++) {
              if (((styles.item(i).NameLocal as string) ?? "").toLowerCase() === name.toLowerCase()) {
                ;(table.Style as string) = name
                styled = true
                break
              }
            }
            if (styled) break
          }
        } catch { /* style may not exist */ }
        if (!styled) {
          try {
            const firstRow = (table.Rows as { Item: (i: number) => Record<string, unknown> }).Item(1)
            ;((firstRow.Range as Record<string, unknown>).Font as Record<string, unknown>).Bold = true
            ;(firstRow.Shading as Record<string, unknown>).BackgroundPatternColor = 0xF3E2D9
          } catch { /* ignore */ }
        }
      }
    })
    // AutoFitBehavior after screen re-enabled — single repaint only
    if (params.autoFitBehavior != null) {
      try { ;(table.AutoFitBehavior as (b: number) => void)(this.numOrEnum(params.autoFitBehavior, AUTO_FIT)) } catch { /* ignore */ }
    }
    // Move cursor reliably past the table with explicit wdParagraph unit (4)
    try {
      const rawTableRange = table.Range as Record<string, unknown>
      const nextRange = (rawTableRange.Next as (unit?: number) => Record<string, unknown> | undefined)(4)
      if (nextRange) {
        const nextStart = nextRange.Start as number
        if (typeof nextStart === "number") {
          doc.getRange(nextStart, nextStart).select()
          sel.collapse(0)
        } else {
          throw new Error("no next range start")
        }
      } else {
        throw new Error("no content after table")
      }
    } catch {
      // Fallback: table at end of document — InsertParagraph to create space
      try {
        const contentEnd = doc.getContent().getEnd()
        const insertAt = doc.getRange(contentEnd, contentEnd).raw
        ;(insertAt.InsertParagraph as () => void)()
        const newEnd = doc.getContent().getEnd()
        doc.getRange(newEnd, newEnd).select()
      } catch { /* ignore */ }
      sel.collapse(0)
    }
    sel.typeParagraph()
    return { rows: params.rows, columns: params.columns, failCount }
  }

  async textToTable(params: { separator?: string; autoFitBehavior?: unknown }): Promise<{ rows: number; columns: number }> {
    this.collapseSelection()
    const sel = this.getSelProxy()
    const range = sel.getRange()
    const sep = params.separator ?? "\t"
    const t = range.convertToTable(sep)
    if (params.autoFitBehavior != null) {
      ;(t as Record<string, unknown>).AutoFitBehavior = this.numOrEnum(params.autoFitBehavior, AUTO_FIT)
    }
    const rows = (t as Record<string, unknown>).Rows as { Count: number }
    const cols = (t as Record<string, unknown>).Columns as { Count: number }
    return { rows: rows.Count, columns: cols.Count }
  }
}
