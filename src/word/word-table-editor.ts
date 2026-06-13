import { WordBase } from "./word-base.js"
import { WordMcpError } from "../security/errors.js"

interface InsertTableParams {
  rows: number; columns: number; data?: string[][]; autoFitBehavior?: unknown
}

interface EditCellParams {
  tableIndex?: number; row: number; column: number; text: string
}

interface AddTableRowParams {
  tableIndex?: number; data?: string[]
}

interface DeleteTableRowParams {
  tableIndex?: number; rowIndex: number
}

export class WordTableEditor extends WordBase {
  private static readonly AUTO_FIT: Record<string, number> = {
    fixed: 0, contents: 1, window: 2,
  }
  private static readonly V_ALIGN: Record<string, number> = {
    top: 0, center: 1, bottom: 3,
  }
  private static readonly BORDER_STYLE: Record<string, number> = {
    none: 0, single: 1, dot: 2, dash_small: 3, dash_large: 4, dash: 5, dash_dot: 6, double: 7,
  }
  private static readonly COLOR_INDEX: Record<string, number> = {
    auto: 0, black: 1, blue: 2, turquoise: 3, bright_green: 4, pink: 5,
    red: 6, yellow: 7, white: 8, dark_blue: 9, teal: 10, green: 11,
    violet: 12, dark_red: 13, dark_yellow: 14, gray_50: 15, gray_25: 16,
  }
  private static readonly UNDERLINE: Record<string, number> = {
    none: 0, single: 1, double: 3, wavy: 11,
  }

  async insertTable(params: InsertTableParams): Promise<{ rows: number; columns: number; failCount: number }> {
    this.collapseSelection()
    const doc = this.requireDoc()
    const sel = this.getSelection()
    const tables = doc.Tables as { Add: (r: unknown, rows: number, cols: number) => Record<string, unknown> }
    const table = tables.Add(sel.Range, params.rows, params.columns)
    if (params.autoFitBehavior != null) {
      ;(table.AutoFitBehavior as (b: number) => void)(this.numOrEnum(params.autoFitBehavior, WordTableEditor.AUTO_FIT))
    }
    let failCount = 0
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
    WordTableEditor.applyDefaultBorders(table)
    // Auto-format: attempt built-in style, fallback to manual header formatting
    if (params.data && params.data.length > 0) {
      let styled = false
      try {
        const styles = doc.Styles as unknown as { Count: number; Item: (i: number) => Record<string, unknown> }
        const candidates = ["Grid Table 4 - Accent 1", "网格表4 - 着色1", "Grid Table 4"]
        for (const name of candidates) {
          for (let i = 1; i <= styles.Count; i++) {
            if (((styles.Item(i).NameLocal as string) ?? "").toLowerCase() === name.toLowerCase()) {
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
    ;(sel.EndKey as (u: number) => void)(6)
    ;(sel.TypeParagraph as () => void)()
    return { rows: params.rows, columns: params.columns, failCount }
  }

  private requireTable(params: { tableIndex?: number }): Record<string, unknown> {
    const doc = this.requireDoc()
    const tables = doc.Tables as { Count: number; Item: (i: number) => Record<string, unknown> }
    const idx = (params.tableIndex ?? 1)
    if (idx < 1 || idx > tables.Count) {
      throw new WordMcpError(`Table index ${idx} out of range (${tables.Count} table(s) exist)`, "TABLE_NOT_FOUND", false, "Use word_insert_table first, or check existing table count by word_get_info.")
    }
    return tables.Item(idx)
  }

  private validateCellIndices(table: Record<string, unknown>, row: number, col: number): void {
    const rowCount = (table.Rows as { Count: number }).Count
    const colCount = (table.Columns as { Count: number }).Count
    if (row < 1 || row > rowCount) {
      throw new WordMcpError(
        `Row index ${row} out of range (1-${rowCount})`,
        "TABLE_INDEX_OUT_OF_RANGE", false,
        `The table has ${rowCount} row(s).`,
      )
    }
    if (col < 1 || col > colCount) {
      throw new WordMcpError(
        `Column index ${col} out of range (1-${colCount})`,
        "TABLE_INDEX_OUT_OF_RANGE", false,
        `The table has ${colCount} column(s).`,
      )
    }
  }

  async editCells(params: { tableIndex?: number; data: string[][] }): Promise<{ rows: number; columns: number; failCount: number; truncatedRows: number; truncatedCols: number }> {
    const table = this.requireTable(params)
    const rows = (table.Rows as { Count: number }).Count
    const cols = (table.Columns as { Count: number }).Count
    let failCount = 0
    const dataRows = params.data.length
    const dataCols = params.data.reduce((m, r) => Math.max(m, r.length), 0)
    for (let r = 0; r < dataRows && r < rows; r++) {
      for (let c = 0; c < params.data[r].length && c < cols; c++) {
        try {
          const cellObj = (table.Cell as (r: number, c: number) => Record<string, unknown>)(r + 1, c + 1)
          const cellRange = cellObj.Range as Record<string, unknown>
          cellRange.Text = params.data[r][c]
        } catch { failCount++ }
      }
    }
    return { rows, columns: cols, failCount, truncatedRows: dataRows > rows ? dataRows - rows : 0, truncatedCols: dataCols > cols ? dataCols - cols : 0 }
  }

  async editTableCell(params: EditCellParams): Promise<void> {
    const table = this.requireTable(params)
    this.validateCellIndices(table, params.row, params.column)
    const cellObj = (table.Cell as (r: number, c: number) => Record<string, unknown>)(params.row, params.column)
    const cellRange = cellObj.Range as Record<string, unknown>
    cellRange.Text = params.text
  }

  async addTableRow(params: AddTableRowParams): Promise<{ writtenCells: number; totalCells: number }> {
    const table = this.requireTable(params)
    const row = (table.Rows as { Add: () => Record<string, unknown> }).Add()
    const cells = row.Cells as { Count: number; Item: (i: number) => Record<string, unknown> }
    const totalCells = cells.Count
    let writtenCells = 0
    if (params.data) {
      for (let c = 0; c < params.data.length && c < totalCells; c++) {
        ;(cells.Item(c + 1).Range as Record<string, unknown>).Text = params.data[c]
        writtenCells++
      }
    }
    return { writtenCells, totalCells }
  }

  async deleteTableRow(params: DeleteTableRowParams): Promise<void> {
    const table = this.requireTable(params)
    const rowCount = (table.Rows as { Count: number }).Count
    if (params.rowIndex < 1 || params.rowIndex > rowCount) {
      throw new WordMcpError(
        `Row index ${params.rowIndex} out of range (1-${rowCount})`,
        "TABLE_INDEX_OUT_OF_RANGE", false,
        `The table has ${rowCount} row(s).`,
      )
    }
    ;(table.Rows as { Item: (i: number) => { Delete: () => void } }).Item(params.rowIndex).Delete()
  }

  async setTableBorders(params: {
    tableIndex?: number
    inside?: { style?: unknown; color?: unknown; size?: number }
    outside?: { style?: unknown; color?: unknown; size?: number }
  }): Promise<void> {
    const table = this.requireTable(params)
    const borders = table.Borders as {
      Item: (t: number) => Record<string, unknown>
      InsideLineStyle: number; OutsideLineStyle: number
    }
    const applyBorder = (items: number[], opts: { style?: unknown; color?: unknown; size?: number }) => {
      for (const t of items) {
        try {
          const b = borders.Item(t) as Record<string, unknown>
          if (opts.style != null) b.LineStyle = this.numOrEnum(opts.style, WordTableEditor.BORDER_STYLE)
          if (opts.color != null) b.ColorIndex = this.numOrEnum(opts.color, WordTableEditor.COLOR_INDEX)
          if (opts.size != null) b.LineWidth = opts.size
        } catch { /* individual border may fail (e.g. double style + incompatible width) */ }
      }
    }
    if (params.outside) {
      applyBorder([1, 2, 3, 4], params.outside)
    }
    if (params.inside) {
      applyBorder([5, 6], params.inside)
    }
  }

  async setTableShading(params: { tableIndex?: number; color: string; target?: string; rowIndex?: number }): Promise<void> {
    const table = this.requireTable(params)
    const hex = params.color.replace("#", "")
    const r = parseInt(hex.slice(0, 2), 16)
    const g = parseInt(hex.slice(2, 4), 16)
    const b = parseInt(hex.slice(4, 6), 16)
    const wordColor = r + g * 256 + b * 65536
    if (params.target === "row") {
      const ri = params.rowIndex ?? 1
      const rowCount = (table.Rows as { Count: number }).Count
      if (ri < 1 || ri > rowCount) {
        throw new WordMcpError(
          `Row index ${ri} out of range (1-${rowCount})`,
          "TABLE_INDEX_OUT_OF_RANGE", false,
          `The table has ${rowCount} row(s).`,
        )
      }
      const row = (table.Rows as { Item: (i: number) => Record<string, unknown> }).Item(ri)
      ;((row.Shading as Record<string, unknown>).BackgroundPatternColor as number) = wordColor
    } else {
      ;((table.Shading as Record<string, unknown>).BackgroundPatternColor as number) = wordColor
    }
  }

  async mergeTableCells(params: { tableIndex?: number; rowStart: number; colStart: number; rowEnd: number; colEnd: number }): Promise<void> {
    const table = this.requireTable(params)
    this.validateCellIndices(table, params.rowStart, params.colStart)
    this.validateCellIndices(table, params.rowEnd, params.colEnd)
    if (params.rowStart > params.rowEnd || params.colStart > params.colEnd) {
      throw new WordMcpError(
        "Merge range is invalid: start must be <= end for both row and column",
        "TABLE_MERGE_INVALID_RANGE", false,
        "Ensure rowStart ≤ rowEnd and colStart ≤ colEnd.",
      )
    }
    const cell1 = (table.Cell as (r: number, c: number) => Record<string, unknown>)(params.rowStart, params.colStart)
    const cell2 = (table.Cell as (r: number, c: number) => Record<string, unknown>)(params.rowEnd, params.colEnd)
    ;(cell1.Merge as (t: unknown) => void)(cell2)
  }

  static applyDefaultBorders(table: Record<string, unknown>): void {
    try {
      const borders = table.Borders as { Item: (t: number) => Record<string, unknown> }
      const apply = (items: number[], style: number, color: number, width: number) => {
        for (const t of items) {
          try {
            const b = borders.Item(t)
            b.LineStyle = style
            b.ColorIndex = color
            b.LineWidth = width
          } catch { /* individual border may fail */ }
        }
      }
      apply([1, 2, 3, 4], 1, 1, 4)
      apply([5, 6], 1, 1, 2)
    } catch { /* table may not support borders */ }
  }

  async setColumnWidth(params: { tableIndex?: number; column: number; width: number }): Promise<void> {
    const table = this.requireTable(params)
    const colCount = (table.Columns as { Count: number }).Count
    if (params.column < 1 || params.column > colCount) {
      throw new WordMcpError(
        `Column index ${params.column} out of range (1-${colCount})`,
        "TABLE_INDEX_OUT_OF_RANGE", false,
        `The table has ${colCount} column(s).`,
      )
    }
    const col = (table.Columns as { Item: (i: number) => Record<string, unknown> }).Item(params.column)
    ;(col.Width as number) = params.width
  }

  async setRowHeight(params: { tableIndex?: number; row: number; height: number }): Promise<void> {
    const table = this.requireTable(params)
    const rowCount = (table.Rows as { Count: number }).Count
    if (params.row < 1 || params.row > rowCount) {
      throw new WordMcpError(
        `Row index ${params.row} out of range (1-${rowCount})`,
        "TABLE_INDEX_OUT_OF_RANGE", false,
        `The table has ${rowCount} row(s).`,
      )
    }
    const row = (table.Rows as { Item: (i: number) => Record<string, unknown> }).Item(params.row)
    ;(row.Height as number) = params.height
  }

  async setCellFont(params: {
    tableIndex?: number; row: number; column: number
    name?: string; size?: number; bold?: boolean; italic?: boolean
    underline?: unknown; color?: unknown; strikethrough?: boolean
  }): Promise<void> {
    const table = this.requireTable(params)
    this.validateCellIndices(table, params.row, params.column)
    const cell = (table.Cell as (r: number, c: number) => Record<string, unknown>)(params.row, params.column)
    const font = (cell.Range as Record<string, unknown>).Font as Record<string, unknown>
    if (params.name != null) font.Name = params.name
    if (params.size != null) font.Size = params.size
    if (params.bold != null) font.Bold = params.bold
    if (params.italic != null) font.Italic = params.italic
    if (params.underline != null) font.Underline = this.numOrEnum(params.underline, WordTableEditor.UNDERLINE)
    if (params.color != null) font.ColorIndex = this.numOrEnum(params.color, WordTableEditor.COLOR_INDEX)
    if (params.strikethrough != null) font.Strikethrough = params.strikethrough ? 1 : 0
  }

  async applyTableStyle(params: { tableIndex?: number; styleName: string }): Promise<void> {
    const table = this.requireTable(params)
    const doc = this.requireDoc()
    const styles = doc.Styles as unknown as { Count: number; Item: (i: number) => Record<string, unknown> }

    const candidates: string[] = [params.styleName]

    const fallbackMap: Record<string, string> = {
      "table grid": "网格型",
      "light list accent 1": "浅色列表 - 着色 1",
      "light list accent 2": "浅色列表 - 着色 2",
      "light list accent 3": "浅色列表 - 着色 3",
      "light list accent 4": "浅色列表 - 着色 4",
      "light list accent 5": "浅色列表 - 着色 5",
      "light list accent 6": "浅色列表 - 着色 6",
      "light shading accent 1": "浅色底纹 - 着色 1",
      "medium shading 1 accent 1": "中型底纹 1 - 着色 1",
      "medium list 1 accent 1": "中型列表 1 - 着色 1",
      "dark list accent 1": "深色列表 - 着色 1",
    }
    const lower = params.styleName.toLowerCase()
    if (fallbackMap[lower]) {
      candidates.push(fallbackMap[lower])
    }

    for (const name of candidates) {
      try {
        for (let i = 1; i <= styles.Count; i++) {
          const s = styles.Item(i)
          const localName = (s.NameLocal as string) ?? ""
          if (localName.toLowerCase() === name.toLowerCase()) {
            ;(table.Style as string) = localName
            return
          }
        }
      } catch { /* try next */ }
    }

    try { ;(table.Style as string) = params.styleName } catch { /* ignore */ }
  }

  async addTableColumn(params: { tableIndex?: number; column?: number }): Promise<void> {
    const table = this.requireTable(params)
    if (params.column != null) {
      const col = (table.Columns as { Item: (i: number) => Record<string, unknown> }).Item(params.column)
      ;((col.Columns as { Add: () => void }).Add)()
    } else {
      ;((table.Columns as { Add: () => void }).Add)()
    }
  }

  async deleteTableColumn(params: { tableIndex?: number; column: number }): Promise<void> {
    const table = this.requireTable(params)
    const colCount = (table.Columns as { Count: number }).Count
    if (params.column < 1 || params.column > colCount) {
      throw new WordMcpError(
        `Column index ${params.column} out of range (1-${colCount})`,
        "TABLE_INDEX_OUT_OF_RANGE", false,
        `The table has ${colCount} column(s).`,
      )
    }
    ;((table.Columns as { Item: (i: number) => { Delete: () => void } }).Item(params.column).Delete)()
  }

  async setCellVerticalAlignment(params: {
    tableIndex?: number; row: number; column: number; alignment: unknown
  }): Promise<void> {
    const table = this.requireTable(params)
    this.validateCellIndices(table, params.row, params.column)
    const cell = (table.Cell as (r: number, c: number) => Record<string, unknown>)(params.row, params.column)
    ;(cell.VerticalAlignment as number) = this.numOrEnum(params.alignment, WordTableEditor.V_ALIGN)
  }

  async tableToText(params: { tableIndex?: number; separator?: string }): Promise<string> {
    const table = this.requireTable(params)
    const sep = params.separator ?? "\t"
    ;(table.ConvertToText as (s: string) => void)(sep)
    return "Table converted to text"
  }

  async textToTable(params: { separator?: string; autoFitBehavior?: unknown }): Promise<{ rows: number; columns: number }> {
    this.collapseSelection()
    const doc = this.requireDoc()
    const sel = this.getSelection()
    if (!sel.Range) throw new WordMcpError("No text selected for conversion", "NO_SELECTION", false, "Use word_select_text, word_find_text, or word_select_all first to select the text to convert.")
    const sep = params.separator ?? "\t"
    const table = (sel.Range as Record<string, unknown>).ConvertToTable as (s: string) => Record<string, unknown>
    const t = table(sep)
    if (params.autoFitBehavior != null) {
      ;(t as Record<string, unknown>).AutoFitBehavior = this.numOrEnum(params.autoFitBehavior, WordTableEditor.AUTO_FIT)
    }
    const rows = (t as Record<string, unknown>).Rows as { Count: number }
    const cols = (t as Record<string, unknown>).Columns as { Count: number }
    return { rows: rows.Count, columns: cols.Count }
  }
}
