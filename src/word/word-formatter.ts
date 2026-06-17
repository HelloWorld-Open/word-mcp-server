import { WordBase } from "./word-base.js"
import type { IWordSession } from "./session.js"
import { WordMcpError } from "../security/errors.js"

export type BorderSide = "top" | "bottom" | "left" | "right"

export interface BorderOptions {
  style: "none" | "single" | "dot" | "dash" | "double"
  color?: string
  size?: number
  sides?: BorderSide[]
}

export interface StyleProfileFont {
  name?: string
  size?: number
  bold?: boolean
  italic?: boolean
  color?: string
  underline?: "none" | "single" | "double" | "wavy"
  highlight?: string
  strikethrough?: boolean
}

export interface StyleProfilePara {
  alignment?: string
  firstLineIndent?: number
  spaceBefore?: number
  spaceAfter?: number
  lineSpacing?: number
  lineSpacingRule?: string
  borders?: BorderOptions
  keepWithNext?: boolean
  pageBreakBefore?: boolean
}

export interface StyleProfile {
  font?: StyleProfileFont
  paragraph?: StyleProfilePara
}

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

export class WordFormatter extends WordBase {
  constructor(session: IWordSession) {
    super(session)
  }

  private static readonly ALIGNMENT: Record<string, number> = {
    left: 0, center: 1, right: 2, justify: 3,
  }
  private static readonly ORIENTATION: Record<string, number> = {
    portrait: 0, landscape: 1,
  }
  private static readonly LINE_SPACING_RULE: Record<string, number> = {
    single: 0, one_point_five: 1, double: 2, at_least: 3, exactly: 4, multiple: 5,
  }
  private static readonly UNDERLINE: Record<string, number> = {
    none: 0, single: 1, double: 3, wavy: 11,
  }
  private static readonly COLOR_INDEX: Record<string, number> = {
    auto: 0, black: 1, blue: 2, turquoise: 3, bright_green: 4, pink: 5,
    red: 6, yellow: 7, white: 8, dark_blue: 9, teal: 10, green: 11,
    violet: 12, dark_red: 13, dark_yellow: 14, gray_50: 15, gray_25: 16,
  }
  private static readonly BORDER_LINE_STYLE: Record<string, number> = {
    none: 0, single: 1, dot: 3, dash: 7, double: 8,
  }
  private static readonly BORDER_SIDE_INDEX: Record<string, number> = {
    top: -1, left: -2, bottom: -3, right: -4,
  }

  private static readonly ALIGNMENT_STRUCT: Record<string, number> = {
    left: 0, center: 1, right: 2,
  }
  private static readonly COLOR_RGB: Record<string, number> = {
    auto: 0, black: 0, blue: 0x0000FF, turquoise: 0x00FFFF, bright_green: 0x00FF00, pink: 0xFF00FF,
    red: 0xFF0000, yellow: 0xFFFF00, white: 0xFFFFFF, dark_blue: 0x000080, teal: 0x008080, green: 0x008000,
    violet: 0x800080, dark_red: 0x800000, dark_yellow: 0x808000, gray_50: 0x808080, gray_25: 0xC0C0C0,
  }

  private static readonly CM_TO_POINTS = 28.3465
  private static readonly NORMAL_STYLES = ["Normal", "正文"]

  private cmToPoints(cm: number): number {
    return cm * WordFormatter.CM_TO_POINTS
  }

  private getLastSectionIndex(): number {
    const doc = this.getDocProxy()
    return doc.getSections().count
  }

  private resolveSectionIndex(sectionIndex?: number): number {
    const doc = this.getDocProxy()
    const count = doc.getSections().count
    if (sectionIndex != null) {
      if (sectionIndex < 1 || sectionIndex > count) {
        throw new WordMcpError(
          `Section index ${sectionIndex} out of range (1-${count})`,
          "SECTION_INDEX_OUT_OF_RANGE",
          false,
          `Document has ${count} section(s).`
        )
      }
      return sectionIndex
    }
    return count
  }

  private saveCursorPosition(): number | undefined {
    try {
      return this.getSelProxy().getStart()
    } catch { return undefined }
  }

  private restoreMainDocCursor(restorePos?: number): void {
    try {
      const doc = this.getDocProxy()
      const targetPos = restorePos ?? doc.getContent().getEnd()
      const range = doc.getRange(targetPos, targetPos)
      range.select()
      this.getSelProxy().collapse(0)
      this.cursor.markInBody()
      this.session.lockPrintView()
    } catch { /* ignore */ }
  }

  async setFont(params: Record<string, unknown>): Promise<void> {
    const font = this.getSelProxy().getFont()
    if (params.name != null) font.Name = params.name
    if (params.size != null) font.Size = params.size
    if (params.bold != null) font.Bold = !!params.bold
    if (params.italic != null) font.Italic = !!params.italic
    if (params.underline != null) font.Underline = this.numOrEnum(params.underline, WordFormatter.UNDERLINE)
    if (params.color != null) font.ColorIndex = this.numOrEnum(params.color, WordFormatter.COLOR_INDEX)
    if (params.strikethrough != null) font.Strikethrough = params.strikethrough ? 1 : 0
    if (params.highlightColor != null) font.HighlightColorIndex = this.numOrEnum(params.highlightColor, WordFormatter.COLOR_INDEX)
    if (params.superscript != null) font.Superscript = params.superscript ? 1 : 0
    if (params.subscript != null) font.Subscript = params.subscript ? 1 : 0
  }

  async setParagraphFormat(params: Record<string, unknown>): Promise<void> {
    const pf = this.getSelProxy().getParagraphFormat()
    if (params.alignment != null) pf.Alignment = this.numOrEnum(params.alignment, WordFormatter.ALIGNMENT)
    if (params.leftIndent != null) pf.LeftIndent = this.cmToPoints(params.leftIndent as number)
    if (params.rightIndent != null) pf.RightIndent = this.cmToPoints(params.rightIndent as number)
    if (params.firstLineIndent != null) pf.FirstLineIndent = this.cmToPoints(params.firstLineIndent as number)
    if (params.spaceBefore != null) pf.SpaceBefore = params.spaceBefore
    if (params.spaceAfter != null) pf.SpaceAfter = params.spaceAfter
    if (params.lineSpacing != null) pf.LineSpacing = params.lineSpacing
    if (params.lineSpacingRule != null) pf.LineSpacingRule = this.numOrEnum(params.lineSpacingRule, WordFormatter.LINE_SPACING_RULE)
  }

  async applyStyle(styleName: string): Promise<void> {
    const sel = this.getSelProxy()
    const names = this.styleCandidates(styleName)
    for (const n of names) {
      try { sel.setStyle(n); return } catch { /* try next */ }
    }
    throw new WordMcpError(`Style not found: ${styleName}`, "STYLE_NOT_FOUND", false, `Available styles: try word_list_styles() to see all styles in the current document.`)
  }

  private styleCandidates(name: string): string[] {
    const candidates = [name]
    const cn = name.match(/^标题\s+(\d+)$/)
    if (cn) candidates.push(`Heading ${cn[1]}`)
    const en = name.match(/^(?:Heading|heading)\s+(\d+)$/)
    if (en) candidates.push(`标题 ${en[1]}`)
    return [...new Set(candidates)]
  }

  async modifyStyle(styleName: string, profile: StyleProfile): Promise<void> {
    const doc = this.getDocProxy()
    const styles = doc.getStyles()
    const candidates = this.styleCandidates(styleName)
    let style: Record<string, unknown> | undefined
    for (const n of candidates) {
      try { style = styles.itemByName(n); break } catch { continue }
    }
    if (!style) {
      throw new WordMcpError(
        `Built-in style not found: ${styleName}`,
        "STYLE_NOT_FOUND", false,
        "Try word_list_styles() to see available styles in the current document."
      )
    }

    await this.session.withScreenOff(async () => {
      if (profile.font) {
        const font = style!.Font as Record<string, unknown>
        if (profile.font!.name != null) font.Name = profile.font!.name
        if (profile.font!.size != null) font.Size = profile.font!.size
        if (profile.font!.bold != null) font.Bold = profile.font!.bold
        if (profile.font!.italic != null) font.Italic = profile.font!.italic
        if (profile.font!.color != null) font.ColorIndex = this.numOrEnum(profile.font!.color, WordFormatter.COLOR_INDEX)
        if (profile.font!.underline != null) font.Underline = this.numOrEnum(profile.font!.underline, WordFormatter.UNDERLINE)
        if (profile.font!.strikethrough != null) font.Strikethrough = profile.font!.strikethrough ? 1 : 0
        if (profile.font!.highlight != null) {
          const idx = WordFormatter.COLOR_INDEX[profile.font!.highlight]
          if (idx !== undefined) {
            font.HighlightColorIndex = idx
          } else if (profile.font!.highlight.startsWith("#")) {
            try {
              const rgb = parseInt(profile.font!.highlight.slice(1), 16)
              const bgr = ((rgb & 0xFF) << 16) | (rgb & 0xFF00) | ((rgb >> 16) & 0xFF)
              const shading = font.Shading as Record<string, unknown>
              shading.BackgroundPatternColor = bgr
            } catch { /* ignore invalid hex */ }
          }
        }
      }

      if (profile.paragraph) {
        const pf = style!.ParagraphFormat as Record<string, unknown>
        if (profile.paragraph!.alignment != null) pf.Alignment = this.numOrEnum(profile.paragraph!.alignment, WordFormatter.ALIGNMENT)
        if (profile.paragraph!.firstLineIndent != null) pf.FirstLineIndent = this.cmToPoints(profile.paragraph!.firstLineIndent)
        if (profile.paragraph!.spaceBefore != null) pf.SpaceBefore = profile.paragraph!.spaceBefore
        if (profile.paragraph!.spaceAfter != null) pf.SpaceAfter = profile.paragraph!.spaceAfter
        if (profile.paragraph!.lineSpacing != null) pf.LineSpacing = profile.paragraph!.lineSpacing
        if (profile.paragraph!.lineSpacingRule != null) pf.LineSpacingRule = this.numOrEnum(profile.paragraph!.lineSpacingRule, WordFormatter.LINE_SPACING_RULE)
        if (profile.paragraph!.keepWithNext != null) pf.KeepWithNext = profile.paragraph!.keepWithNext
        if (profile.paragraph!.pageBreakBefore != null) pf.PageBreakBefore = profile.paragraph!.pageBreakBefore
        if (profile.paragraph!.borders) {
          const b = profile.paragraph!.borders
          const sides = b.sides ?? ["top", "bottom", "left", "right"]
          const borders = pf.Borders as { Item: (i: number) => Record<string, unknown> }
          for (const side of sides) {
            const bi = WordFormatter.BORDER_SIDE_INDEX[side]
            if (bi == null) continue
            try {
              const border = borders.Item(bi)
              border.LineStyle = this.numOrEnum(b.style, WordFormatter.BORDER_LINE_STYLE)
              if (b.color != null) border.ColorIndex = this.numOrEnum(b.color, WordFormatter.COLOR_INDEX)
              if (b.size != null) border.LineWidth = b.size
            } catch { /* ignore per-side border failures */ }
          }
        }
      }
    })
  }

  async setPageSetup(params: Record<string, unknown>): Promise<void> {
    const sections = this.getDocProxy().getSections()
    const ps = sections.item(sections.count).getPageSetup()
    if (params.topMargin != null) ps.TopMargin = this.cmToPoints(params.topMargin as number)
    if (params.bottomMargin != null) ps.BottomMargin = this.cmToPoints(params.bottomMargin as number)
    if (params.leftMargin != null) ps.LeftMargin = this.cmToPoints(params.leftMargin as number)
    if (params.rightMargin != null) ps.RightMargin = this.cmToPoints(params.rightMargin as number)
    if (params.orientation != null) ps.Orientation = this.numOrEnum(params.orientation, WordFormatter.ORIENTATION)
    if (params.pageWidth != null) ps.PageWidth = this.cmToPoints(params.pageWidth as number)
    if (params.pageHeight != null) ps.PageHeight = this.cmToPoints(params.pageHeight as number)
  }

  async setDocumentProperties(params: Record<string, unknown>): Promise<void> {
    const doc = this.getDocProxy()
    try {
      const props = doc.getBuiltInDocumentProperties()
      if (params.title != null) (props.itemByName("Title") as Record<string, unknown>).Value = params.title
      if (params.author != null) (props.itemByName("Author") as Record<string, unknown>).Value = params.author
      if (params.subject != null) (props.itemByName("Subject") as Record<string, unknown>).Value = params.subject
      if (params.keywords != null) (props.itemByName("Keywords") as Record<string, unknown>).Value = params.keywords
      if (params.comments != null) (props.itemByName("Comments") as Record<string, unknown>).Value = params.comments
      if (params.category != null) (props.itemByName("Category") as Record<string, unknown>).Value = params.category
    } catch {
    }
  }

  async applyBodyIndent(indentCm: number): Promise<number> {
    const doc = this.getDocProxy()
    const sel = this.getSelProxy()
    const origStart = sel.getStart()
    const origEnd = sel.getEnd()
    let count = 0
    try {
      const content = doc.getContent()
      const rng = doc.getRange(content.getStart(), content.getEnd())
      const find = rng.getFind()
      const replacement = find.Replacement as Record<string, unknown>
      const replPfmt = replacement.ParagraphFormat as Record<string, unknown>
      ;(find.ClearFormatting as () => void)()
      find.Style = "Normal"
      ;(replacement.ClearFormatting as () => void)()
      replPfmt.FirstLineIndent = this.cmToPoints(indentCm)
      replacement.Text = "^&"
      find.Text = ""
      find.Forward = true
      find.Format = true
      find.Wrap = 1
      ;(find.Execute as (...a: unknown[]) => boolean)("", false, false, false, false, false, true, 1, true, "^&", 2)
      count = doc.getParagraphs().count
    } finally {
      try {
        const restore = doc.getRange(origStart, origEnd)
        restore.select()
      } catch { }
    }
    return count
  }

  async setTrackChanges(enable: boolean): Promise<void> {
    const doc = this.getDocProxy()
    doc.setTrackRevisions(enable)
  }

  resetParagraphStyle(): void {
    try {
      const sel = this.getSelProxy()
      for (const name of WordFormatter.NORMAL_STYLES) {
        try { sel.setStyle(name); return } catch { /* try next */ }
      }
    } catch { /* ignore */ }
  }

  async acceptAllChanges(): Promise<number> {
    const doc = this.getDocProxy()
    const revisions = doc.getRevisions()
    const count = revisions.count
    if (count > 0) revisions.acceptAll()
    return count
  }

  async rejectAllChanges(): Promise<number> {
    const doc = this.getDocProxy()
    const revisions = doc.getRevisions()
    const count = revisions.count
    if (count > 0) revisions.rejectAll()
    return count
  }

  async listStyles(): Promise<Array<{ name: string; type: number; builtIn: boolean }>> {
    const doc = this.getDocProxy()
    const styles = doc.getStyles()
    const result: Array<{ name: string; type: number; builtIn: boolean }> = []
    for (let i = 1; i <= styles.count; i++) {
      const style = styles.item(i)
      if (style.InUse) {
        result.push({
          name: style.NameLocal as string,
          type: style.Type as number,
          builtIn: style.BuiltIn as boolean,
        })
      }
    }
    return result
  }

  async setHeader(text: string, alignment?: unknown, sectionIndex?: number): Promise<void> {
    const savedCursorPos = this.saveCursorPosition()
    const si = this.resolveSectionIndex(sectionIndex)
    const hdr = this.getDocProxy().getSections().item(si).getHeader(1)
    hdr.setContent(text, alignment != null ? this.numOrEnum(alignment, WordFormatter.ALIGNMENT_STRUCT) as number : undefined)
    this.restoreMainDocCursor(savedCursorPos)
  }

  async setFooter(text: string, alignment?: unknown, sectionIndex?: number): Promise<void> {
    const savedCursorPos = this.saveCursorPosition()
    const si = this.resolveSectionIndex(sectionIndex)
    const ftr = this.getDocProxy().getSections().item(si).getFooter(1)
    ftr.setContent(text, alignment != null ? this.numOrEnum(alignment, WordFormatter.ALIGNMENT_STRUCT) as number : undefined)
    this.restoreMainDocCursor(savedCursorPos)
  }

  async setPageNumbers(target: "header" | "footer", alignment?: unknown, sectionIndex?: number): Promise<void> {
    const savedCursorPos = this.saveCursorPosition()
    const si = this.resolveSectionIndex(sectionIndex)
    const hf = target === "header"
      ? this.getDocProxy().getSections().item(si).getHeader(1)
      : this.getDocProxy().getSections().item(si).getFooter(1)

    // 1. 倒序遍历删除已有 PAGE 域（wdFieldPage = 33）
    const fields = hf.getFields()
    for (let i = fields.count; i >= 1; i--) {
      if (fields.item(i).type === 33) {
        fields.item(i).delete()
      }
    }

    // 2. 快照：读取当前状态为纯 JS 数据，避免后续 COM 可变性导致的不一致
    const pre = { endPos: hf.getEnd(), isEmpty: hf.getText().trim().length === 0 }
    const insertPos = pre.isEmpty ? pre.endPos : pre.endPos - 1

    // 3. 进入页眉/页脚编辑模式并在预计算位置设置插入点
    hf.select()
    const rng = hf.getRange()
    rng.setRange(insertPos, insertPos)
    rng.select()

    const sel = this.getSelProxy()
    if (!pre.isEmpty) {
      sel.typeText(" ")
    }
    sel.getRange().addField(33)

    // 4. 设置段落对齐和页码对齐
    const alignVal = alignment != null
      ? this.numOrEnum(alignment, WordFormatter.ALIGNMENT_STRUCT)
      : 1
    hf.setAlignment(alignVal as number)
    hf.setPageNumbersAlignment(alignVal as number)

    this.restoreMainDocCursor(savedCursorPos)
  }

  async insertToc(): Promise<void> {
    this.collapseSelection()
    const doc = this.getDocProxy()
    ;((doc.getTablesOfContents() as { Add: (r: unknown) => void }).Add(this.getSelProxy().getRange().raw))
    const sel = this.getSelProxy()
    this.goToEnd()
    this.getSelProxy().collapse(0)
    sel.typeParagraph()
  }

  async addBookmark(name: string): Promise<void> {
    this.collapseSelection()
    this.getDocProxy().getBookmarks().add(name)
    this.goToEnd()
    this.getSelProxy().collapse(0)
  }

  async addComment(text: string): Promise<void> {
    this.collapseSelection()
    this.getDocProxy().getComments().add(this.getSelProxy().getRange().raw, text)
    this.goToEnd()
    this.getSelProxy().collapse(0)
  }

  async setWatermark(params: { text: string; remove?: boolean; fontSize?: number; color?: unknown }): Promise<void> {
    const savedCursorPos = this.saveCursorPosition()
    const sections = this.getDocProxy().getSections()
    await this.session.withScreenOff(async () => {
      const addToSection = (si: number) => {
        const hdr = sections.item(si).getHeader(1).raw
        if (params.remove) {
          const shapes = hdr.Shapes as { Count: number; Item: (i: number) => Record<string, unknown> }
          for (let i = shapes.Count; i >= 1; i--) { ;(shapes.Item(i).Delete as () => void)() }
          return
        }
        const shape = (hdr.Shapes as { AddTextEffect: (preset: number, text: string, font: string, size: number, bold: boolean, italic: boolean, left: number, top: number) => Record<string, unknown> })
          .AddTextEffect(0, params.text, "Arial", params.fontSize ?? 48, false, false, 0, 0)
        shape.RelativeHorizontalPosition = 0
        shape.RelativeVerticalPosition = 0
        shape.Left = -999995
        shape.Top = -999995
        shape.Rotation = -45
        ;(shape.Fill as Record<string, unknown>).Visible = true
        ;(shape.Line as Record<string, unknown>).Visible = false
        if (params.color != null) {
          const cf = (shape.Fill as Record<string, unknown>).ForeColor as Record<string, unknown>
          cf.RGB = this.numOrEnum(params.color, WordFormatter.COLOR_RGB)
        }
        ;((shape.WrapFormat as Record<string, unknown>).AllowOverlap as boolean) = true
        ;(shape.ZOrder as (a: number) => void)(4)
      }
      if (!params.remove) {
        for (let i = 1; i <= sections.count; i++) addToSection(i)
      } else {
        if (sections.count > 0) addToSection(1)
      }
    })
    this.restoreMainDocCursor(savedCursorPos)
  }

  private static readonly V_ALIGN: Record<string, number> = {
    top: 0, center: 1, bottom: 3,
  }
  private static readonly BORDER_STYLE: Record<string, number> = {
    none: 0, single: 1, dot: 2, dash_small: 3, dash_large: 4, dash: 5, dash_dot: 6, double: 7,
  }

  private requireTable(params: { tableIndex?: number }): Record<string, unknown> {
    const doc = this.getDocProxy()
    const tables = doc.getTables()
    const idx = (params.tableIndex ?? 1)
    if (idx < 1 || idx > tables.count) {
      throw new WordMcpError(`Table index ${idx} out of range (${tables.count} table(s) exist)`,
        "TABLE_NOT_FOUND", false, "Use word_insert_table first, or check existing table count by word_get_info.")
    }
    return tables.item(idx)
  }

  private validateCellIndices(table: Record<string, unknown>, row: number, col: number): void {
    const rowCount = (table.Rows as { Count: number }).Count
    const colCount = (table.Columns as { Count: number }).Count
    if (row < 1 || row > rowCount) {
      throw new WordMcpError(`Row index ${row} out of range (1-${rowCount})`, "TABLE_INDEX_OUT_OF_RANGE", false,
        `The table has ${rowCount} row(s).`)
    }
    if (col < 1 || col > colCount) {
      throw new WordMcpError(`Column index ${col} out of range (1-${colCount})`, "TABLE_INDEX_OUT_OF_RANGE", false,
        `The table has ${colCount} column(s).`)
    }
  }

  async editCells(params: { tableIndex?: number; data: string[][] }): Promise<{ rows: number; columns: number; failCount: number; truncatedRows: number; truncatedCols: number }> {
    const table = this.requireTable(params)
    const rows = (table.Rows as { Count: number }).Count
    const cols = (table.Columns as { Count: number }).Count
    let failCount = 0
    const dataRows = params.data.length
    const dataCols = params.data.reduce((m, r) => Math.max(m, r.length), 0)
    const TIME_BUDGET = 50
    let batchStart = Date.now()
    await this.session.withScreenOff(async () => {
      for (let r = 0; r < dataRows && r < rows; r++) {
        for (let c = 0; c < params.data[r].length && c < cols; c++) {
          try {
            const cellObj = (table.Cell as (r: number, c: number) => Record<string, unknown>)(r + 1, c + 1)
            const cellRange = cellObj.Range as Record<string, unknown>
            cellRange.Text = params.data[r][c]
          } catch { failCount++ }
        }
        if (Date.now() - batchStart >= TIME_BUDGET) {
          try { ;(this.getWord() as Record<string, unknown>).ScreenUpdating = true } catch { /* ignore */ }
          try { ;((this.getWord() as Record<string, unknown>).ScreenRefresh as () => void)() } catch { /* ignore */ }
          await new Promise(resolve => setImmediate(resolve))
          try { ;(this.getWord() as Record<string, unknown>).ScreenUpdating = false } catch { /* ignore */ }
          batchStart = Date.now()
        }
      }
    })
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
      throw new WordMcpError(`Row index ${params.rowIndex} out of range (1-${rowCount})`, "TABLE_INDEX_OUT_OF_RANGE", false,
        `The table has ${rowCount} row(s).`)
    }
    ;(table.Rows as { Item: (i: number) => { Delete: () => void } }).Item(params.rowIndex).Delete()
  }

  async setTableBorders(params: { tableIndex?: number; inside?: { style?: unknown; color?: unknown; size?: number }; outside?: { style?: unknown; color?: unknown; size?: number } }): Promise<void> {
    const table = this.requireTable(params)
    const borders = table.Borders as { Item: (t: number) => Record<string, unknown>; InsideLineStyle: number; OutsideLineStyle: number }
    const applyBorder = (items: number[], opts: { style?: unknown; color?: unknown; size?: number }) => {
      for (const t of items) {
        try {
          const b = borders.Item(t) as Record<string, unknown>
          if (opts.style != null) b.LineStyle = this.numOrEnum(opts.style, WordFormatter.BORDER_STYLE)
          if (opts.color != null) b.ColorIndex = this.numOrEnum(opts.color, WordFormatter.COLOR_INDEX)
          if (opts.size != null) b.LineWidth = opts.size
        } catch { /* ignore per-border failure */ }
      }
    }
    if (params.outside) applyBorder([1, 2, 3, 4], params.outside)
    if (params.inside) applyBorder([5, 6], params.inside)
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
        throw new WordMcpError(`Row index ${ri} out of range (1-${rowCount})`, "TABLE_INDEX_OUT_OF_RANGE", false,
          `The table has ${rowCount} row(s).`)
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
      throw new WordMcpError("Merge range is invalid: start must be <= end for both row and column", "TABLE_MERGE_INVALID_RANGE", false,
        "Ensure rowStart ≤ rowEnd and colStart ≤ colEnd.")
    }
    const cell1 = (table.Cell as (r: number, c: number) => Record<string, unknown>)(params.rowStart, params.colStart)
    const cell2 = (table.Cell as (r: number, c: number) => Record<string, unknown>)(params.rowEnd, params.colEnd)
    ;(cell1.Merge as (t: unknown) => void)(cell2)
  }

  async setColumnWidth(params: { tableIndex?: number; column: number; width: number }): Promise<void> {
    const table = this.requireTable(params)
    const colCount = (table.Columns as { Count: number }).Count
    if (params.column < 1 || params.column > colCount) {
      throw new WordMcpError(`Column index ${params.column} out of range (1-${colCount})`, "TABLE_INDEX_OUT_OF_RANGE", false,
        `The table has ${colCount} column(s).`)
    }
    const col = (table.Columns as { Item: (i: number) => Record<string, unknown> }).Item(params.column)
    ;(col.Width as number) = params.width
  }

  async setRowHeight(params: { tableIndex?: number; row: number; height: number }): Promise<void> {
    const table = this.requireTable(params)
    const rowCount = (table.Rows as { Count: number }).Count
    if (params.row < 1 || params.row > rowCount) {
      throw new WordMcpError(`Row index ${params.row} out of range (1-${rowCount})`, "TABLE_INDEX_OUT_OF_RANGE", false,
        `The table has ${rowCount} row(s).`)
    }
    const row = (table.Rows as { Item: (i: number) => Record<string, unknown> }).Item(params.row)
    ;(row.Height as number) = params.height
  }

  async setCellFont(params: { tableIndex?: number; row: number; column: number; name?: string; size?: number; bold?: boolean; italic?: boolean; underline?: unknown; color?: unknown; strikethrough?: boolean }): Promise<void> {
    const table = this.requireTable(params)
    this.validateCellIndices(table, params.row, params.column)
    const cell = (table.Cell as (r: number, c: number) => Record<string, unknown>)(params.row, params.column)
    const font = (cell.Range as Record<string, unknown>).Font as Record<string, unknown>
    if (params.name != null) font.Name = params.name
    if (params.size != null) font.Size = params.size
    if (params.bold != null) font.Bold = params.bold
    if (params.italic != null) font.Italic = params.italic
    if (params.underline != null) font.Underline = this.numOrEnum(params.underline, WordFormatter.UNDERLINE)
    if (params.color != null) font.ColorIndex = this.numOrEnum(params.color, WordFormatter.COLOR_INDEX)
    if (params.strikethrough != null) font.Strikethrough = params.strikethrough ? 1 : 0
  }

  async applyTableStyle(params: { tableIndex?: number; styleName: string }): Promise<void> {
    const table = this.requireTable(params)
    const doc = this.getDocProxy()
    const styles = doc.getStyles()
    const candidates: string[] = [params.styleName]
    const fallbackMap: Record<string, string> = {
      "table grid": "网格型", "light list accent 1": "浅色列表 - 着色 1",
      "light list accent 2": "浅色列表 - 着色 2", "light list accent 3": "浅色列表 - 着色 3",
      "light list accent 4": "浅色列表 - 着色 4", "light list accent 5": "浅色列表 - 着色 5",
      "light list accent 6": "浅色列表 - 着色 6", "light shading accent 1": "浅色底纹 - 着色 1",
      "medium shading 1 accent 1": "中型底纹 1 - 着色 1", "medium list 1 accent 1": "中型列表 1 - 着色 1",
      "dark list accent 1": "深色列表 - 着色 1",
    }
    const lower = params.styleName.toLowerCase()
    if (fallbackMap[lower]) candidates.push(fallbackMap[lower])
    for (const name of candidates) {
      try {
        for (let i = 1; i <= styles.count; i++) {
          const s = styles.item(i)
          const localName = (s.NameLocal as string) ?? ""
          if (localName.toLowerCase() === name.toLowerCase()) { ;(table.Style as string) = localName; return }
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
      throw new WordMcpError(`Column index ${params.column} out of range (1-${colCount})`, "TABLE_INDEX_OUT_OF_RANGE", false,
        `The table has ${colCount} column(s).`)
    }
    ;((table.Columns as { Item: (i: number) => { Delete: () => void } }).Item(params.column).Delete)()
  }

  async setCellVerticalAlignment(params: { tableIndex?: number; row: number; column: number; alignment: unknown }): Promise<void> {
    const table = this.requireTable(params)
    this.validateCellIndices(table, params.row, params.column)
    const cell = (table.Cell as (r: number, c: number) => Record<string, unknown>)(params.row, params.column)
    ;(cell.VerticalAlignment as number) = this.numOrEnum(params.alignment, WordFormatter.V_ALIGN)
  }

  async tableToText(params: { tableIndex?: number; separator?: string }): Promise<string> {
    const table = this.requireTable(params)
    ;(table.ConvertToText as (s: string) => void)(params.separator ?? "\t")
    return "Table converted to text"
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
}
