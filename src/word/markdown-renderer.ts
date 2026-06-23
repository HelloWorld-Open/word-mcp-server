import type { Block, InlineSegment } from "./markdown-parser.js"
import { parseInline } from "./markdown-parser.js"
import { WordFormatter } from "./word-formatter.js"
import { ContextSanitizer } from "./context-sanitizer.js"
import type { ISelectionProxy, IDocumentProxy } from "./com-proxy/types.js"

const sanitizeText = ContextSanitizer.sanitizeText

export interface RenderContext {
  getSelection(): ISelectionProxy
  requireDoc(): IDocumentProxy
  goToEnd(): void
  withScreenOff<T>(fn: () => Promise<T>): Promise<T>
}

export class MarkdownRenderer {
  constructor(private com: RenderContext) {}

  async renderBlock(
    block: Block,
    bi: number,
    totalBlocks: number,
    pos: "end" | "cursor" = "end",
    skipEnd?: boolean,
  ): Promise<number> {
    const atCursor = pos === "cursor"
    const shouldEnd = !atCursor && !skipEnd
    if (block.type === "hr") {
      if (shouldEnd) this.com.goToEnd()
      this.com.getSelection().addHorizontalLine()
      this.com.getSelection().typeParagraph()
      return 0
    }

    if (block.type === "pagebreak") {
      if (shouldEnd) this.com.goToEnd()
      this.com.getSelection().insertBreak(7)
      return 0
    }

    if (block.type === "image") {
      if (shouldEnd) this.com.goToEnd()
      const doc = this.com.requireDoc()
      try {
        doc.getInlineShapes().addPicture(block.url)
        this.com.getSelection().typeParagraph()
      } catch {
        this.com.getSelection().typeText(sanitizeText(`[图片: ${block.alt}]`))
      }
      return block.alt.length + block.url.length
    }

    if (block.type === "heading" || block.type === "paragraph") {
      const doc = this.com.requireDoc()
      let basePos = -1
      if (atCursor) {
        const sel = this.com.getSelection()
        try { basePos = sel.getRange().getStart() } catch { /* ignore */ }
      } else {
        try { basePos = doc.getContent().getEnd() } catch { /* ignore */ }
        if (shouldEnd) this.com.goToEnd()
      }
      const sel = this.com.getSelection()
      const segs = parseInline(block.text)
      const fullText = segs.map(s => s.text).join("")

      if (basePos >= 0) {
        sel.typeText(sanitizeText(fullText))
        if (block.type === "heading") {
          sel.moveStart(1, -fullText.length)
          MarkdownRenderer.applyHeadingStyle(sel.raw, Math.min(block.level, 9))
          sel.collapse(0)
        }
        let offset = 0
        for (const seg of segs) {
          const segLen = seg.text.length
          const needsFormat = seg.bold || seg.italic || seg.code || seg.strikethrough || seg.link
          if (needsFormat) {
            const range = doc.getRange(basePos + offset, basePos + offset + segLen)
            if (seg.link) {
              try { ;(range.getHyperlinks() as { Add: (r: Record<string, unknown>, a: string) => void }).Add(range.raw, seg.link) } catch { /* ignore */ }
            }
            if (seg.bold || seg.italic || seg.strikethrough || seg.code) {
              const rangeFont = range.getFont()
              if (seg.bold) try { rangeFont.Bold = true } catch { /* ignore */ }
              if (seg.italic) try { rangeFont.Italic = true } catch { /* ignore */ }
              if (seg.strikethrough) try { rangeFont.Strikethrough = true } catch { /* ignore */ }
              if (seg.code) {
                try { rangeFont.Name = "Consolas" } catch { /* ignore */ }
                try { rangeFont.Size = 10.5 } catch { /* ignore */ }
                try { ;(range.getShading() as Record<string, unknown>).BackgroundPatternColor = 0xF0F0F0 } catch { /* ignore */ }
              }
            }
          }
          offset += segLen
        }
      } else {
        for (const seg of segs) {
          const font = sel.getFont()
          let changedBold = false, changedItalic = false, changedStrikethrough = false
          if (seg.bold) { try { font.Bold = true; changedBold = true } catch { /* ignore */ } }
          if (seg.italic) { try { font.Italic = true; changedItalic = true } catch { /* ignore */ } }
          if (seg.strikethrough) { try { font.Strikethrough = true; changedStrikethrough = true } catch { /* ignore */ } }
          if (seg.code) {
            try { font.Name = "Consolas" } catch { /* ignore */ }
            try { font.Size = 10.5 } catch { /* ignore */ }
          }
          if (seg.link) {
            const cleaned = sanitizeText(seg.text)
            sel.typeText(cleaned)
            try {
              sel.moveStart(1, -cleaned.length)
              const range = sel.getRange()
              ;(range.getHyperlinks() as { Add: (r: Record<string, unknown>, a: string) => void }).Add(range.raw, seg.link)
              const linkFont = range.getFont()
              if (seg.bold) try { linkFont.Bold = true } catch { /* ignore */ }
              if (seg.italic) try { linkFont.Italic = true } catch { /* ignore */ }
              if (seg.strikethrough) try { linkFont.Strikethrough = true } catch { /* ignore */ }
              sel.collapse(0)
            } catch { /* ignore */ }
          } else {
            sel.typeText(sanitizeText(seg.text))
            if (seg.code) {
              try {
                const cleaned = sanitizeText(seg.text)
                sel.moveStart(1, -cleaned.length)
                ;(sel.getRange().getShading() as Record<string, unknown>).BackgroundPatternColor = 0xF0F0F0
                sel.collapse(0)
              } catch { /* ignore */ }
            }
          }
          if (changedBold) try { font.Bold = false } catch { /* ignore */ }
          if (changedItalic) try { font.Italic = false } catch { /* ignore */ }
          if (changedStrikethrough) try { font.Strikethrough = false } catch { /* ignore */ }
        }
        if (block.type === "heading") {
          sel.moveStart(1, -fullText.length)
          MarkdownRenderer.applyHeadingStyle(sel.raw, Math.min(block.level, 9))
          sel.collapse(0)
        }
      }
      if (bi < totalBlocks - 1) { sel.typeParagraph() }
      return block.text.length
    }

    if (block.type === "bullet_list" || block.type === "numbered_list") {
      if (shouldEnd) this.com.goToEnd()
      const sel = this.com.getSelection()
      const lf = sel.getRange().getListFormat()
      if (block.type === "bullet_list") {
        ;(lf.ApplyBulletDefault as () => void)()
      } else {
        ;(lf.ApplyNumberDefault as () => void)()
      }
      for (let idx = 0; idx < block.items.length; idx++) {
        for (let indent = 0; indent < block.items[idx].indent; indent++) {
          try { ;(lf.IncreaseIndent as () => void)() } catch { /* IncreaseIndent may not be available */ }
        }
        const segs = parseInline(block.items[idx].text)
        for (const seg of segs) MarkdownRenderer.typeSeg(sel.raw, seg)
        if (idx < block.items.length - 1) { sel.typeParagraph() }
      }
      sel.typeParagraph()
      const freshLf = sel.getRange().getListFormat()
      ;(freshLf.RemoveNumbers as () => void)()
      return block.items.reduce((s, item) => s + item.text.length, 0)
    }

    if (block.type === "table") {
      if (shouldEnd) this.com.goToEnd()
      const doc = this.com.requireDoc()
      const rows = block.rows.length
      const cols = Math.max(...block.rows.map(r => r.length), 1)
      if (rows === 0 || cols === 0) return 0
      const range = this.com.getSelection().getRange().raw
      const tables = doc.getTables()
      const table = tables.add(range, rows, cols)
      // Wrap all cell-fill + styling inside withScreenOff to avoid per-call UI repaints
      await this.com.withScreenOff(async () => {
        for (let r = 0; r < rows; r++) {
          for (let c = 0; c < cols; c++) {
            try {
              const text = sanitizeText(block.rows[r][c] ?? "")
              ;((table.Cell as (r: number, c: number) => { Range: { Text: string } })(r + 1, c + 1).Range.Text as string) = text
            } catch { /* cell may fail individually */ }
          }
        }
        WordFormatter.applyDefaultBorders(table)
        try { ;((table.Rows as Record<string, unknown>).Alignment as number) = 1 } catch { /* ignore */ }
        try {
          const headerRow = (table.Rows as { Item: (i: number) => Record<string, unknown> }).Item(1)
          ;((headerRow.Range as Record<string, unknown>).Font as Record<string, unknown>).Bold = true
          ;(headerRow.Shading as Record<string, unknown>).BackgroundPatternColor = 0xD9E2F3
        } catch { /* ignore */ }
      })
      // AutoFitBehavior after screen re-enabled — single repaint only
      try { ;(table.AutoFitBehavior as (b: number) => void)(1) } catch { /* ignore */ }
      // Move cursor reliably past the table with explicit wdParagraph unit (4)
      try {
        const rawTableRange = table.Range as Record<string, unknown>
        const nextRange = (rawTableRange.Next as (unit?: number) => Record<string, unknown> | undefined)(4)
        if (nextRange) {
          const nextStart = nextRange.Start as number
          if (typeof nextStart === "number") {
            doc.getRange(nextStart, nextStart).select()
            this.com.getSelection().collapse(0)
          } else {
            throw new Error("no next range start")
          }
        } else {
          throw new Error("no content after table")
        }
      } catch {
        // Table is at end of document — InsertParagraph creates a new paragraph outside the table
        try {
          const contentEnd = doc.getContent().getEnd()
          const insertAt = doc.getRange(contentEnd, contentEnd).raw
          ;(insertAt.InsertParagraph as () => void)()
          const newEnd = doc.getContent().getEnd()
          doc.getRange(newEnd, newEnd).select()
          this.com.getSelection().collapse(0)
        } catch { /* ignore */ }
      }
      if (bi < totalBlocks - 1) { this.com.getSelection().typeParagraph() }
      return block.rows.reduce((s, r) => s + r.reduce((a, t) => a + t.length, 0), 0)
    }

    if (block.type === "blockquote") {
      if (shouldEnd) this.com.goToEnd()
      const sel = this.com.getSelection()
      const pf = sel.getParagraphFormat()
      const prevIndent = (pf.LeftIndent as number) ?? 0
      ;(pf.LeftIndent as number) = prevIndent + 28.35
      ;(sel.getFont() as Record<string, unknown>).Italic = true
      const segs = parseInline(block.text)
      const fullText = segs.map(s => s.text).join("")
      for (const seg of segs) MarkdownRenderer.typeSeg(sel.raw, seg)
      ;(sel.getFont() as Record<string, unknown>).Italic = false
      try {
        sel.moveStart(1, -fullText.length)
        const rng = sel.getRange().raw
        ;(rng.Shading as Record<string, unknown>).BackgroundPatternColor = 0xF5F5F5
        const borders = rng.Borders as { Item: (i: number) => Record<string, unknown> }
        const b = borders.Item(1)
        b.LineStyle = 1
        b.ColorIndex = 15
        b.LineWidth = 8
        sel.collapse(0)
      } catch { /* ignore */ }
      ;(pf.LeftIndent as number) = prevIndent
      if (bi < totalBlocks - 1) { sel.typeParagraph() }
      return block.text.length
    }

    if (block.type === "codeblock") {
      if (shouldEnd) this.com.goToEnd()
      const sel = this.com.getSelection()
      const doc = this.com.requireDoc()
      const codeLines = block.text.split("\n")
      const startPos = doc.getContent().getEnd()

      for (let li = 0; li < codeLines.length; li++) {
        sel.typeText(sanitizeText(codeLines[li]))
        if (li < codeLines.length - 1) {
          sel.typeParagraph()
        }
      }

      const endPos = doc.getContent().getEnd()
      try {
        const codeRange = doc.getRange(startPos, endPos)
        ;(codeRange.getFont() as Record<string, unknown>).Name = "Consolas"
        ;(codeRange.getFont() as Record<string, unknown>).Size = 10.5
        ;(codeRange.getShading() as Record<string, unknown>).BackgroundPatternColor = 0xF5F5F5
      } catch { /* ignore */ }

      try {
        const endRange = doc.getRange(endPos, endPos)
        endRange.select()
      } catch { /* ignore */ }
      if (bi < totalBlocks - 1) { this.com.getSelection().typeParagraph() }
      return block.text.length
    }

    return 0
  }

  static blockCost(block: Block): number {
    switch (block.type) {
      case "heading":
      case "paragraph":
      case "blockquote":
        return 1
      case "bullet_list":
      case "numbered_list":
        return Math.max(1, Math.round(block.items.reduce((s, i) => s + 1 + i.indent, 0) * 0.8))
      case "table":
        return Math.max(1, 3 + Math.round(block.rows.length * (block.rows[0]?.length ?? 1) * 0.3))
      case "codeblock":
        return Math.max(1, Math.round(block.text.split("\n").length * 0.5))
      case "hr":
      case "pagebreak":
        return 1
      case "image":
        return 2
      default:
        return 1
    }
  }

  static applyHeadingStyle(sel: Record<string, unknown>, level: number): void {
    const candidates = [`Heading ${level}`, `标题 ${level}`]
    for (const name of candidates) {
      try { sel.Style = name; return } catch { /* try next */ }
    }
  }

  static typeSeg(sel: Record<string, unknown>, seg: InlineSegment): void {
    const font = sel.Font as Record<string, unknown>
    try { font.Bold = seg.bold } catch { /* ignore */ }
    try { font.Italic = seg.italic } catch { /* ignore */ }
    try { font.Strikethrough = seg.strikethrough } catch { /* ignore */ }
    if (seg.code) {
      try { font.Name = "Consolas" } catch { /* ignore */ }
      try { font.Size = 10.5 } catch { /* ignore */ }
    }

    if (seg.link) {
      const cleaned = sanitizeText(seg.text)
      ;(sel.TypeText as (t: string) => void)(cleaned)
      try {
        ;(sel.MoveStart as (unit: number, count: number) => void)(1, -cleaned.length)
        const range = sel.Range as Record<string, unknown>
        ;(range.Hyperlinks as { Add: (r: unknown, a: string) => void }).Add(range, seg.link)
        const linkFont = range.Font as Record<string, unknown>
        if (seg.bold) try { linkFont.Bold = true } catch { /* ignore */ }
        if (seg.italic) try { linkFont.Italic = true } catch { /* ignore */ }
        if (seg.strikethrough) try { linkFont.Strikethrough = true } catch { /* ignore */ }
        ;(sel.Collapse as (d: number) => void)(0)
      } catch { /* ignore */ }
    } else {
      ;(sel.TypeText as (t: string) => void)(sanitizeText(seg.text))
      if (seg.code) {
        try {
          const cleaned = sanitizeText(seg.text)
          ;(sel.MoveStart as (unit: number, count: number) => void)(1, -cleaned.length)
          ;((sel.Range as Record<string, unknown>).Shading as Record<string, unknown>).BackgroundPatternColor = 0xF0F0F0
          ;(sel.Collapse as (d: number) => void)(0)
        } catch { /* ignore */ }
      }
    }

    if (seg.code) {
      const prevName = font.Name as string
      const prevSize = font.Size as number
      if (prevName && prevName !== "Consolas") try { font.Name = prevName } catch { /* ignore */ }
      if (prevSize && prevSize !== 10.5) try { font.Size = prevSize } catch { /* ignore */ }
    }
  }
}
