import type { IWordSession } from "./session.js"
import { WordBase } from "./word-base.js"
import { WordTableEditor } from "./word-table-editor.js"

function sanitizeText(text: string): string {
  return text.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, "")
}

interface InlineSegment {
  text: string
  bold: boolean
  italic: boolean
  code: boolean
  strikethrough: boolean
  link?: string
}

interface ListItem {
  text: string
  indent: number
}

type Block =
  | { type: "heading"; level: number; text: string }
  | { type: "paragraph"; text: string }
  | { type: "bullet_list"; items: ListItem[] }
  | { type: "numbered_list"; items: ListItem[] }
  | { type: "table"; rows: string[][] }
  | { type: "hr" }
  | { type: "pagebreak" }
  | { type: "image"; alt: string; url: string }
  | { type: "blockquote"; text: string }
  | { type: "codeblock"; text: string }

export class WordMarkdown extends WordBase {
  constructor(session: IWordSession) { super(session) }

  async writeBlocks(markdown: string): Promise<{ blocks: number; chars: number }> {
    const blocks = this.parseBlocks(markdown)
    if (blocks.length === 0) return { blocks: 0, chars: 0 }
    let totalChars = 0
    this.collapseSelection()
    for (let bi = 0; bi < blocks.length; bi++) {
      totalChars += this.renderBlock(blocks[bi], bi, blocks.length, "end")
    }
    try {
      ;(this.getSelection().TypeParagraph as () => void)()
    } catch { /* ignore */ }
    return { blocks: blocks.length, chars: totalChars }
  }

  async insertAtCursor(markdown: string): Promise<{ blocks: number; chars: number }> {
    const blocks = this.parseBlocks(markdown)
    if (blocks.length === 0) return { blocks: 0, chars: 0 }
    let totalChars = 0
    this.collapseSelection()
    for (let bi = 0; bi < blocks.length; bi++) {
      totalChars += this.renderBlock(blocks[bi], bi, blocks.length, "cursor")
    }
    return { blocks: blocks.length, chars: totalChars }
  }

  private renderBlock(block: Block, bi: number, totalBlocks: number, pos: "end" | "cursor" = "end"): number {
    const atCursor = pos === "cursor"
    if (block.type === "hr") {
      if (!atCursor) this.goToEnd()
      ;((this.getSelection().InlineShapes as Record<string, unknown>).AddHorizontalLineStandard as () => void)()
      ;(this.getSelection().TypeParagraph as () => void)()
      return 0
    }

    if (block.type === "pagebreak") {
      if (!atCursor) this.goToEnd()
      ;(this.getSelection().InsertBreak as (t: number) => void)(7)
      return 0
    }

    if (block.type === "image") {
      if (!atCursor) this.goToEnd()
      const doc = this.requireDoc()
      const inlineShapes = doc.InlineShapes as { AddPicture: (p: string) => Record<string, unknown> }
      try {
        inlineShapes.AddPicture(block.url)
        ;(this.getSelection().TypeParagraph as () => void)()
      } catch {
        ;(this.getSelection().TypeText as (t: string) => void)(sanitizeText(`[图片: ${block.alt}]`))
      }
      return block.alt.length + block.url.length
    }

    if (block.type === "heading" || block.type === "paragraph") {
      const doc = this.requireDoc()
      let basePos = -1
      if (atCursor) {
        const sel = this.getSelection()
        try { basePos = (sel.Range as Record<string, unknown>).Start as number } catch { /* ignore */ }
      } else {
        try { basePos = (doc.Content as Record<string, unknown>).End as number } catch { /* ignore */ }
        this.goToEnd()
      }
      const sel = this.getSelection()
      const segs = this.parseInline(block.text)
      const fullText = segs.map(s => s.text).join("")

      if (basePos >= 0) {
        ;(sel.TypeText as (t: string) => void)(sanitizeText(fullText))
        if (block.type === "heading") {
          ;(sel.MoveStart as (u: number, c: number) => void)(1, -fullText.length)
          this.applyHeadingStyle(sel, Math.min(block.level, 9))
          ;(sel.Collapse as (d: number) => void)(0)
        }
        let offset = 0
        for (const seg of segs) {
          const segLen = seg.text.length
          const needsFormat = seg.bold || seg.italic || seg.code || seg.strikethrough || seg.link
          if (needsFormat) {
            const range = (doc.Range as (s: number, e: number) => Record<string, unknown>)(basePos + offset, basePos + offset + segLen)
            if (seg.link) {
              try { ;(range.Hyperlinks as { Add: (r: unknown, a: string) => void }).Add(range, seg.link) } catch { /* ignore */ }
            }
            if (seg.bold || seg.italic || seg.strikethrough || seg.code) {
              const rangeFont = range.Font as Record<string, unknown>
              if (seg.bold) try { rangeFont.Bold = true } catch { /* ignore */ }
              if (seg.italic) try { rangeFont.Italic = true } catch { /* ignore */ }
              if (seg.strikethrough) try { rangeFont.Strikethrough = true } catch { /* ignore */ }
              if (seg.code) {
                try { rangeFont.Name = "Consolas" } catch { /* ignore */ }
                try { rangeFont.Size = 10.5 } catch { /* ignore */ }
                try { ;(range.Shading as Record<string, unknown>).BackgroundPatternColor = 0xF0F0F0 } catch { /* ignore */ }
              }
            }
          }
          offset += segLen
        }
      } else {
        for (const seg of segs) {
          const font = sel.Font as Record<string, unknown>
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
            ;(sel.TypeText as (t: string) => void)(cleaned)
            try {
              ;(sel.MoveStart as (u: number, c: number) => void)(1, -cleaned.length)
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
                ;(sel.MoveStart as (u: number, c: number) => void)(1, -cleaned.length)
                ;((sel.Range as Record<string, unknown>).Shading as Record<string, unknown>).BackgroundPatternColor = 0xF0F0F0
                ;(sel.Collapse as (d: number) => void)(0)
              } catch { /* ignore */ }
            }
          }
          if (changedBold) try { font.Bold = false } catch { /* ignore */ }
          if (changedItalic) try { font.Italic = false } catch { /* ignore */ }
          if (changedStrikethrough) try { font.Strikethrough = false } catch { /* ignore */ }
        }
        if (block.type === "heading") {
          ;(sel.MoveStart as (u: number, c: number) => void)(1, -fullText.length)
          this.applyHeadingStyle(sel, Math.min(block.level, 9))
          ;(sel.Collapse as (d: number) => void)(0)
        }
      }
      if (bi < totalBlocks - 1) { ;(sel.TypeParagraph as () => void)() }
      return block.text.length
    }

    if (block.type === "bullet_list" || block.type === "numbered_list") {
      if (!atCursor) this.goToEnd()
      const sel = this.getSelection()
      const lf = (sel.Range as Record<string, unknown>).ListFormat as Record<string, unknown>
      if (block.type === "bullet_list") {
        ;(lf.ApplyBulletDefault as () => void)()
      } else {
        ;(lf.ApplyNumberDefault as () => void)()
      }
      for (let idx = 0; idx < block.items.length; idx++) {
        for (let indent = 0; indent < block.items[idx].indent; indent++) {
          try { ;(lf.IncreaseIndent as () => void)() } catch { /* IncreaseIndent may not be available */ }
        }
        const segs = this.parseInline(block.items[idx].text)
        for (const seg of segs) this.typeSeg(sel, seg)
        if (idx < block.items.length - 1) { ;(sel.TypeParagraph as () => void)() }
      }
      ;(sel.TypeParagraph as () => void)()
      const freshLf = (sel.Range as Record<string, unknown>).ListFormat as Record<string, unknown>
      ;(freshLf.RemoveNumbers as () => void)()
      return block.items.reduce((s, item) => s + item.text.length, 0)
    }

    if (block.type === "table") {
      if (!atCursor) this.goToEnd()
      const doc = this.requireDoc()
      const rows = block.rows.length
      const cols = Math.max(...block.rows.map(r => r.length), 1)
      if (rows === 0 || cols === 0) return 0
      const range = this.getSelection().Range
      const tables = doc.Tables as {
        Add: (range: unknown, rows: number, cols: number) => Record<string, unknown>
      }
      const table = tables.Add(range, rows, cols)
      for (let r = 0; r < rows; r++) {
        for (let c = 0; c < cols; c++) {
          const text = block.rows[r][c] ?? ""
          ;((table.Cell as (r: number, c: number) => { Range: { Text: string } })(r + 1, c + 1).Range.Text as string) = text
        }
      }
      WordTableEditor.applyDefaultBorders(table)
      try { ;(table.AutoFitBehavior as (b: number) => void)(1) } catch { /* ignore */ }
      try {
        const headerRow = (table.Rows as { Item: (i: number) => Record<string, unknown> }).Item(1)
        ;((headerRow.Range as Record<string, unknown>).Font as Record<string, unknown>).Bold = true
        ;(headerRow.Shading as Record<string, unknown>).BackgroundPatternColor = 0xD9E2F3
      } catch { /* ignore */ }
      try { ;(this.getSelection().EndKey as (u: number) => void)(6) } catch { /* ignore */ }
      if (bi < totalBlocks - 1) { ;(this.getSelection().TypeParagraph as () => void)() }
      return block.rows.reduce((s, r) => s + r.reduce((a, t) => a + t.length, 0), 0)
    }

    if (block.type === "blockquote") {
      if (!atCursor) this.goToEnd()
      const sel = this.getSelection()
      const pf = sel.ParagraphFormat as Record<string, unknown>
      const prevIndent = (pf.LeftIndent as number) ?? 0
      ;(pf.LeftIndent as number) = prevIndent + 28.35
      ;(sel.Font as Record<string, unknown>).Italic = true
      const segs = this.parseInline(block.text)
      const fullText = segs.map(s => s.text).join("")
      for (const seg of segs) this.typeSeg(sel, seg)
      ;(sel.Font as Record<string, unknown>).Italic = false
      try {
        ;(sel.MoveStart as (u: number, c: number) => void)(1, -fullText.length)
        const rng = sel.Range as Record<string, unknown>
        ;(rng.Shading as Record<string, unknown>).BackgroundPatternColor = 0xF5F5F5
        const borders = rng.Borders as { Item: (i: number) => Record<string, unknown> }
        const b = borders.Item(1)
        b.LineStyle = 1
        b.ColorIndex = 15
        b.LineWidth = 8
        ;(sel.Collapse as (d: number) => void)(0)
      } catch { /* ignore */ }
      ;(pf.LeftIndent as number) = prevIndent
      if (bi < totalBlocks - 1) { ;(sel.TypeParagraph as () => void)() }
      return block.text.length
    }

    if (block.type === "codeblock") {
      if (!atCursor) this.goToEnd()
      const sel = this.getSelection()
      const doc = this.requireDoc()
      const codeLines = block.text.split("\n")
      const startPos = (doc.Content as Record<string, unknown>).End as number

      for (let li = 0; li < codeLines.length; li++) {
        ;(sel.TypeText as (t: string) => void)(sanitizeText(codeLines[li]))
        if (li < codeLines.length - 1) {
          ;(sel.TypeParagraph as () => void)()
        }
      }

      const endPos = (doc.Content as Record<string, unknown>).End as number
      try {
        const codeRange = (doc.Range as (s: number, e: number) => Record<string, unknown>)(startPos, endPos)
        ;(codeRange.Font as Record<string, unknown>).Name = "Consolas"
        ;(codeRange.Font as Record<string, unknown>).Size = 10.5
        ;(codeRange.Shading as Record<string, unknown>).BackgroundPatternColor = 0xF5F5F5
      } catch { /* ignore */ }

      try {
        const endRange = (doc.Range as (s: number, e: number) => Record<string, unknown>)(endPos, endPos)
        ;(endRange.Select as () => void)()
      } catch { /* ignore */ }
      if (bi < totalBlocks - 1) { ;(this.getSelection().TypeParagraph as () => void)() }
      return block.text.length
    }

    return 0
  }

  private applyHeadingStyle(sel: Record<string, unknown>, level: number): void {
    const candidates = [`Heading ${level}`, `标题 ${level}`]
    for (const name of candidates) {
      try { sel.Style = name; return } catch { /* try next */ }
    }
  }

  private typeSeg(sel: Record<string, unknown>, seg: InlineSegment): void {
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

  private parseInline(text: string): InlineSegment[] {
    const segments: InlineSegment[] = []
    let current = ""
    let bold = false
    let italic = false
    let code = false
    let strikethrough = false
    let i = 0

    while (i < text.length) {
      if (text[i] === "\\" && i + 1 < text.length) {
        current += text[i + 1]
        i += 2
        continue
      }

      if (text[i] === "`") {
        if (current) segments.push({ text: current, bold, italic, code, strikethrough })
        current = ""
        code = !code
        i++
        continue
      }

      if (code) {
        current += text[i]
        i++
        continue
      }

      if (text[i] === "[" && !code) {
        const close = text.indexOf("](", i + 1)
        if (close !== -1) {
          const closeParen = text.indexOf(")", close + 2)
          if (closeParen !== -1) {
            if (current) segments.push({ text: current, bold, italic, code, strikethrough })
            current = ""
            const linkText = text.slice(i + 1, close)
            const url = text.slice(close + 2, closeParen)
            segments.push({ text: linkText, bold, italic, code: false, strikethrough, link: url })
            i = closeParen + 1
            continue
          }
        }
      }

      if (text[i] === "~" && text[i + 1] === "~") {
        if (current) segments.push({ text: current, bold, italic, code, strikethrough })
        current = ""
        strikethrough = !strikethrough
        i += 2
        continue
      }

      if (text[i] === "*" && text[i + 1] === "*") {
        if (current) segments.push({ text: current, bold, italic, code, strikethrough })
        current = ""
        bold = !bold
        i += 2
        continue
      }

      if (text[i] === "*") {
        if (current) segments.push({ text: current, bold, italic, code, strikethrough })
        current = ""
        italic = !italic
        i++
        continue
      }

      current += text[i]
      i++
    }

    if (current || bold || italic || code || strikethrough) {
      segments.push({ text: current, bold, italic, code, strikethrough })
    }
    bold = false; italic = false; code = false; strikethrough = false
    return segments
  }

  private isLooseTableRow(line: string): boolean {
    if (!line.includes("|")) return false
    const cells = line.split("|").filter(c => !/^[-:\s]*$/.test(c))
    return cells.length >= 2
  }

  private parseBlocks(markdown: string): Block[] {
    const lines = markdown.split("\n")
    const blocks: Block[] = []
    let i = 0

    while (i < lines.length) {
      const raw = lines[i]
      const trimmed = raw.trim()

      if (trimmed === "") { i++; continue }

      if (/^#{1,6}\s+/.test(trimmed)) {
        const level = trimmed.match(/^#+/)![0].length
        const text = trimmed.replace(/^#+\s+/, "")
        blocks.push({ type: "heading", level, text })
        i++
        continue
      }

      if (trimmed === "[[pagebreak]]") {
        blocks.push({ type: "pagebreak" })
        i++
        continue
      }

      if (/^(-{3,}|\*{3,}|_{3,})\s*$/.test(trimmed)) {
        blocks.push({ type: "hr" })
        i++
        continue
      }

      if (trimmed.startsWith("```")) {
        i++
        const codeLines: string[] = []
        while (i < lines.length && !lines[i].trim().startsWith("```")) {
          codeLines.push(lines[i])
          i++
        }
        if (i < lines.length) i++
        blocks.push({ type: "codeblock", text: codeLines.join("\n") })
        continue
      }

      if (trimmed.startsWith("> ")) {
        blocks.push({ type: "blockquote", text: trimmed.slice(2) })
        i++
        continue
      }

      if (/^[-*+]\s/.test(trimmed)) {
        const items: ListItem[] = []
        while (i < lines.length && /^(\s*)[-*+]\s/.test(lines[i])) {
          const indent = Math.floor((lines[i].match(/^(\s*)/)?.[1]?.length ?? 0) / 2)
          const text = lines[i].trim().replace(/^[-*+]\s+/, "")
          if (text.trim()) {
            items.push({ text, indent })
          }
          i++
        }
        blocks.push({ type: "bullet_list", items })
        continue
      }

      if (/^\d+[.)]\s/.test(trimmed)) {
        const items: ListItem[] = []
        while (i < lines.length && /^\s*\d+[.)]\s/.test(lines[i])) {
          const text = lines[i].trim().replace(/^\d+[.)]\s+/, "")
          items.push({ text, indent: 0 })
          i++
        }
        blocks.push({ type: "numbered_list", items })
        continue
      }

      if (trimmed.startsWith("|") || this.isLooseTableRow(trimmed)) {
        const rows: string[][] = []
        while (i < lines.length && (lines[i].trim().startsWith("|") || this.isLooseTableRow(lines[i].trim()))) {
          const line = lines[i].trim()
          const parts = line.split("|").map(c => c.trim())
          const cells = parts.length >= 2 && line.trim().startsWith("|")
            ? parts.slice(1, parts.length - (line.trim().endsWith("|") ? 1 : 0))
            : parts
          if (!cells.every(c => /^[-:\s]+$/.test(c))) {
            rows.push(cells)
          }
          i++
        }
        blocks.push({ type: "table", rows })
        continue
      }

      if (trimmed.startsWith("![")) {
        const match = trimmed.match(/^!\[([^\]]*)\]\(([^)]+)\)/)
        if (match) {
          blocks.push({ type: "image", alt: match[1], url: match[2] })
          i++
          continue
        }
      }

      const paraLines: string[] = []
      while (i < lines.length) {
        const t = lines[i].trim()
        if (t === "") break
        if (/^(#{1,6}\s|[-*+]\s|\d+[.)]\s|> )/.test(t)) break
        if (t.startsWith("![")) break
        if (t === "[[pagebreak]]") break
        if (/^(-{3,}|\*{3,}|_{3,})\s*$/.test(t)) break
        if (t.startsWith("|") || this.isLooseTableRow(t) || t.startsWith("```")) break
        paraLines.push(t)
        i++
      }
      if (paraLines.length) {
        blocks.push({ type: "paragraph", text: paraLines.join("\n") })
      } else {
        i++
      }
    }

    return blocks
  }
}
