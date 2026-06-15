import type { IWordSession } from "./session.js"
import { type HeadingEntry } from "./types.js"
import { WordMcpError } from "../security/errors.js"
import { ContextSanitizer } from "./context-sanitizer.js"

export interface Offset {
  direction: "before" | "after"
  count?: number
}

export interface HeadingLocator {
  by: "heading"
  match?: string
  matchMode?: "exact" | "contains" | "startsWith" | "regex"
  occurrence?: number
  offset?: Offset
  level?: number
}

export interface ParagraphLocator {
  by: "paragraph"
  match?: string
  matchMode?: "contains" | "regex"
  occurrence?: number
  offset?: Offset
}

export interface TableLocator {
  by: "table"
  occurrence?: number
  offset?: Offset
}

export interface BookmarkLocator {
  by: "bookmark"
  name: string
  offset?: Offset
}

export type Locator = HeadingLocator | ParagraphLocator | TableLocator | BookmarkLocator

export interface ResolvedPosition {
  found: boolean
  paragraphIndex: number
  headingContext: string | null
  tableIndex?: number
  error?: string
}

interface TableEntry {
  paragraphIndex: number
  rangeStart: number
  rangeEnd: number
}

export class PositionMap {
  private headings: HeadingEntry[] = []
  private tables: TableEntry[] = []
  private dirty = true
  private lastParaCount = 0
  private lastContentEnd = 0
  private refreshPromise: Promise<void> | null = null
  private cachedTexts: string[] = []
  private cachedParaStarts: number[] = []

  constructor(private session: IWordSession) {}

  get cachedParaCount(): number { return this.lastParaCount }

  scheduleRefresh(): void {
    if (!this.dirty || this.refreshPromise) return
    this.refreshPromise = (async () => {
      do {
        try {
          await this.refresh()
        } catch {
          this.dirty = true
        }
      } while (this.dirty)
    })().finally(() => { this.refreshPromise = null })
  }

  async fetchActualParaCount(): Promise<number> {
    const doc = this.getDoc()
    return (doc.Paragraphs as { Count: number }).Count as number
  }

  async paraCountMatches(expected: number): Promise<boolean> {
    try {
      const doc = this.getDoc()
      const actual = (doc.Paragraphs as { Count: number }).Count as number
      return actual === expected
    } catch {
      return false
    }
  }

  private getDoc(): Record<string, unknown> {
    const doc = this.session.activeDoc ?? (this.session.application as Record<string, unknown>).ActiveDocument as Record<string, unknown> | undefined
    if (!doc) throw new WordMcpError("No document is open", "NO_DOCUMENT", false, "Use word_document or word_stream_start first.")
    return doc
  }

  async refresh(): Promise<void> {
    const doc = this.getDoc()
    const paras = doc.Paragraphs as { Count: number; Item: (i: number) => Record<string, unknown> }
    const count = paras.Count as number
    this.headings = []
    this.tables = []
    this.lastParaCount = count

    // 获取全文并本地计算段落位置（1 次 COM 调用，替代 888 次逐段调用）
    const fullText = (doc.Content as Record<string, unknown>).Text as string
    const rawTexts = fullText.split('\r')
    if (rawTexts.length > 0 && rawTexts[rawTexts.length - 1] === '') rawTexts.pop()
    const allTexts = rawTexts.slice(0, count)
    while (allTexts.length < count) allTexts.push('')

    const paraStarts = Array.from({ length: count + 2 }, () => 0)
    let textPos = 0
    for (let i = 1; i <= count; i++) {
      paraStarts[i] = textPos
      textPos += allTexts[i - 1].length + 1
    }
    paraStarts[count + 1] = fullText.length

    const binarySearchPara = (startPos: number): number => {
      let lo = 1, hi = count
      while (lo <= hi) {
        const mid = Math.floor((lo + hi) / 2)
        if (startPos >= paraStarts[mid]) {
          if (mid >= count || startPos < paraStarts[mid + 1]) return mid
          lo = mid + 1
        } else {
          hi = mid - 1
        }
      }
      return Math.min(count, Math.max(1, lo))
    }

    // 先构建表格索引（包括字符范围），用于过滤表格内的标题检测
    const tables = doc.Tables as { Count: number; Item: (i: number) => Record<string, unknown> }
    const tableCount = tables.Count as number
    for (let t = 1; t <= tableCount; t++) {
      const table = tables.Item(t)
      const range = table.Range as Record<string, unknown>
      const start = range.Start as number
      const end = range.End as number
      const pi = binarySearchPara(start)
      this.tables.push({ paragraphIndex: pi, rangeStart: start, rangeEnd: end })
    }

    // 使用 Range.Find 搜索标题（不影响用户光标），排除表格内部段落
    const headingEntries: Array<{ pi: number; level: number }> = []
    const headingSet = new Set<number>()

    const styleFormats: Array<(l: number) => string> = [
      (l) => `Heading ${l}`,
      (l) => `标题 ${l}`,
    ]

    for (const fmt of styleFormats) {
      for (let level = 1; level <= 9; level++) {
        try {
          const contentRange = (doc.Content as Record<string, unknown>).Duplicate as Record<string, unknown>
          const find = contentRange.Find as Record<string, unknown>
          ;(find.ClearFormatting as () => void)()
          find.Style = fmt(level)
          find.Text = ""
          find.Forward = true
          find.Wrap = 0

          while (true) {
            const found = (find.Execute as (...a: unknown[]) => boolean)(
              "", false, false, false, false, false, true, 0, true, "", 0
            )
            if (!found) break
            const start = contentRange.Start as number
            const pi = binarySearchPara(start)

            // 跳过表格内部的段落（表格单元格可能误继承标题样式）
            const insideTable = this.tables.some(t => start >= t.rangeStart && start < t.rangeEnd)
            if (!headingSet.has(pi) && pi >= 1 && pi <= count && !insideTable) {
              headingSet.add(pi)
              headingEntries.push({ pi, level })
            }
          }
        } catch { /* skip this heading level */ }
      }
    }

    headingEntries.sort((a, b) => a.pi - b.pi)
    for (const e of headingEntries) {
      const text = ContextSanitizer.stripBel(allTexts[e.pi - 1] ?? "")
      this.headings.push({ text, level: e.level, paragraphIndex: e.pi })
    }

    this.lastContentEnd = fullText.length
    this.cachedTexts = allTexts
    this.cachedParaStarts = paraStarts
    this.dirty = false
  }

  async ensureFresh(): Promise<void> {
    if (this.refreshPromise) {
      await this.refreshPromise
      return
    }
    if (this.dirty) { await this.refresh(); return }
    try {
      const doc = this.getDoc()
      const currentCount = (doc.Paragraphs as { Count: number }).Count as number
      const currentContentEnd = (doc.Content as Record<string, unknown>).End as number
      if (currentCount !== this.lastParaCount) {
        await this.refresh()
      } else if (currentContentEnd !== this.lastContentEnd) {
        await this.refreshContentOnly()
      }
    } catch {
      await this.refresh()
    }
  }

  private async refreshContentOnly(): Promise<void> {
    const doc = this.getDoc()
    const fullText = (doc.Content as Record<string, unknown>).Text as string
    const rawTexts = fullText.split('\r')
    if (rawTexts.length > 0 && rawTexts[rawTexts.length - 1] === '') rawTexts.pop()
    const allTexts = rawTexts.slice(0, this.lastParaCount)

    for (const h of this.headings) {
      const newText = (allTexts[h.paragraphIndex - 1] ?? "").replace(/\r?\n$/, "")
      h.text = newText
    }

    this.cachedTexts = allTexts
    this.lastContentEnd = fullText.length
    this.dirty = false
  }

  markDirty(): void {
    this.dirty = true
  }

  getHeadings(): HeadingEntry[] {
    return this.headings
  }

  async resolve(locator: Locator, skipFresh?: boolean): Promise<ResolvedPosition> {
    if (!skipFresh) await this.ensureFresh()

    switch (locator.by) {
      case "heading": return this.resolveHeading(locator)
      case "paragraph": return this.resolveParagraph(locator)
      case "table": return this.resolveTable(locator)
      case "bookmark": return this.resolveBookmark(locator)
    }
  }

  private async resolveHeading(locator: HeadingLocator): Promise<ResolvedPosition> {
    let candidates = this.headings

    if (locator.level != null) {
      candidates = candidates.filter(h => h.level === locator.level)
    }

    const matchText = locator.match
    if (matchText != null) {
      const mode = locator.matchMode ?? "exact"
      if (mode === "exact") {
        candidates = candidates.filter(h => h.text === matchText)
      } else if (mode === "contains") {
        candidates = candidates.filter(h => h.text.includes(matchText))
      } else if (mode === "startsWith") {
        candidates = candidates.filter(h => h.text.startsWith(matchText))
      } else if (mode === "regex" && matchText) {
        try {
          const re = new RegExp(matchText)
          candidates = candidates.filter(h => re.test(h.text))
        } catch {
          return { found: false, paragraphIndex: 0, headingContext: null, error: `Invalid regex: ${matchText}` }
        }
      }
    }

    const occ = (locator.occurrence ?? 1) - 1
    if (occ < 0 || occ >= candidates.length) {
      return { found: false, paragraphIndex: 0, headingContext: null, error: `Heading not found: ${locator.match ?? `#${occ + 1}`}` }
    }

    const match = candidates[occ]
    const base = this.applyOffset(match.paragraphIndex, locator.offset)
    return { ...base, headingContext: `${match.text} (H${match.level})` }
  }

  private async resolveParagraph(locator: ParagraphLocator): Promise<ResolvedPosition> {
    const allTexts = this.cachedTexts
    if (allTexts.length === 0) {
      return this.resolveParagraphCOM(locator)
    }

    const candidates: number[] = []
    const matchText = locator.match
    if (matchText) {
      const mode = locator.matchMode ?? "contains"
      for (let i = 0; i < allTexts.length; i++) {
        const text = allTexts[i].replace(/\r?\n$/, "")
        if (mode === "contains" && text.includes(matchText)) {
          candidates.push(i + 1)
        } else if (mode === "regex") {
          try {
            const re = new RegExp(matchText)
            if (re.test(text)) candidates.push(i + 1)
          } catch {
            return { found: false, paragraphIndex: 0, headingContext: null, error: `Invalid regex: ${matchText}` }
          }
        }
      }
    } else {
      for (let i = 1; i <= allTexts.length; i++) candidates.push(i)
    }

    const occ = (locator.occurrence ?? 1) - 1
    if (occ < 0 || occ >= candidates.length) {
      return { found: false, paragraphIndex: 0, headingContext: null, error: `Paragraph not found: ${locator.match ?? `#${occ + 1}`}` }
    }

    const paraIdx = candidates[occ]
    const base = this.applyOffset(paraIdx, locator.offset)
    const ctx = this.getHeadingContext(base.paragraphIndex)
    return { ...base, headingContext: ctx ?? base.headingContext }
  }

  private async resolveParagraphCOM(locator: ParagraphLocator): Promise<ResolvedPosition> {
    const doc = this.getDoc()
    const count = (doc.Paragraphs as { Count: number }).Count as number

    const candidates: number[] = []
    const matchText = locator.match
    if (matchText) {
      const mode = locator.matchMode ?? "contains"
      const fullText = (doc.Content as Record<string, unknown>).Text as string
      const rawTexts = fullText.split('\r')
      if (rawTexts.length > 0 && rawTexts[rawTexts.length - 1] === '') rawTexts.pop()
      const allTexts = rawTexts.slice(0, count)
      for (let i = 0; i < allTexts.length; i++) {
        const text = allTexts[i].replace(/\r?\n$/, "")
        if (mode === "contains" && text.includes(matchText)) {
          candidates.push(i + 1)
        } else if (mode === "regex") {
          try {
            const re = new RegExp(matchText)
            if (re.test(text)) candidates.push(i + 1)
          } catch {
            return { found: false, paragraphIndex: 0, headingContext: null, error: `Invalid regex: ${matchText}` }
          }
        }
      }
    } else {
      for (let i = 1; i <= count; i++) candidates.push(i)
    }

    const occ = (locator.occurrence ?? 1) - 1
    if (occ < 0 || occ >= candidates.length) {
      return { found: false, paragraphIndex: 0, headingContext: null, error: `Paragraph not found: ${locator.match ?? `#${occ + 1}`}` }
    }

    const paraIdx = candidates[occ]
    const base = this.applyOffset(paraIdx, locator.offset)
    const ctx = this.getHeadingContext(base.paragraphIndex)
    return { ...base, headingContext: ctx ?? base.headingContext }
  }

  private async resolveTable(locator: TableLocator): Promise<ResolvedPosition> {
    if (this.tables.length === 0) {
      return { found: false, paragraphIndex: 0, headingContext: null, error: "No tables in document" }
    }

    const occ = (locator.occurrence ?? 1) - 1
    if (occ < 0 || occ >= this.tables.length) {
      return { found: false, paragraphIndex: 0, headingContext: null, error: `Table #${occ + 1} not found (${this.tables.length} tables)` }
    }

    const tableEntry = this.tables[occ]
    const base = this.applyOffset(tableEntry.paragraphIndex, locator.offset)
    const ctx = this.getHeadingContext(base.paragraphIndex)
    return { ...base, tableIndex: occ + 1, headingContext: ctx ?? base.headingContext }
  }

  private async resolveBookmark(locator: BookmarkLocator): Promise<ResolvedPosition> {
    const doc = this.getDoc()
    const bookmarks = doc.Bookmarks as { Count: number; Item: (i: number) => Record<string, unknown> }
    const count = bookmarks.Count as number
    for (let i = 1; i <= count; i++) {
      const b = bookmarks.Item(i)
      if ((b.Name as string) === locator.name) {
        const range = b.Range as Record<string, unknown>
        const start = range.Start as number
        if (this.cachedParaStarts.length > 0) {
          const paraCount = this.cachedParaStarts.length - 2
          const ps = this.cachedParaStarts
          let lo = 1, hi = paraCount
          while (lo <= hi) {
            const mid = Math.floor((lo + hi) / 2)
            if (start >= ps[mid]) {
              if (mid >= paraCount || start < ps[mid + 1]) {
                const base = this.applyOffset(mid, locator.offset)
                const ctx = this.getHeadingContext(base.paragraphIndex)
                return { ...base, headingContext: ctx ?? base.headingContext }
              }
              lo = mid + 1
            } else {
              hi = mid - 1
            }
          }
          const base = this.applyOffset(1, locator.offset)
          return { ...base, headingContext: null }
        }
        const paraCount = (doc.Paragraphs as { Count: number }).Count as number
        const fullText = (doc.Content as Record<string, unknown>).Text as string
        const rawTexts = fullText.split('\r')
        if (rawTexts.length > 0 && rawTexts[rawTexts.length - 1] === '') rawTexts.pop()
        const allTexts = rawTexts.slice(0, paraCount)
        const paraStarts = Array.from({ length: paraCount + 2 }, () => 0)
        let textPos = 0
        for (let i = 1; i <= paraCount; i++) {
          paraStarts[i] = textPos
          textPos += allTexts[i - 1].length + 1
        }
        paraStarts[paraCount + 1] = fullText.length
        let lo = 1, hi = paraCount
        while (lo <= hi) {
          const mid = Math.floor((lo + hi) / 2)
          if (start >= paraStarts[mid]) {
            if (mid >= paraCount || start < paraStarts[mid + 1]) {
              const base = this.applyOffset(mid, locator.offset)
              const ctx = this.getHeadingContext(base.paragraphIndex)
              return { ...base, headingContext: ctx ?? base.headingContext }
            }
            lo = mid + 1
          } else {
            hi = mid - 1
          }
        }
        const base = this.applyOffset(1, locator.offset)
        return { ...base, headingContext: null }
      }
    }
    return { found: false, paragraphIndex: 0, headingContext: null, error: `Bookmark "${locator.name}" not found` }
  }

  private applyOffset(basePara: number, offset?: Offset): ResolvedPosition {
    if (!offset) {
      const ctx = this.getHeadingContext(basePara)
      return { found: true, paragraphIndex: basePara, headingContext: ctx }
    }
    const dir = offset.direction === "after" ? 1 : -1
    const count = offset.count ?? 1
    const target = Math.max(1, basePara + dir * count)
    const ctx = this.getHeadingContext(target)
    return { found: true, paragraphIndex: target, headingContext: ctx }
  }

  private getHeadingContext(paraIndex: number): string | null {
    const headings = this.headings
    if (headings.length === 0) return null
    let lo = 0, hi = headings.length - 1
    let best: HeadingEntry | null = null
    while (lo <= hi) {
      const mid = Math.floor((lo + hi) / 2)
      if (headings[mid].paragraphIndex <= paraIndex) {
        best = headings[mid]
        lo = mid + 1
      } else {
        hi = mid - 1
      }
    }
    return best ? `${best.text} (H${best.level})` : null
  }
}
