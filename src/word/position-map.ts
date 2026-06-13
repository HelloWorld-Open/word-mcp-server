import type { IWordSession } from "./session.js"
import { type HeadingEntry } from "./types.js"
import { WordMcpError } from "../security/errors.js"

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
}

export class PositionMap {
  private headings: HeadingEntry[] = []
  private tables: TableEntry[] = []
  private dirty = true
  private lastParaCount = 0
  private lastContentEnd = 0

  constructor(private session: IWordSession) {}

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

    const paraStarts = new Array(count + 2)
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

    // 使用 Find.Style 搜索标题（字符串赋在 find 自身上，与 find.Text 同模式，避免嵌套 COM 属性问题）
    const sel = (doc.Application as Record<string, unknown>).Selection as Record<string, unknown>
    const origStart = sel.Start as number
    const origEnd = sel.End as number
    const find = sel.Find as Record<string, unknown>

    const headingEntries: Array<{ pi: number; level: number }> = []
    const headingSet = new Set<number>()

    const styleFormats: Array<(l: number) => string> = [
      (l) => `Heading ${l}`,
      (l) => `标题 ${l}`,
    ]

    for (const fmt of styleFormats) {
      for (let level = 1; level <= 9; level++) {
        try {
          ;(find.ClearFormatting as () => void)()
          find.Style = fmt(level)
          find.Text = ""
          find.Forward = true
          find.Wrap = 0

          ;(sel.HomeKey as (u: number) => void)(6)

          while (true) {
            const found = (find.Execute as (...a: unknown[]) => boolean)(
              "", false, false, false, false, false, true, 0, true, "", 0
            )
            if (!found) break
            const start = sel.Start as number
            const pi = binarySearchPara(start)
            if (!headingSet.has(pi) && pi >= 1 && pi <= count) {
              headingSet.add(pi)
              headingEntries.push({ pi, level })
            }
          }
        } catch { /* skip this heading level */ }
      }
    }

    try {
      const restore = (doc.Range as (s: number, e: number) => Record<string, unknown>)(origStart, origEnd)
      ;(restore.Select as () => void)()
    } catch { /* ignore */ }

    headingEntries.sort((a, b) => a.pi - b.pi)
    for (const e of headingEntries) {
      const text = (allTexts[e.pi - 1] ?? "").replace(/\r?\n$/, "")
      this.headings.push({ text, level: e.level, paragraphIndex: e.pi })
    }

    // 表格段落映射：用二分查找替代 O(n²) 内循环
    const tables = doc.Tables as { Count: number; Item: (i: number) => Record<string, unknown> }
    const tableCount = tables.Count as number
    for (let t = 1; t <= tableCount; t++) {
      const table = tables.Item(t)
      const start = (table.Range as Record<string, unknown>).Start as number
      const pi = binarySearchPara(start)
      this.tables.push({ paragraphIndex: pi })
    }

    this.lastContentEnd = fullText.length
    this.dirty = false
  }

  async ensureFresh(): Promise<void> {
    if (this.dirty) { await this.refresh(); return }
    try {
      const doc = this.getDoc()
      const currentCount = (doc.Paragraphs as { Count: number }).Count as number
      const currentContentEnd = (doc.Content as Record<string, unknown>).End as number
      if (currentCount !== this.lastParaCount) {
        await this.refresh()
      } else if (currentContentEnd !== this.lastContentEnd) {
        this.lastContentEnd = currentContentEnd
      }
    } catch {
      await this.refresh()
    }
  }

  markDirty(): void {
    this.dirty = true
  }

  getHeadings(): HeadingEntry[] {
    return this.headings
  }

  async resolve(locator: Locator): Promise<ResolvedPosition> {
    await this.ensureFresh()

    switch (locator.by) {
      case "heading": return this.resolveHeading(locator)
      case "paragraph": return this.resolveParagraph(locator)
      case "table": return this.resolveTable(locator)
      case "bookmark": return this.resolveBookmark(locator)
    }
  }

  private async resolveHeading(locator: HeadingLocator): Promise<ResolvedPosition> {
    let candidates = this.headings

    const matchText = locator.match
    if (matchText != null) {
      const mode = locator.matchMode ?? "exact"
      if (mode === "exact") {
        candidates = this.headings.filter(h => h.text === matchText)
      } else if (mode === "contains") {
        candidates = this.headings.filter(h => h.text.includes(matchText))
      } else if (mode === "startsWith") {
        candidates = this.headings.filter(h => h.text.startsWith(matchText))
      } else if (mode === "regex" && matchText) {
        try {
          const re = new RegExp(matchText)
          candidates = this.headings.filter(h => re.test(h.text))
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
        const paraCount = (doc.Paragraphs as { Count: number }).Count as number
        const fullText = (doc.Content as Record<string, unknown>).Text as string
        const rawTexts = fullText.split('\r')
        if (rawTexts.length > 0 && rawTexts[rawTexts.length - 1] === '') rawTexts.pop()
        const allTexts = rawTexts.slice(0, paraCount)
        const paraStarts = new Array(paraCount + 2)
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
    let best: HeadingEntry | null = null
    for (const h of this.headings) {
      if (h.paragraphIndex <= paraIndex) {
        if (!best || h.paragraphIndex > best.paragraphIndex) {
          best = h
        }
      }
    }
    return best ? `${best.text} (H${best.level})` : null
  }
}
