import type { IWordSession } from "./session.js"
import { type HeadingEntry } from "./types.js"
import { WordMcpError } from "../security/errors.js"
import { ContextSanitizer } from "./context-sanitizer.js"
import type { IDocumentProxy, IRangeProxy, ISelectionProxy } from "./com-proxy/types.js"

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
  private _tablesDirty = true
  private _headingsDirty = true
  private lastParaCount = 0
  private lastContentEnd = 0
  private refreshPromise: Promise<void> | null = null
  private cachedTexts: string[] = []
  private cachedParaStarts: number[] = []
  private _docVersion = 0

  constructor(private session: IWordSession) {}

  get cachedParaCount(): number { return this.lastParaCount }
  get cachedHeadingCount(): number { return this.headings.length }
  get cachedTableCount(): number { return this.tables.length }
  get docVersion(): number { return this._docVersion }

  scheduleRefresh(): void {
    if (!this.dirty || this.refreshPromise) return
    this.refreshPromise = (async () => {
      let attempts = 0
      const MAX_ATTEMPTS = 3
      do {
        attempts++
        try {
          await this.refresh()
        } catch (e) {
          this.session.logger?.warn(`[PositionMap] refresh failed (attempt ${attempts}/${MAX_ATTEMPTS})`)
          this.dirty = true
          if (attempts >= MAX_ATTEMPTS) {
            this.dirty = false
            break
          }
          await new Promise(r => setTimeout(r, 100 * Math.pow(2, attempts - 1)))
        }
      } while (this.dirty)
    })().finally(() => { this.refreshPromise = null })
  }

  async fetchActualParaCount(): Promise<number> {
    return this.session.getDocProxy().getParagraphs().count
  }

  async paraCountMatches(expected: number): Promise<boolean> {
    try {
      const actual = this.session.getDocProxy().getParagraphs().count
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
    this.headings = []
    this.tables = []
    await this.refreshParagraphs()
    await this.refreshTables()
    await this.refreshHeadings()
    this._tablesDirty = false
    this._headingsDirty = false
    this.dirty = false
  }

  private async refreshParagraphs(): Promise<void> {
    const doc = this.getDoc()
    const docProxy = this.session.getDocProxy()
    const paras = doc.Paragraphs as { Count: number; Item: (i: number) => Record<string, unknown> }
    const count = paras.Count as number
    this.lastParaCount = count

    const fullText = docProxy.getContent().getText()
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

    this.lastContentEnd = fullText.length
    this.cachedTexts = allTexts
    this.cachedParaStarts = paraStarts
  }

  private async refreshTables(): Promise<void> {
    const doc = this.getDoc()
    const count = this.lastParaCount
    const paraStarts = this.cachedParaStarts
    if (count === 0 || paraStarts.length < 2) return

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
  }

  private async refreshHeadings(): Promise<void> {
    const docProxy = this.session.getDocProxy()
    const allTexts = this.cachedTexts
    const count = this.lastParaCount
    const paraStarts = this.cachedParaStarts
    if (count === 0 || paraStarts.length < 2) return

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

    this.headings = []
    const headingEntries: Array<{ pi: number; level: number }> = []
    const headingSet = new Set<number>()

    const styleFormats: Array<(l: number) => string> = [
      (l) => `Heading ${l}`,
      (l) => `标题 ${l}`,
    ]

    for (const fmt of styleFormats) {
      for (let level = 1; level <= 9; level++) {
        try {
          const contentRange = docProxy.getContent().duplicate()
          const find = contentRange.getFind()
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
            const start = contentRange.getStart()
            const pi = binarySearchPara(start)

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
  }

  async ensureFresh(): Promise<void> {
    await this.ensureFreshInternal(true, true)
  }

  private async ensureFreshInternal(includeTables: boolean, includeHeadings: boolean): Promise<void> {
    if (this.refreshPromise) {
      await this.refreshPromise
      return
    }

    if (this.dirty) {
      await this.refreshParagraphs()
      this._tablesDirty = true
      this._headingsDirty = true
    }

    try {
      const docProxy = this.session.getDocProxy()
      const currentCount = docProxy.getParagraphs().count
      const currentContentEnd = docProxy.getContent().getEnd()

      if (currentCount !== this.lastParaCount) {
        await this.refreshParagraphs()
        this._tablesDirty = true
        this._headingsDirty = true
      } else if (currentContentEnd !== this.lastContentEnd && !this.dirty) {
        await this.refreshContentOnly()
      }
    } catch {
      await this.refreshParagraphs()
      this._tablesDirty = true
      this._headingsDirty = true
    }

    if (includeTables && this._tablesDirty && this.lastParaCount > 0) {
      this.tables = []
      await this.refreshTables()
      this._tablesDirty = false
    }
    if (includeHeadings && this._headingsDirty && this.lastParaCount > 0) {
      this.headings = []
      await this.refreshHeadings()
      this._headingsDirty = false
    }

    this.dirty = false
  }

  async ensureFreshParagraphsOnly(): Promise<void> {
    await this.ensureFreshInternal(false, false)
  }

  private async refreshContentOnly(): Promise<void> {
    const fullText = this.session.getDocProxy().getContent().getText()
    const rawTexts = fullText.split('\r')
    if (rawTexts.length > 0 && rawTexts[rawTexts.length - 1] === '') rawTexts.pop()
    const allTexts = rawTexts.slice(0, this.lastParaCount)

    for (const h of this.headings) {
      const newText = (allTexts[h.paragraphIndex - 1] ?? "").replace(/\r?\n$/, "")
      h.text = newText
    }

    this.cachedTexts = allTexts
    this.lastContentEnd = fullText.length
    this._tablesDirty = false
    this._headingsDirty = false
    this.dirty = false
  }

  markDirty(): void {
    this.dirty = true
    this._docVersion++
  }

  getHeadings(): HeadingEntry[] {
    return this.headings
  }

  getHeadingPath(paraIndex: number): HeadingEntry[] {
    const headings = this.headings
    if (headings.length === 0) return []

    let lo = 0, hi = headings.length - 1
    let nearestIdx = -1
    while (lo <= hi) {
      const mid = Math.floor((lo + hi) / 2)
      if (headings[mid].paragraphIndex <= paraIndex) {
        nearestIdx = mid
        lo = mid + 1
      } else {
        hi = mid - 1
      }
    }

    if (nearestIdx === -1) return []

    const path: HeadingEntry[] = [headings[nearestIdx]]
    let currentLevel = headings[nearestIdx].level
    for (let i = nearestIdx - 1; i >= 0; i--) {
      if (headings[i].level < currentLevel) {
        path.unshift(headings[i])
        currentLevel = headings[i].level
      }
    }
    return path
  }

  getParagraphIndex(charPos: number): number {
    const starts = this.cachedParaStarts
    if (starts.length < 2) return 1
    const paraCount = starts.length - 2
    let lo = 1, hi = paraCount
    while (lo <= hi) {
      const mid = Math.floor((lo + hi) / 2)
      if (charPos >= starts[mid]) {
        if (mid >= paraCount || charPos < starts[mid + 1]) return mid
        lo = mid + 1
      } else {
        hi = mid - 1
      }
    }
    return Math.min(paraCount, Math.max(1, lo))
  }

  async resolve(locator: Locator, skipFresh?: boolean): Promise<ResolvedPosition> {
    if (!skipFresh) {
      switch (locator.by) {
        case "heading":
          await this.ensureFreshInternal(true, true)
          break
        case "paragraph":
          await this.ensureFreshInternal(true, true)
          break
        case "table":
          await this.ensureFreshInternal(true, false)
          break
        case "bookmark":
          await this.ensureFreshInternal(false, false)
          break
      }
    }

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
    const docProxy = this.session.getDocProxy()
    const count = docProxy.getParagraphs().count

    const candidates: number[] = []
    const matchText = locator.match
    if (matchText) {
      const mode = locator.matchMode ?? "contains"
      const fullText = docProxy.getContent().getText()
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
        const docProxy = this.session.getDocProxy()
        const paraCount = docProxy.getParagraphs().count
        const fullText = docProxy.getContent().getText()
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
