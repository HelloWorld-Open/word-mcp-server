import { WordBase } from "./word-base.js"
import type { WordFormatting } from "./formatting.js"
import { WordMcpError } from "../security/errors.js"

export class WordTextEditor extends WordBase {
  constructor(session: import("./session.js").IWordSession, private formatting?: WordFormatting) {
    super(session)
  }
  private static readonly GOTO_ITEM: Record<string, number> = {
    page: 1, section: 2, line: 3, bookmark: 4, comment: 5, footnote: 6, endnote: 7,
    field: 8, table: 9, graphic: 10, equation: 11, object: 12, heading: 13,
  }
  private static readonly GOTO_DIRECTION: Record<string, number> = {
    first: 1, next: 2, previous: 3, last: -1,
  }

  async insertParagraph(count?: number): Promise<void> {
    this.collapseSelection()
    const sel = this.getSelection()
    const n = Math.max(1, count ?? 1)
    for (let i = 0; i < n; i++) {
      ;(sel.TypeParagraph as () => void)()
    }
  }

  async insertPageBreak(): Promise<void> {
    this.collapseSelection()
    ;(this.getSelection().InsertBreak as (t: number) => void)(7)
  }

  async insertHorizontalLine(): Promise<void> {
    this.collapseSelection()
    const inlineShapes = (this.getSelection().InlineShapes as { AddHorizontalLineStandard: () => void })
    inlineShapes.AddHorizontalLineStandard()
  }

  async findText(
    searchText: string,
    options?: { matchCase?: boolean; matchWholeWord?: boolean; direction?: string; wrap?: boolean }
  ): Promise<string> {
    this.collapseSelection()
    const sel = this.getSelection()
    const find = sel.Find as Record<string, unknown>
    ;(find.ClearFormatting as () => void)()
    if (options?.matchCase != null) find.MatchCase = !!options.matchCase
    if (options?.matchWholeWord != null) find.MatchWholeWord = !!options.matchWholeWord
    const forward = options?.direction !== "backward"
    const wrap = options?.wrap ?? true
    ;(find.Execute as (...args: unknown[]) => void)(
      searchText, !!options?.matchCase, !!options?.matchWholeWord,
      false, false, false, forward, wrap ? 1 : 0, false, "", 0
    )
    const found = (sel.Type as number) !== 1
    if (!found) return ""
    const range = sel.Range as Record<string, unknown>
    const text = (range.Text as string) ?? ""
    const start = range.Start as number
    return `Found at position ${start}. Text: "${text.slice(0, 200)}"`
  }

  async findReplace(
    findText: string,
    replaceWith: string,
    options?: { matchCase?: boolean; matchWholeWord?: boolean; replaceAll?: boolean; wrap?: boolean }
  ): Promise<void> {
    const find = this.getSelection().Find as Record<string, unknown>
    ;(find.ClearFormatting as () => void)()
    ;(((find as Record<string, unknown>).Replacement as Record<string, unknown>).ClearFormatting as () => void)()
    ;((find as Record<string, unknown>).Replacement as Record<string, unknown>).Text = replaceWith
    if (options?.matchCase != null) find.MatchCase = !!options.matchCase
    if (options?.matchWholeWord != null) find.MatchWholeWord = !!options.matchWholeWord
    const replaceMode = options?.replaceAll !== false ? 2 : 1
    const wrap = options?.wrap !== false ? 1 : 0
    this.collapseSelection()
    ;(find.Execute as (...args: unknown[]) => void)(
      findText, !!options?.matchCase, !!options?.matchWholeWord,
      false, false, false, true, wrap, false, replaceWith, replaceMode
    )
  }

  async goTo(what?: unknown, which?: unknown): Promise<void> {
    const sel = this.getSelection()
    const w = this.numOrEnum(what ?? "page", WordTextEditor.GOTO_ITEM)
    if (what === "end") {
      this.collapseSelection()
      const sel2 = this.getSelection()
      try { ;(sel2.EndKey as (u: number) => void)(6) } catch { }
      try { ;(sel2.Collapse as (d: number) => void)(0) } catch { }
      return
    }
    ;(sel.GoTo as (w: number, wh: number) => void)(w, this.numOrEnum(which ?? "first", WordTextEditor.GOTO_DIRECTION))
  }

  async goToParagraph(index: number): Promise<void> {
    const doc = this.requireDoc()
    const paras = doc.Paragraphs as { Count: number; Item: (i: number) => Record<string, unknown> }
    if (index < 1 || index > paras.Count) {
      throw new WordMcpError(
        `Paragraph index ${index} out of range (${paras.Count} paragraphs)`,
        "PARAGRAPH_NOT_FOUND", false,
        "Use word_get_structure to list paragraphs with their indices."
      )
    }
    const p = paras.Item(index)
    ;((p.Range as Record<string, unknown>).Select as () => void)()
  }

  async selectAll(): Promise<void> {
    ;(this.getSelection().WholeStory as () => void)()
  }

  async selectText(start: number, length: number): Promise<void> {
    const doc = this.requireDoc()
    const range = (doc.Range as (s: number, e: number) => Record<string, unknown>)(start, start + length)
    ;(range.Select as () => void)()
  }

  async selectCurrentWord(): Promise<void> {
    const sel = this.getSelection()
    ;((sel as Record<string, unknown>).Expand as (u: number) => void)(2)
  }

  async selectCurrentParagraph(): Promise<void> {
    const sel = this.getSelection()
    ;((sel as Record<string, unknown>).Expand as (u: number) => void)(4)
  }

  async deleteSelection(): Promise<void> {
    this.requireSelection()
    ;(this.getSelection().Delete as () => void)()
  }

  async backspace(count?: number): Promise<void> {
    this.collapseSelection()
    const sel = this.getSelection()
    const n = Math.max(1, count ?? 1)
    for (let i = 0; i < n; i++) {
      ;(sel.TypeBackspace as () => void)()
    }
  }

  async copy(): Promise<void> {
    this.requireSelection()
    ;(this.getSelection().Copy as () => void)()
  }

  async cut(): Promise<void> {
    this.requireSelection()
    ;(this.getSelection().Cut as () => void)()
  }

  async paste(): Promise<void> {
    this.collapseSelection()
    ;(this.getSelection().Paste as () => void)()
    ;(this.getSelection().EndKey as (u: number) => void)(6)
  }

  async undo(count?: number): Promise<void> {
    const doc = this.requireDoc()
    const n = Math.max(1, count ?? 1)
    for (let i = 0; i < n; i++) {
      ;((doc as Record<string, unknown>).Undo as () => void)()
    }
  }

  async redo(count?: number): Promise<void> {
    const doc = this.requireDoc()
    const n = Math.max(1, count ?? 1)
    for (let i = 0; i < n; i++) {
      ;((doc as Record<string, unknown>).Redo as () => void)()
    }
  }

  async getCursorInfo(): Promise<{ hasSelection: boolean; selectedText: string; start: number; end: number }> {
    const sel = this.getSelection()
    const selType = sel.Type as number
    const range = sel.Range as Record<string, unknown>
    return {
      hasSelection: selType !== 1,
      selectedText: selType !== 1 ? ((range.Text as string) ?? "") : "",
      start: range.Start as number,
      end: range.End as number,
    }
  }

  async insertList(type: "bullet" | "number", items: string[]): Promise<void> {
    this.collapseSelection()
    const sel = this.getSelection()
    const word = this.getWord()
    const applyList = () => {
      const lf = (sel.Range as Record<string, unknown>).ListFormat as Record<string, unknown>
      if (type === "bullet") {
        ;(lf.ApplyBulletDefault as () => void)()
      } else {
        ;(lf.ApplyNumberDefault as () => void)()
      }
    }
    const removeList = () => {
      const lf = (sel.Range as Record<string, unknown>).ListFormat as Record<string, unknown>
      ;(lf.RemoveNumbers as () => void)()
    }
    try {
      try { word.ScreenUpdating = false } catch { /* ignore */ }
      applyList()
      const TIME_BUDGET = 50
      let batchStart = Date.now()
      for (let i = 0; i < items.length; i++) {
        ;(sel.TypeText as (t: string) => void)(items[i])
        if (i < items.length - 1) {
          ;(sel.TypeParagraph as () => void)()
        }
        if (i < items.length - 1 && Date.now() - batchStart >= TIME_BUDGET) {
          try { word.ScreenUpdating = true } catch { /* ignore */ }
          try { ;(word.ScreenRefresh as () => void)() } catch { /* ignore */ }
          await new Promise(resolve => setImmediate(resolve))
          try { word.ScreenUpdating = false } catch { /* ignore */ }
          batchStart = Date.now()
        }
      }
      ;(sel.TypeParagraph as () => void)()
      removeList()
    } finally {
      try { word.ScreenUpdating = true } catch { /* ignore */ }
    }
  }

  async addHyperlink(
    text: string,
    address: string,
    subAddress?: string,
    screenTip?: string,
  ): Promise<void> {
    const sel = this.getSelection()
    const doc = this.requireDoc()
    const hyperlinks = doc.Hyperlinks as {
      Add: (anchor: unknown, address: string, subAddress?: string, screenTip?: string, textToDisplay?: string) => void
    }

    const hasSelection = (sel.Type as number) !== 1
    if (hasSelection) {
      hyperlinks.Add(sel.Range, address, subAddress, screenTip, text)
    } else {
      this.collapseSelection()
      const range = this.getSelection().Range
      hyperlinks.Add(range, address, subAddress, screenTip, text)
    }
    ;(sel.EndKey as (u: number) => void)(6)
  }

  async addFootnote(text: string): Promise<void> {
    this.collapseSelection()
    const doc = this.requireDoc()
    const footnotes = doc.Footnotes as { Add: (range: unknown, text: string) => void }
    await this.session.withScreenOff(async () => {
      try {
        footnotes.Add(this.getSelection().Range, text)
      } catch {
        // COM 失败回退：在文档末尾写入脚注标记文本
        try {
          this.formatting?.resetParagraphStyle()
          this.cursor.goToEnd()
          const sel = this.getSelection()
          ;(sel.TypeText as (t: string) => void)(`[脚注: ${text}]`)
        } catch { /* ignore */ }
      }
    })
    const end = (doc.Content as Record<string, unknown>).End as number
    const endRange = (doc.Range as (s: number, e: number) => Record<string, unknown>)(end, end)
    ;(endRange.Select as () => void)()
    ;(this.getSelection().Collapse as (d: number) => void)(0)
  }

  async insertSectionBreak(type?: string): Promise<void> {
    this.collapseSelection()
    const map: Record<string, number> = {
      nextPage: 8, continuous: 9, evenPage: 10, oddPage: 11,
    }
    ;(this.getSelection().Collapse as (d: number) => void)(1) // WdCollapseStart: don't replace selected paragraph
    ;(this.getSelection().InsertBreak as (t: number) => void)(map[type ?? "nextPage"] ?? 8)
  }

  async setColumns(count: number, spacing?: number): Promise<void> {
    const doc = this.requireDoc()
    const sections = doc.Sections as { Count: number; Item: (i: number) => Record<string, unknown> }
    const si = sections.Count
    const ps = sections.Item(si).PageSetup as Record<string, unknown>
    const textColumns = ps.TextColumns as Record<string, unknown>
    ;(textColumns.SetCount as (c: number) => void)(count)
    if (spacing != null) {
      const spacingPoints = Math.round(spacing * 28.3465)
      ;(textColumns.Spacing as number) = spacingPoints
    }
  }

  async insertFile(path: string): Promise<void> {
    this.collapseSelection()
    const sel = this.getSelection()
    ;((sel.Range as Record<string, unknown>).InsertFile as (p: string) => void)(path)
  }
}
