import { WordBase } from "./word-base.js"
import { WordMcpError } from "../security/errors.js"
import type { IDocumentProxy, ISelectionProxy, IRangeProxy } from "./com-proxy/types.js"

export class WordCursor extends WordBase {
  constructor(session: import("./session.js").IWordSession) {
    super(session)
  }
  private static readonly GOTO_ITEM: Record<string, number> = {
    page: 1, section: 2, line: 3, bookmark: 4, comment: 5, footnote: 6, endnote: 7,
    field: 8, table: 9, graphic: 10, equation: 11, object: 12, heading: 13,
  }
  private static readonly GOTO_DIRECTION: Record<string, number> = {
    first: 1, next: 2, previous: 3, last: -1,
  }

  async findText(
    searchText: string,
    options?: { matchCase?: boolean; matchWholeWord?: boolean; direction?: string; wrap?: boolean }
  ): Promise<string> {
    this.collapseSelection()
    const sel = this.getSelProxy()
    const find = sel.getFind()
    ;(find.ClearFormatting as () => void)()
    if (options?.matchCase != null) find.MatchCase = !!options.matchCase
    if (options?.matchWholeWord != null) find.MatchWholeWord = !!options.matchWholeWord
    const forward = options?.direction !== "backward"
    const wrap = options?.wrap ?? true
    ;(find.Execute as (...args: unknown[]) => void)(
      searchText, !!options?.matchCase, !!options?.matchWholeWord,
      false, false, false, forward, wrap ? 1 : 0, false, "", 0
    )
    const found = sel.getType() !== 1
    if (!found) return ""
    const range = sel.getRange()
    const text = range.getText() ?? ""
    const start = range.getStart()
    return `Found at position ${start}. Text: "${text.slice(0, 200)}"`
  }

  async findReplace(
    findText: string,
    replaceWith: string,
    options?: { matchCase?: boolean; matchWholeWord?: boolean; replaceAll?: boolean; wrap?: boolean }
  ): Promise<void> {
    this.collapseSelection()
    const find = this.getSelProxy().getFind()
    ;(find.ClearFormatting as () => void)()
    ;(((find as Record<string, unknown>).Replacement as Record<string, unknown>).ClearFormatting as () => void)()
    ;((find as Record<string, unknown>).Replacement as Record<string, unknown>).Text = replaceWith
    if (options?.matchCase != null) find.MatchCase = !!options.matchCase
    if (options?.matchWholeWord != null) find.MatchWholeWord = !!options.matchWholeWord
    const replaceMode = options?.replaceAll !== false ? 2 : 1
    const wrap = options?.wrap !== false ? 1 : 0
    ;(find.Execute as (...args: unknown[]) => void)(
      findText, !!options?.matchCase, !!options?.matchWholeWord,
      false, false, false, true, wrap, false, replaceWith, replaceMode
    )
  }

  async goTo(what?: unknown, which?: unknown): Promise<void> {
    const sel = this.getSelProxy()
    const w = this.numOrEnum(what ?? "page", WordCursor.GOTO_ITEM)
    if (what === "end") {
      this.goToEnd()
      return
    }
    sel.goTo(w, this.numOrEnum(which ?? "first", WordCursor.GOTO_DIRECTION))
  }

  async goToParagraph(index: number): Promise<void> {
    const doc = this.getDocProxy()
    const paras = doc.getParagraphs()
    if (index < 1 || index > paras.count) {
      throw new WordMcpError(
        `Paragraph index ${index} out of range (${paras.count} paragraphs)`,
        "PARAGRAPH_NOT_FOUND", false,
        "Use word_get_structure to list paragraphs with their indices."
      )
    }
    const p = paras.item(index)
    ;((p.Range as Record<string, unknown>).Select as () => void)()
  }

  async selectAll(): Promise<void> {
    this.getSelProxy().wholeStory()
  }

  async selectText(start: number, length: number): Promise<void> {
    const doc = this.getDocProxy()
    const range = doc.getRange(start, start + length)
    range.select()
  }

  async selectCurrentWord(): Promise<void> {
    const sel = this.getSelProxy()
    sel.expand(2)
  }

  async selectCurrentParagraph(): Promise<void> {
    const sel = this.getSelProxy()
    sel.expand(4)
  }

  async deleteSelection(): Promise<void> {
    this.requireSelection()
    this.getSelProxy().delete()
  }

  async backspace(count?: number): Promise<void> {
    this.collapseSelection()
    const sel = this.getSelProxy()
    const n = Math.max(1, count ?? 1)
    for (let i = 0; i < n; i++) {
      sel.typeBackspace()
    }
  }

  async copy(): Promise<void> {
    this.requireSelection()
    this.getSelProxy().copy()
  }

  async cut(): Promise<void> {
    this.requireSelection()
    this.getSelProxy().cut()
  }

  async paste(): Promise<void> {
    this.collapseSelection()
    this.getSelProxy().paste()
    this.goToEnd()
  }

  async undo(count?: number): Promise<void> {
    const doc = this.getDocProxy()
    const n = Math.max(1, count ?? 1)
    for (let i = 0; i < n; i++) {
      doc.undo()
    }
  }

  async redo(count?: number): Promise<void> {
    const doc = this.getDocProxy()
    const n = Math.max(1, count ?? 1)
    for (let i = 0; i < n; i++) {
      doc.redo()
    }
  }

  async getCursorInfo(): Promise<{ hasSelection: boolean; selectedText: string; start: number; end: number }> {
    const sel = this.getSelProxy()
    const selType = sel.getType()
    const range = sel.getRange()
    return {
      hasSelection: selType !== 1,
      selectedText: selType !== 1 ? (range.getText() ?? "") : "",
      start: range.getStart(),
      end: range.getEnd(),
    }
  }

  async insertFile(path: string): Promise<void> {
    this.collapseSelection()
    const range = this.getSelProxy().getRange()
    range.insertFile(path)
  }
}
