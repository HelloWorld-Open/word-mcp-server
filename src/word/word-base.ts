import type { IWordSession } from "./session.js"
import { WordMcpError } from "../security/errors.js"

export class WordBase {
  constructor(protected session: IWordSession) {}

  protected numOrEnum<T>(val: unknown, map: Record<string, T>): T | number {
    if (typeof val === "string") return map[val] ?? (val as unknown as T)
    return val as number
  }

  protected getSelection(): Record<string, unknown> {
    return this.session.comCall(() =>
      (this.getWord().Selection as Record<string, unknown>) as Record<string, unknown>
    )
  }

  protected requireSelection(): void {
    const sel = this.getSelection()
    const start = sel.Start as number
    const end = sel.End as number
    if (start === end) throw new WordMcpError(
      "No text is selected",
      "NO_SELECTION",
      false,
      "Use word_select_all(), word_select_text(), word_find_text(), or word_select_current_word() first."
    )
  }

  protected getWord(): Record<string, unknown> {
    return this.session.comCall(() =>
      this.session.application as Record<string, unknown>
    )
  }

  protected requireDoc(): Record<string, unknown> {
    const doc = this.session.activeDoc ?? (this.getWord().ActiveDocument as Record<string, unknown>)
    if (!doc) throw new WordMcpError("No document is open", "NO_DOCUMENT", false, "Use word_document(path) to open a file, or word_create to create a new document.")
    return doc
  }

  private _cachedStart = -1
  private _cachedEnd = -1

  protected collapseSelection(): void {
    try {
      const sel = this.getSelection()
      const storyType = sel.StoryType as number
      if (storyType === 1) {
        const start = sel.Start as number
        const end = sel.End as number
        if (start === this._cachedStart && end === this._cachedEnd) return
        this._cachedStart = start
        this._cachedEnd = end
      }

      const doc = this.session.activeDoc ?? (this.getWord().ActiveDocument as Record<string, unknown>)
      if (!doc) return

      if (storyType !== 1) {
        const contentEnd = (doc.Content as Record<string, unknown>).End as number
        const endRange = (doc.Range as (s: number, e: number) => Record<string, unknown>)(contentEnd, contentEnd)
        ;(endRange.Select as () => void)()
        ;((this.getSelection()).Collapse as (d: number) => void)(0)
        this.session.wasInNonBody = true
        return
      }

      if (this.session.wasInNonBody) {
        ;(sel.TypeParagraph as () => void)()
        this.session.wasInNonBody = false
      }

      const wdWithInTable = 12
      try {
        if ((sel.Information as (t: number) => boolean)(wdWithInTable)) {
          const table = (sel.Tables as { Item: (i: number) => Record<string, unknown> }).Item(1)
          ;((table.Range as Record<string, unknown>).Select as () => void)()
          ;((this.getSelection()).Collapse as (d: number) => void)(0)
        }
      } catch {
      }

      try {
        const shapes = sel.ShapeRange as { Count: number } | undefined
        if (shapes && shapes.Count > 0) {
          const contentEnd = (doc.Content as Record<string, unknown>).End as number
          const endRange = (doc.Range as (s: number, e: number) => Record<string, unknown>)(contentEnd, contentEnd)
          ;(endRange.Select as () => void)()
          ;((this.getSelection()).Collapse as (d: number) => void)(0)
        }
      } catch { }
    } catch {
    }
  }

  protected sleep = (ms: number) => new Promise(r => setTimeout(r, ms))
}
