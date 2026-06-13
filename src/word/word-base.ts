import type { IWordSession } from "./session.js"
import { CursorPosition } from "./cursor-position.js"
import { WordMcpError } from "../security/errors.js"

export class WordBase {
  protected cursor: CursorPosition

  constructor(protected session: IWordSession) {
    this.cursor = new CursorPosition(session)
  }

  protected numOrEnum<T>(val: unknown, map: Record<string, T>): T | number {
    if (typeof val === "string") {
      const found = map[val]
      if (found !== undefined) return found
      const num = Number(val)
      if (!isNaN(num)) return num
      return 0
    }
    return val as number
  }

  protected getSelection(): Record<string, unknown> {
    this.cursor.markSelectionRead()
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
    if (!doc) throw new WordMcpError("No document is open", "NO_DOCUMENT", false, "Use word_document(path) to open a file, or word_stream_start to create a new document.")
    return doc
  }

  protected collapseSelection(): void {
    this.cursor.ensureMainBody()
  }

  protected goToEnd(): void {
    this.cursor.goToEnd()
  }

  protected sleep = (ms: number) => new Promise(r => setTimeout(r, ms))
}
