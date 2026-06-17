import type { IWordSession } from "./session.js"
import type { IDocumentProxy, ISelectionProxy } from "./com-proxy/types.js"
import { ContextSanitizer, type ICursorContext } from "./context-sanitizer.js"
import { WordMcpError } from "../security/errors.js"

export class WordBase {
  protected cursor: ICursorContext

  constructor(protected session: IWordSession, cursor?: ICursorContext) {
    this.cursor = cursor ?? new ContextSanitizer(session)
  }

  protected getSelProxy(): ISelectionProxy {
    this.cursor.markSelectionRead()
    return this.session.getSelectionProxy()
  }

  protected getDocProxy(): IDocumentProxy {
    return this.session.getDocProxy()
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

  protected requireSelection(): void {
    const sel = this.getSelProxy()
    const start = sel.getStart()
    const end = sel.getEnd()
    if (start === end) throw new WordMcpError(
      "No text is selected",
      "NO_SELECTION",
      false,
      "Use word_select_all(), word_select_text(), word_find_text(), or word_select_current({scope:\"word\"}) first."
    )
  }

  protected getWord(): Record<string, unknown> {
    return this.session.comCall(() =>
      this.session.application as Record<string, unknown>
    )
  }

  protected collapseSelection(): void {
    this.cursor.ensureMainBody()
  }

  protected goToEnd(): void {
    this.cursor.goToEnd()
  }

  protected sleep = (ms: number) => new Promise(r => setTimeout(r, ms))
}
