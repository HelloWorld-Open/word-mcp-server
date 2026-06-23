import type { IWordSession } from "./session.js"
import type { ISelectionProxy, IDocumentProxy, IRangeProxy } from "./com-proxy/types.js"
import { WordMcpError } from "../security/errors.js"

export interface ICursorContext {
  ensureMainBody(): void
  markInBody(): void
  markSelectionRead(): void
  goToEnd(): void
  reset(): void
}

export class ContextSanitizer implements ICursorContext {
  private static readonly WD_WITHIN_TABLE = 12

  private wasInNonBody = false
  private cachedStart = -1
  private cachedEnd = -1
  private collapseReady = false

  constructor(private session: IWordSession) {}

  // ===================== Static text sanitizers =====================

  static sanitizeText(text: string): string {
    return text.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, "")
  }

  static stripBel(text: string): string {
    return text.replace(/[\r\x07]+$/, "")
  }

  // ===================== Cursor context management =====================

  ensureMainBody(): void {
    try {
      if (this.collapseReady) return

      const sel = this.getSelection()
      const storyType = sel.getStoryType()
      if (storyType === 1) {
        const start = sel.getStart()
        const end = sel.getEnd()
        if (start === this.cachedStart && end === this.cachedEnd) {
          this.collapseReady = true
          return
        }
        this.cachedStart = start
        this.cachedEnd = end
      }

      const doc = this.session.activeDoc ?? (this.getWord().ActiveDocument as Record<string, unknown>)
      if (!doc) return

      if (storyType !== 1) {
        const contentEnd = this.session.getDocProxy().getContent().getEnd()
        const endRange = this.session.getDocProxy().getRange(contentEnd, contentEnd)
        endRange.select()
        this.getSelection().collapse(0)
        this.wasInNonBody = true
        return
      }

      if (this.wasInNonBody) {
        sel.typeParagraph()
        this.wasInNonBody = false
      }

      try {
        if (sel.getInformation(ContextSanitizer.WD_WITHIN_TABLE)) {
          const table = (sel.getTables() as { Item: (i: number) => Record<string, unknown> }).Item(1)
          // Use Document.Range to position cursor AFTER the table.
          // Range.Next() guarantees the cursor lands outside table structure.
          // Fallback: InsertParagraph at document end creates new paragraph outside table.
          const rawTableRange = table.Range as Record<string, unknown>
          const nextRange = (rawTableRange.Next as () => Record<string, unknown> | undefined)()
          if (nextRange) {
            const nextStart = nextRange.Start as number
            if (typeof nextStart === "number") {
              this.session.getDocProxy().getRange(nextStart, nextStart).select()
              this.getSelection().collapse(0)
            } else {
              throw new Error("no next range start")
            }
          } else {
            throw new Error("no content after table")
          }
          this.cachedStart = -1
          this.cachedEnd = -1
          this.collapseReady = false
        }
      } catch { /* table check may fail */ }

      try {
        const shapes = sel.getShapeRange() as { Count: number } | undefined
        if (shapes && shapes.Count > 0) {
          const contentEnd = this.session.getDocProxy().getContent().getEnd()
          const endRange = this.session.getDocProxy().getRange(contentEnd, contentEnd)
          endRange.select()
          this.getSelection().collapse(0)
        }
      } catch { /* shape check may fail */ }
    } catch (e) { this.session.logger?.warn({ err: e }, "ensureMainBody failed") }
  }

  markInBody(): void {
    this.wasInNonBody = false
    this.collapseReady = false
  }

  markSelectionRead(): void {
    this.collapseReady = false
  }

  goToEnd(): void {
    try {
      if (!this.session.activeDoc) return
      const doc = this.session.getDocProxy()
      const end = doc.getContent().getEnd()
      doc.getRange(end, end).select()
    } catch (e) { this.session.logger?.warn({ err: e }, "goToEnd failed") }
  }

  reset(): void {
    this.cachedStart = -1
    this.cachedEnd = -1
    this.collapseReady = false
    this.wasInNonBody = false
  }

  // ===================== Private COM helpers =====================

  private getSelection(): ISelectionProxy {
    this.collapseReady = false
    const sel = this.session.getSelectionProxy()
    if (!sel) {
      throw new WordMcpError("COM Selection proxy returned null — Word COM connection may be transiently unavailable", "COM_SELECTION_NULL", true, "The operation will be retried automatically. If the problem persists, close Word and try again.")
    }
    return sel
  }

  private getWord(): Record<string, unknown> {
    return this.session.comCall(() =>
      this.session.application as Record<string, unknown>
    )
  }
}
