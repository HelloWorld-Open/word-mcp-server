import type { IWordSession } from "./session.js"

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

  static cleanCellText(text: string): string {
    return text.replace(/[\r\x07]+$/, "")
  }

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
      const storyType = sel.StoryType as number
      if (storyType === 1) {
        const start = sel.Start as number
        const end = sel.End as number
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
        const contentEnd = (doc.Content as Record<string, unknown>).End as number
        const endRange = (doc.Range as (s: number, e: number) => Record<string, unknown>)(contentEnd, contentEnd)
        ;(endRange.Select as () => void)()
        ;((this.getSelection()).Collapse as (d: number) => void)(0)
        this.wasInNonBody = true
        return
      }

      if (this.wasInNonBody) {
        ;(sel.TypeParagraph as () => void)()
        this.wasInNonBody = false
      }

      try {
        if ((sel.Information as (t: number) => boolean)(ContextSanitizer.WD_WITHIN_TABLE)) {
          const table = (sel.Tables as { Item: (i: number) => Record<string, unknown> }).Item(1)
          ;((table.Range as Record<string, unknown>).Select as () => void)()
          ;((this.getSelection()).Collapse as (d: number) => void)(0)
        }
      } catch { /* table check may fail */ }

      try {
        const shapes = sel.ShapeRange as { Count: number } | undefined
        if (shapes && shapes.Count > 0) {
          const contentEnd = (doc.Content as Record<string, unknown>).End as number
          const endRange = (doc.Range as (s: number, e: number) => Record<string, unknown>)(contentEnd, contentEnd)
          ;(endRange.Select as () => void)()
          ;((this.getSelection()).Collapse as (d: number) => void)(0)
        }
      } catch { /* shape check may fail */ }
    } catch { /* ignore */ }
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
      const word = this.session.comCall(() => this.session.application as Record<string, unknown>)
      const doc = this.session.activeDoc ?? (word.ActiveDocument as Record<string, unknown>)
      if (!doc) return
      const end = (doc.Content as Record<string, unknown>).End as number
      const rng = (doc.Range as (s: number, e: number) => Record<string, unknown>)(end, end)
      ;(rng.Select as () => void)()
    } catch { /* ignore */ }
  }

  reset(): void {
    this.cachedStart = -1
    this.cachedEnd = -1
    this.collapseReady = false
    this.wasInNonBody = false
  }

  // ===================== Private COM helpers =====================

  private getSelection(): Record<string, unknown> {
    this.collapseReady = false
    const sel = this.session.comCall(() =>
      (this.getWord().Selection as Record<string, unknown>) as Record<string, unknown>
    ) as Record<string, unknown>
    if (!sel) {
      throw new Error("COM Selection proxy returned null — Word COM connection may be transiently unavailable")
    }
    return sel
  }

  private getWord(): Record<string, unknown> {
    return this.session.comCall(() =>
      this.session.application as Record<string, unknown>
    )
  }
}
