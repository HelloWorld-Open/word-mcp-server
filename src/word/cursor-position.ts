import type { IWordSession } from "./session.js"

export class CursorPosition {
  private wasInNonBody = false
  private cachedStart = -1
  private cachedEnd = -1
  private collapseReady = false

  constructor(private session: IWordSession) {}

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
    } catch { }
  }

  reset(): void {
    this.cachedStart = -1
    this.cachedEnd = -1
    this.collapseReady = false
    this.wasInNonBody = false
  }

  private getSelection(): Record<string, unknown> {
    this.collapseReady = false
    return this.session.comCall(() =>
      (this.getWord().Selection as Record<string, unknown>) as Record<string, unknown>
    )
  }

  private getWord(): Record<string, unknown> {
    return this.session.comCall(() =>
      this.session.application as Record<string, unknown>
    )
  }
}
