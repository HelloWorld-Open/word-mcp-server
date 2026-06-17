import type { ISelectionProxy, IRangeProxy } from "./types.js"

export class SelectionProxy implements ISelectionProxy {
  constructor(
    private rawSel: Record<string, unknown>,
    private session: { wrapRange: (raw: Record<string, unknown>) => IRangeProxy; log?: (level: string, msg: string) => void },
  ) {}

  get raw(): Record<string, unknown> {
    return this.rawSel
  }

  getStart(): number {
    try { return this.rawSel.Start as number } catch { this.session.log?.("debug", "SelectionProxy.getStart failed"); return 0 }
  }

  getEnd(): number {
    try { return this.rawSel.End as number } catch { this.session.log?.("debug", "SelectionProxy.getEnd failed"); return 0 }
  }

  getStoryType(): number {
    try { return this.rawSel.StoryType as number } catch { this.session.log?.("debug", "SelectionProxy.getStoryType failed"); return 1 }
  }

  getStyle(): string | number {
    try { return this.rawSel.Style as string | number } catch { this.session.log?.("debug", "SelectionProxy.getStyle failed"); return "" }
  }

  setStyle(v: string | number): void {
    try { (this.rawSel as Record<string, unknown>).Style = v } catch { this.session.log?.("warn", "SelectionProxy.setStyle failed") }
  }

  getType(): number {
    try { return this.rawSel.Type as number } catch { this.session.log?.("debug", "SelectionProxy.getType failed"); return 1 }
  }

  typeText(text: string): void {
    try { ;(this.rawSel.TypeText as (t: string) => void)(text) } catch { this.session.log?.("error", "SelectionProxy.typeText failed") }
  }

  typeParagraph(): void {
    try { ;(this.rawSel.TypeParagraph as () => void)() } catch { this.session.log?.("error", "SelectionProxy.typeParagraph failed") }
  }

  typeBackspace(): void {
    try { ;(this.rawSel.TypeBackspace as () => void)() } catch { this.session.log?.("warn", "SelectionProxy op failed") }
  }

  collapse(direction?: number): void {
    try { ;(this.rawSel.Collapse as (d?: number) => void)(direction ?? 0) } catch { this.session.log?.("warn", "SelectionProxy op failed") }
  }

  /** CAUTION: EndKey(wdStory) may not move past a table when the cursor is inside one.
   *  Use Document.Range(tableEnd, tableEnd).select() for reliable cursor movement past tables. */
  endKey(unit: number): void {
    try { ;(this.rawSel.EndKey as (u: number) => void)(unit) } catch { this.session.log?.("warn", "SelectionProxy op failed") }
  }

  moveStart(unit: number, count?: number): void {
    try { ;(this.rawSel.MoveStart as (u: number, c?: number) => void)(unit, count ?? 1) } catch { this.session.log?.("warn", "SelectionProxy op failed") }
  }

  homeKey(unit: number): void {
    try { ;(this.rawSel.HomeKey as (u: number) => void)(unit) } catch { this.session.log?.("warn", "SelectionProxy op failed") }
  }

  wholeStory(): void {
    try { ;(this.rawSel.WholeStory as () => void)() } catch { this.session.log?.("warn", "SelectionProxy op failed") }
  }

  delete(): void {
    try { ;(this.rawSel.Delete as () => void)() } catch { this.session.log?.("warn", "SelectionProxy op failed") }
  }

  copy(): void {
    try { ;(this.rawSel.Copy as () => void)() } catch { this.session.log?.("warn", "SelectionProxy op failed") }
  }

  cut(): void {
    try { ;(this.rawSel.Cut as () => void)() } catch { this.session.log?.("warn", "SelectionProxy op failed") }
  }

  paste(): void {
    try { ;(this.rawSel.Paste as () => void)() } catch { this.session.log?.("warn", "SelectionProxy op failed") }
  }

  select(): void {
    try { ;(this.rawSel.Select as () => void)() } catch { this.session.log?.("warn", "SelectionProxy op failed") }
  }

  expand(unit: number): void {
    try { ;(this.rawSel.Expand as (u: number) => void)(unit) } catch { this.session.log?.("warn", "SelectionProxy op failed") }
  }

  goTo(what: number, which?: number, count?: number): Record<string, unknown> {
    try {
      return (this.rawSel.GoTo as (w: number, wh?: number, c?: number) => Record<string, unknown>)(what, which ?? 1, count ?? 1)
    } catch { this.session.log?.("warn", "SelectionProxy.goTo failed")
      return {}
    }
  }

  insertBreak(type: number): void {
    try { ;(this.rawSel.InsertBreak as (t: number) => void)(type) } catch { this.session.log?.("error", "SelectionProxy.insertBreak failed") }
  }

  getRange(): IRangeProxy {
    try {
      const raw = this.rawSel.Range as Record<string, unknown>
      return this.session.wrapRange(raw)
    } catch { this.session.log?.("debug", "SelectionProxy.getRange failed")
      return this.session.wrapRange({} as Record<string, unknown>)
    }
  }

  getInformation(type: number): number | boolean {
    try {
      return (this.rawSel.Information as (t: number) => number | boolean)(type)
    } catch { this.session.log?.("debug", "SelectionProxy.getInformation failed")
      return false
    }
  }

  addHorizontalLine(): void {
    try {
      ;((this.rawSel.InlineShapes as Record<string, unknown>).AddHorizontalLineStandard as () => void)()
    } catch { this.session.log?.("warn", "SelectionProxy op failed") }
  }

  getFind(): Record<string, unknown> {
    try { return this.rawSel.Find as Record<string, unknown> } catch { this.session.log?.("debug", "SelectionProxy.getFind failed"); return {} }
  }

  getFont(): Record<string, unknown> {
    try { return this.rawSel.Font as Record<string, unknown> } catch { this.session.log?.("debug", "SelectionProxy.getFont failed"); return {} }
  }

  getParagraphFormat(): Record<string, unknown> {
    try { return this.rawSel.ParagraphFormat as Record<string, unknown> } catch { this.session.log?.("debug", "SelectionProxy.getParagraphFormat failed"); return {} }
  }

  getInlineShapes(): Record<string, unknown> {
    try { return this.rawSel.InlineShapes as Record<string, unknown> } catch { this.session.log?.("debug", "SelectionProxy.getInlineShapes failed"); return {} }
  }

  getTables(): Record<string, unknown> {
    try { return this.rawSel.Tables as Record<string, unknown> } catch { this.session.log?.("debug", "SelectionProxy.getTables failed"); return {} }
  }

  getShapeRange(): Record<string, unknown> {
    try { return this.rawSel.ShapeRange as Record<string, unknown> } catch { this.session.log?.("debug", "SelectionProxy.getShapeRange failed"); return {} }
  }
}
