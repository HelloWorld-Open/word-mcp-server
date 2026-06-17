import type { IRangeProxy } from "./types.js"

export class RangeProxy implements IRangeProxy {
  constructor(
    private rawRange: Record<string, unknown>,
    private session: { wrapRange: (raw: Record<string, unknown>) => IRangeProxy; log?: (level: string, msg: string) => void },
  ) {}

  get raw(): Record<string, unknown> {
    return this.rawRange
  }

  getText(): string {
    try { return this.rawRange.Text as string } catch { this.session.log?.("debug", "RangeProxy.getText failed"); return "" }
  }

  setText(val: string): void {
    try { this.rawRange.Text = val } catch { this.session.log?.("warn", "RangeProxy op failed") }
  }

  getStart(): number {
    try { return this.rawRange.Start as number } catch { this.session.log?.("debug", "RangeProxy.getStart/End failed"); return 0 }
  }

  getEnd(): number {
    try { return this.rawRange.End as number } catch { this.session.log?.("debug", "RangeProxy.getStart/End failed"); return 0 }
  }

  setStart(v: number): void {
    try { this.rawRange.Start = v } catch { this.session.log?.("warn", "RangeProxy op failed") }
  }

  setEnd(v: number): void {
    try { this.rawRange.End = v } catch { this.session.log?.("warn", "RangeProxy op failed") }
  }

  setRange(start: number, end: number): void {
    try { ;(this.rawRange.SetRange as (s: number, e: number) => void)(start, end) } catch { this.session.log?.("warn", "RangeProxy op failed") }
  }

  getBold(): boolean | undefined {
    try { return this.rawRange.Bold as boolean } catch { this.session.log?.("debug", "RangeProxy.getBold/Italic failed"); return undefined }
  }

  setBold(val: boolean): void {
    try { this.rawRange.Bold = val } catch { this.session.log?.("warn", "RangeProxy op failed") }
  }

  getItalic(): boolean | undefined {
    try { return this.rawRange.Italic as boolean } catch { this.session.log?.("debug", "RangeProxy.getBold/Italic failed"); return undefined }
  }

  setItalic(val: boolean): void {
    try { this.rawRange.Italic = val } catch { this.session.log?.("warn", "RangeProxy op failed") }
  }

  select(): void {
    try { ;(this.rawRange.Select as () => void)() } catch { this.session.log?.("warn", "RangeProxy op failed") }
  }

  duplicate(): IRangeProxy {
    try {
      const raw = this.rawRange.Duplicate as Record<string, unknown>
      return this.session.wrapRange(raw)
    } catch { this.session.log?.("debug", "RangeProxy.duplicate failed")
      return this.session.wrapRange({} as Record<string, unknown>)
    }
  }

  getFind(): Record<string, unknown> {
    return this.rawRange.Find as Record<string, unknown>
  }

  getFont(): Record<string, unknown> {
    return this.rawRange.Font as Record<string, unknown>
  }

  getParagraphFormat(): Record<string, unknown> {
    return this.rawRange.ParagraphFormat as Record<string, unknown>
  }

  getShading(): Record<string, unknown> {
    return this.rawRange.Shading as Record<string, unknown>
  }

  getListFormat(): Record<string, unknown> {
    return this.rawRange.ListFormat as Record<string, unknown>
  }

  getHyperlinks(): Record<string, unknown> {
    return this.rawRange.Hyperlinks as Record<string, unknown>
  }

  insertFile(path: string): void {
    try { ;(this.rawRange.InsertFile as (p: string) => void)(path) } catch { this.session.log?.("warn", "RangeProxy op failed") }
  }

  convertToTable(separator: string): Record<string, unknown> {
    try {
      return (this.rawRange.ConvertToTable as (s: string) => Record<string, unknown>)(separator)
    } catch { this.session.log?.("warn", "RangeProxy.convertToTable failed")
      return {}
    }
  }

  addField(type: number): void {
    try {
      ;(this.rawRange.Fields as { Add: (r: unknown, t: number) => void }).Add(this.rawRange, type)
    } catch { this.session.log?.("warn", "RangeProxy.addField failed") }
  }
}
