import type { IWordSession } from "./session.js"
import { WordMcpError } from "../security/errors.js"

export class WordFormatting {
  private static readonly ALIGNMENT: Record<string, number> = {
    left: 0, center: 1, right: 2, justify: 3,
  }
  private static readonly ORIENTATION: Record<string, number> = {
    portrait: 0, landscape: 1,
  }
  private static readonly LINE_SPACING_RULE: Record<string, number> = {
    single: 0, one_point_five: 1, double: 2, at_least: 3, exactly: 4, multiple: 5,
  }
  private static readonly UNDERLINE: Record<string, number> = {
    none: 0, single: 1, double: 3, wavy: 11,
  }
  private static readonly COLOR_INDEX: Record<string, number> = {
    auto: 0, black: 1, blue: 2, turquoise: 3, bright_green: 4, pink: 5,
    red: 6, yellow: 7, white: 8, dark_blue: 9, teal: 10, green: 11,
    violet: 12, dark_red: 13, dark_yellow: 14, gray_50: 15, gray_25: 16,
  }

  constructor(private session: IWordSession) {}

  private requireDoc(): Record<string, unknown> {
    const doc = this.session.activeDoc ?? (this.session.application as Record<string, unknown>).ActiveDocument as Record<string, unknown> | undefined
    if (!doc) throw new WordMcpError("No document is open", "NO_DOCUMENT", false, "Use word_document(path) to open a file, or word_create to create a new document.")
    return doc
  }

  private getSelection(): Record<string, unknown> {
    this.requireDoc()
    return (this.session.application as Record<string, unknown>).Selection as Record<string, unknown>
  }

  private getDoc(): Record<string, unknown> {
    return this.requireDoc()
  }

  private numOrEnum<T>(val: unknown, map: Record<string, T>): T | number {
    if (typeof val === "string") return map[val] ?? (val as unknown as T)
    return val as number
  }

  async setFont(params: Record<string, unknown>): Promise<void> {
    const font = this.getSelection().Font as Record<string, unknown>
    if (params.name != null) font.Name = params.name
    if (params.size != null) font.Size = params.size
    if (params.bold != null) font.Bold = !!params.bold
    if (params.italic != null) font.Italic = !!params.italic
    if (params.underline != null) font.Underline = this.numOrEnum(params.underline, WordFormatting.UNDERLINE)
    if (params.color != null) font.ColorIndex = this.numOrEnum(params.color, WordFormatting.COLOR_INDEX)
    if (params.strikethrough != null) font.Strikethrough = params.strikethrough ? 1 : 0
    if (params.highlightColor != null) font.HighlightColorIndex = this.numOrEnum(params.highlightColor, WordFormatting.COLOR_INDEX)
    if (params.superscript != null) font.Superscript = params.superscript ? 1 : 0
    if (params.subscript != null) font.Subscript = params.subscript ? 1 : 0
  }

  async setParagraphFormat(params: Record<string, unknown>): Promise<void> {
    const pf = this.getSelection().ParagraphFormat as Record<string, unknown>
    if (params.alignment != null) pf.Alignment = this.numOrEnum(params.alignment, WordFormatting.ALIGNMENT)
    if (params.leftIndent != null) pf.LeftIndent = this.cmToPoints(params.leftIndent as number)
    if (params.rightIndent != null) pf.RightIndent = this.cmToPoints(params.rightIndent as number)
    if (params.firstLineIndent != null) pf.FirstLineIndent = this.cmToPoints(params.firstLineIndent as number)
    if (params.spaceBefore != null) pf.SpaceBefore = params.spaceBefore
    if (params.spaceAfter != null) pf.SpaceAfter = params.spaceAfter
    if (params.lineSpacing != null) pf.LineSpacing = params.lineSpacing
    if (params.lineSpacingRule != null) pf.LineSpacingRule = this.numOrEnum(params.lineSpacingRule, WordFormatting.LINE_SPACING_RULE)
  }

  async applyStyle(styleName: string): Promise<void> {
    const sel = this.getSelection()
    const names = this.styleCandidates(styleName)
    for (const n of names) {
      try { sel.Style = n; return } catch { /* try next */ }
    }
    throw new WordMcpError(`Style not found: ${styleName}`, "STYLE_NOT_FOUND", false, `Available styles: try word_list_styles() to see all styles in the current document.`)
  }

  private styleCandidates(name: string): string[] {
    const candidates = [name]
    const cn = name.match(/^标题\s+(\d+)$/)
    if (cn) candidates.push(`Heading ${cn[1]}`)
    const en = name.match(/^(?:Heading|heading)\s+(\d+)$/)
    if (en) candidates.push(`标题 ${en[1]}`)
    return [...new Set(candidates)]
  }

  private cmToPoints(cm: number): number {
    return cm * 28.3465
  }

  async setPageSetup(params: Record<string, unknown>): Promise<void> {
    const doc = this.getDoc()
    const sections = doc.Sections as { Count: number; Item: (i: number) => Record<string, unknown> }
    const si = sections.Count
    const ps = sections.Item(si).PageSetup as Record<string, unknown>
    if (params.topMargin != null) ps.TopMargin = this.cmToPoints(params.topMargin as number)
    if (params.bottomMargin != null) ps.BottomMargin = this.cmToPoints(params.bottomMargin as number)
    if (params.leftMargin != null) ps.LeftMargin = this.cmToPoints(params.leftMargin as number)
    if (params.rightMargin != null) ps.RightMargin = this.cmToPoints(params.rightMargin as number)
    if (params.orientation != null) ps.Orientation = this.numOrEnum(params.orientation, WordFormatting.ORIENTATION)
    if (params.pageWidth != null) ps.PageWidth = this.cmToPoints(params.pageWidth as number)
    if (params.pageHeight != null) ps.PageHeight = this.cmToPoints(params.pageHeight as number)
  }

  async setDocumentProperties(params: Record<string, unknown>): Promise<void> {
    const doc = this.getDoc()
    try {
      const props = doc.BuiltInDocumentProperties as { Item: (n: string) => Record<string, unknown> }
      if (params.title != null) props.Item("Title").Value = params.title
      if (params.author != null) props.Item("Author").Value = params.author
      if (params.subject != null) props.Item("Subject").Value = params.subject
      if (params.keywords != null) props.Item("Keywords").Value = params.keywords
      if (params.comments != null) props.Item("Comments").Value = params.comments
      if (params.category != null) props.Item("Category").Value = params.category
    } catch {
      // Suppress property errors on some document formats
    }
  }

  async setTrackChanges(enable: boolean): Promise<void> {
    const doc = this.getDoc()
    doc.TrackRevisions = enable
  }

  async acceptAllChanges(): Promise<number> {
    const doc = this.getDoc()
    const revisions = doc.Revisions as { Count: number; AcceptAll: () => void }
    const count = revisions.Count
    if (count > 0) revisions.AcceptAll()
    return count
  }

  async rejectAllChanges(): Promise<number> {
    const doc = this.getDoc()
    const revisions = doc.Revisions as { Count: number; RejectAll: () => void }
    const count = revisions.Count
    if (count > 0) revisions.RejectAll()
    return count
  }

  async listStyles(): Promise<Array<{ name: string; type: number; builtIn: boolean }>> {
    const doc = this.getDoc()
    const styles = doc.Styles as unknown as ArrayLike<Record<string, unknown>>
    const result: Array<{ name: string; type: number; builtIn: boolean }> = []
    for (let i = 1; i <= (styles as unknown as { Count: number }).Count; i++) {
      const style = (styles as unknown as { Item: (i: number) => Record<string, unknown> }).Item(i)
      if (style.InUse) {
        result.push({
          name: style.NameLocal as string,
          type: style.Type as number,
          builtIn: style.BuiltIn as boolean,
        })
      }
    }
    return result
  }
}
