import { WordBase } from "./word-base.js"
import { WordMcpError } from "../security/errors.js"

export interface StyleProfileFont {
  name?: string
  size?: number
  bold?: boolean
  italic?: boolean
  color?: string
}

export interface StyleProfilePara {
  alignment?: string
  firstLineIndent?: number
  spaceBefore?: number
  spaceAfter?: number
  lineSpacing?: number
  lineSpacingRule?: string
}

export interface StyleProfile {
  font?: StyleProfileFont
  paragraph?: StyleProfilePara
}

export class WordFormatting extends WordBase {
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

  async modifyStyle(styleName: string, profile: StyleProfile): Promise<void> {
    const doc = this.requireDoc()
    const styles = doc.Styles as { Item: (n: string) => Record<string, unknown> }
    const candidates = this.styleCandidates(styleName)
    let style: Record<string, unknown> | undefined
    for (const n of candidates) {
      try { style = styles.Item(n); break } catch { continue }
    }
    if (!style) {
      throw new WordMcpError(
        `Built-in style not found: ${styleName}`,
        "STYLE_NOT_FOUND", false,
        "Try word_list_styles() to see available styles in the current document."
      )
    }

    if (profile.font) {
      const font = style.Font as Record<string, unknown>
      if (profile.font.name != null) font.Name = profile.font.name
      if (profile.font.size != null) font.Size = profile.font.size
      if (profile.font.bold != null) font.Bold = profile.font.bold
      if (profile.font.italic != null) font.Italic = profile.font.italic
      if (profile.font.color != null) font.ColorIndex = this.numOrEnum(profile.font.color, WordFormatting.COLOR_INDEX)
    }

    if (profile.paragraph) {
      const pf = style.ParagraphFormat as Record<string, unknown>
      if (profile.paragraph.alignment != null) pf.Alignment = this.numOrEnum(profile.paragraph.alignment, WordFormatting.ALIGNMENT)
      if (profile.paragraph.firstLineIndent != null) pf.FirstLineIndent = this.cmToPoints(profile.paragraph.firstLineIndent)
      if (profile.paragraph.spaceBefore != null) pf.SpaceBefore = profile.paragraph.spaceBefore
      if (profile.paragraph.spaceAfter != null) pf.SpaceAfter = profile.paragraph.spaceAfter
      if (profile.paragraph.lineSpacing != null) pf.LineSpacing = profile.paragraph.lineSpacing
      if (profile.paragraph.lineSpacingRule != null) pf.LineSpacingRule = this.numOrEnum(profile.paragraph.lineSpacingRule, WordFormatting.LINE_SPACING_RULE)
    }
  }

  private cmToPoints(cm: number): number {
    return cm * 28.3465
  }

  async setPageSetup(params: Record<string, unknown>): Promise<void> {
    const doc = this.requireDoc()
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
    const doc = this.requireDoc()
    try {
      const props = doc.BuiltInDocumentProperties as { Item: (n: string) => Record<string, unknown> }
      if (params.title != null) props.Item("Title").Value = params.title
      if (params.author != null) props.Item("Author").Value = params.author
      if (params.subject != null) props.Item("Subject").Value = params.subject
      if (params.keywords != null) props.Item("Keywords").Value = params.keywords
      if (params.comments != null) props.Item("Comments").Value = params.comments
      if (params.category != null) props.Item("Category").Value = params.category
    } catch {
    }
  }

  async applyBodyIndent(indentCm: number): Promise<number> {
    const doc = this.requireDoc()
    const sel = this.getSelection()
    const origStart = sel.Start as number
    const origEnd = sel.End as number
    let count = 0
    try {
      const rng = (doc.Range as (s: number, e: number) => Record<string, unknown>)(
        (doc.Content as Record<string, unknown>).Start as number,
        (doc.Content as Record<string, unknown>).End as number,
      )
      const find = rng.Find as Record<string, unknown>
      const replacement = find.Replacement as Record<string, unknown>
      const replPfmt = replacement.ParagraphFormat as Record<string, unknown>
      ;(find.ClearFormatting as () => void)()
      find.Style = "Normal"
      ;(replacement.ClearFormatting as () => void)()
      replPfmt.FirstLineIndent = this.cmToPoints(indentCm)
      replacement.Text = "^&"
      find.Text = ""
      find.Forward = true
      find.Format = true
      find.Wrap = 1
      ;(find.Execute as (...a: unknown[]) => boolean)("", false, false, false, false, false, true, 1, true, "^&", 2)
      count = (doc.Paragraphs as { Count: number }).Count
    } finally {
      try {
        const restore = (doc.Range as (s: number, e: number) => Record<string, unknown>)(origStart, origEnd)
        ;(restore.Select as () => void)()
      } catch { }
    }
    return count
  }

  async setTrackChanges(enable: boolean): Promise<void> {
    const doc = this.requireDoc()
    doc.TrackRevisions = enable
  }

  async acceptAllChanges(): Promise<number> {
    const doc = this.requireDoc()
    const revisions = doc.Revisions as { Count: number; AcceptAll: () => void }
    const count = revisions.Count
    if (count > 0) revisions.AcceptAll()
    return count
  }

  async rejectAllChanges(): Promise<number> {
    const doc = this.requireDoc()
    const revisions = doc.Revisions as { Count: number; RejectAll: () => void }
    const count = revisions.Count
    if (count > 0) revisions.RejectAll()
    return count
  }

  async listStyles(): Promise<Array<{ name: string; type: number; builtIn: boolean }>> {
    const doc = this.requireDoc()
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
