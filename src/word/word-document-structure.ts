import { WordBase } from "./word-base.js"

export class WordDocumentStructure extends WordBase {
  private static readonly ALIGNMENT: Record<string, number> = {
    left: 0, center: 1, right: 2,
  }
  private static readonly COLOR_RGB: Record<string, number> = {
    auto: 0, black: 0, blue: 0x0000FF, turquoise: 0x00FFFF, bright_green: 0x00FF00, pink: 0xFF00FF,
    red: 0xFF0000, yellow: 0xFFFF00, white: 0xFFFFFF, dark_blue: 0x000080, teal: 0x008080, green: 0x008000,
    violet: 0x800080, dark_red: 0x800000, dark_yellow: 0x808000, gray_50: 0x808080, gray_25: 0xC0C0C0,
  }

  private getLastSectionIndex(): number {
    const doc = this.requireDoc()
    return (doc.Sections as { Count: number }).Count
  }

  private restoreMainDocCursor(): void {
    try {
      const doc = this.requireDoc()
      const contentEnd = (doc.Content as Record<string, unknown>).End as number
      const range = (doc.Range as (s: number, e: number) => Record<string, unknown>)(contentEnd, contentEnd)
      ;(range.Select as () => void)()
      ;((this.getSelection()).Collapse as (d: number) => void)(0)
      this.session.wasInNonBody = false
    } catch { /* ignore */ }
  }

  async setHeader(text: string, alignment?: unknown): Promise<void> {
    const doc = this.requireDoc()
    const si = this.getLastSectionIndex()
    const section = (doc.Sections as { Item: (i: number) => Record<string, unknown> }).Item(si)
    const headerObj = (section.Headers as { Item: (i: number) => Record<string, unknown> }).Item(1)
    const hdrRange = headerObj.Range as Record<string, unknown>
    hdrRange.Text = text
    if (alignment != null) {
      ;(hdrRange.ParagraphFormat as Record<string, unknown>).Alignment = this.numOrEnum(alignment, WordDocumentStructure.ALIGNMENT)
    }
    this.restoreMainDocCursor()
  }

  async setFooter(text: string, alignment?: unknown): Promise<void> {
    const doc = this.requireDoc()
    const si = this.getLastSectionIndex()
    const section = (doc.Sections as { Item: (i: number) => Record<string, unknown> }).Item(si)
    const footerObj = (section.Footers as { Item: (i: number) => Record<string, unknown> }).Item(1)
    const ftrRange = footerObj.Range as Record<string, unknown>
    ftrRange.Text = text
    if (alignment != null) {
      ;(ftrRange.ParagraphFormat as Record<string, unknown>).Alignment = this.numOrEnum(alignment, WordDocumentStructure.ALIGNMENT)
    }
    this.restoreMainDocCursor()
  }

  async setPageNumbers(target: "header" | "footer"): Promise<void> {
    const doc = this.requireDoc()
    const si = this.getLastSectionIndex()
    const section = (doc.Sections as { Item: (i: number) => Record<string, unknown> }).Item(si)
    const container = target === "header"
      ? (section.Headers as { Item: (i: number) => Record<string, unknown> }).Item(1)
      : (section.Footers as { Item: (i: number) => Record<string, unknown> }).Item(1)
    const range = container.Range as Record<string, unknown>
    range.Text = ""
    ;(range.ParagraphFormat as Record<string, unknown>).Alignment = 1
    const fields = range.Fields as { Add: (r: unknown, type: number) => void }
    fields.Add(range, 33)
    this.restoreMainDocCursor()
  }

  async insertToc(): Promise<void> {
    this.collapseSelection()
    const doc = this.requireDoc()
    ;(doc.TablesOfContents as { Add: (r: unknown) => void }).Add(this.getSelection().Range)
    const sel = this.getSelection()
    ;(sel.EndKey as (u: number) => void)(6)
    ;(sel.TypeParagraph as () => void)()
  }

  async addBookmark(name: string): Promise<void> {
    this.collapseSelection()
    const doc = this.requireDoc()
    ;(doc.Bookmarks as { Add: (n: string) => void }).Add(name)
    const sel = this.getSelection()
    ;(sel.EndKey as (u: number) => void)(6)
  }

  async addComment(text: string): Promise<void> {
    this.collapseSelection()
    const doc = this.requireDoc()
    ;(doc.Comments as { Add: (r: unknown, t: string) => void }).Add(this.getSelection().Range, text)
    const sel = this.getSelection()
    ;(sel.EndKey as (u: number) => void)(6)
  }

  async setWatermark(params: { text: string; remove?: boolean; fontSize?: number; color?: unknown }): Promise<void> {
    const doc = this.requireDoc()
    const sections = doc.Sections as { Count: number; Item: (i: number) => Record<string, unknown> }
    const addToSection = (si: number) => {
      const hdr = (sections.Item(si).Headers as { Item: (i: number) => Record<string, unknown> }).Item(1)
      if (params.remove) {
        const shapes = hdr.Shapes as { Count: number; Item: (i: number) => Record<string, unknown> }
        for (let i = shapes.Count; i >= 1; i--) { ;(shapes.Item(i).Delete as () => void)() }
        return
      }
      const shape = (hdr.Shapes as { AddTextEffect: (preset: number, text: string, font: string, size: number, bold: boolean, italic: boolean, left: number, top: number) => Record<string, unknown> })
        .AddTextEffect(0, params.text, "Arial", params.fontSize ?? 48, false, false, 0, 0)
      shape.RelativeHorizontalPosition = 0
      shape.RelativeVerticalPosition = 0
      shape.Left = -999995
      shape.Top = -999995
      shape.Rotation = -45
      ;(shape.Fill as Record<string, unknown>).Visible = true
      ;(shape.Line as Record<string, unknown>).Visible = false
      if (params.color != null) {
        const cf = (shape.Fill as Record<string, unknown>).ForeColor as Record<string, unknown>
        cf.RGB = this.numOrEnum(params.color, WordDocumentStructure.COLOR_RGB)
      }
      ;((shape.WrapFormat as Record<string, unknown>).AllowOverlap as boolean) = true
      ;(shape.ZOrder as (a: number) => void)(4)
    }
    if (!params.remove) {
      for (let i = 1; i <= sections.Count; i++) addToSection(i)
    } else {
      if (sections.Count > 0) addToSection(1)
    }
    this.restoreMainDocCursor()
  }
}
