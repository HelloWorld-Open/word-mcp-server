import type { IWordSession } from "./session.js"
import type { DocumentInfo, HeadingEntry } from "./types.js"
import type { PositionMap } from "./position-map.js"
import { WordMcpError } from "../security/errors.js"

export class WordDocument {
  constructor(
    private session: IWordSession,
    private positionMap?: PositionMap,
  ) {}

  private requireDoc(): Record<string, unknown> {
    const doc = this.session.activeDoc ?? (this.session.application as Record<string, unknown>).ActiveDocument as Record<string, unknown> | undefined
    if (!doc) throw new WordMcpError("No document is open", "NO_DOCUMENT", false, "Use word_document(path) to open a file, or word_create to create a new document.")
    return doc
  }

  private getDoc(): Record<string, unknown> {
    return this.requireDoc()
  }

  async getInfo(): Promise<DocumentInfo> {
    const doc = this.getDoc()
    const stat = (n: number): number => {
      try { return (doc.ComputeStatistics as (s: number) => number)(n) as number } catch { return 0 }
    }
    const countOf = (key: string): number => {
      try {
        const col = doc[key] as { Count?: number } | undefined
        return col?.Count ?? 0
      } catch { return 0 }
    }
    return {
      name: doc.Name as string ?? "",
      fullName: doc.FullName as string ?? "",
      path: doc.Path as string ?? "",
      wordCount: stat(0),
      paragraphCount: countOf("Paragraphs"),
      pageCount: stat(2),
      characterCount: stat(3),
      sectionCount: countOf("Sections"),
      saved: (() => { try { return doc.Saved as boolean } catch { return false } })(),
    }
  }

  async getFullText(): Promise<string> {
    const doc = this.getDoc()
    const content = doc.Content as Record<string, unknown>
    const text = (content.Text as string) ?? ""
    return text
  }

  async getParagraphText(index: number): Promise<string> {
    const doc = this.getDoc()
    const paras = doc.Paragraphs as unknown as { Count: number; Item: (i: number) => Record<string, unknown> }
    const total = paras.Count as number
    if (index < 1 || index > total) {
      throw new WordMcpError(
        `Paragraph index ${index} out of range (1-${total})`,
        "PARAGRAPH_INDEX_OUT_OF_RANGE",
        false,
        "Use word_get_info() to check the paragraph count.",
      )
    }
    const p = paras.Item(index)
    return ((p.Range as Record<string, unknown>).Text as string ?? "").replace(/\r?\n$/, "")
  }

  async getTableData(tableIndex: number): Promise<{ tableCount: number; rows: number; columns: number; data: string[][] }> {
    const doc = this.getDoc()
    const tables = doc.Tables as { Count: number; Item: (i: number) => Record<string, unknown> }
    const tableCount = tables.Count as number
    if (tableCount === 0) {
      throw new WordMcpError("No tables exist in the document", "NO_TABLES", false, "Use word_insert_table first.")
    }
    if (tableIndex < 1 || tableIndex > tableCount) {
      throw new WordMcpError(
        `Table index ${tableIndex} out of range (1-${tableCount})`,
        "TABLE_INDEX_OUT_OF_RANGE",
        false,
        `There are ${tableCount} table(s) in the document.`,
      )
    }
    const table = tables.Item(tableIndex)
    const rows = (table.Rows as { Count: number }).Count
    const columns = (table.Columns as { Count: number }).Count
    const raw = (table.Range as Record<string, unknown>).Text as string ?? ""
    const rowsRaw = raw.split(/\r\x07/)
    const data: string[][] = []
    for (let r = 0; r < rowsRaw.length && data.length < rows; r++) {
      const cells = rowsRaw[r].split("\x07")
      const rowData: string[] = []
      for (let c = 0; c < cells.length; c++) {
        const cell = cells[c].replace(/[\r]+$/, "")
        rowData.push(cell)
      }
      if (rowData.length > 0) {
        while (rowData.length < columns) rowData.push("")
        data.push(rowData)
      }
    }
    return { tableCount, rows, columns, data }
  }

  async getComments(): Promise<{ author: string; text: string; index: number }[]> {
    const doc = this.getDoc()
    const comments = doc.Comments as { Count: number; Item: (i: number) => Record<string, unknown> }
    const count = comments.Count as number
    const result: { author: string; text: string; index: number }[] = []
    for (let i = 1; i <= count; i++) {
      const c = comments.Item(i)
      result.push({
        author: c.Author as string,
        text: ((c.Range as Record<string, unknown>).Text as string ?? "").replace(/[\r\x07]+$/, ""),
        index: i,
      })
    }
    return result
  }

  async getBookmarks(): Promise<{ name: string; index: number }[]> {
    const doc = this.getDoc()
    const bookmarks = doc.Bookmarks as { Count: number; Item: (i: number) => Record<string, unknown> }
    const count = bookmarks.Count as number
    const result: { name: string; index: number }[] = []
    for (let i = 1; i <= count; i++) {
      const b = bookmarks.Item(i)
      result.push({ name: b.Name as string, index: i })
    }
    return result
  }

  async getLists(): Promise<{
    listCount: number
    lists: { type: string; items: { level: number; text: string; prefix: string }[] }[]
  }> {
    const doc = this.getDoc()
    const lists = doc.Lists as { Count: number; Item: (i: number) => Record<string, unknown> }
    const listCount = lists.Count as number
    const result: { type: string; items: { level: number; text: string; prefix: string }[] }[] = []
    for (let i = 1; i <= listCount; i++) {
      const list = lists.Item(i)
      const lf = list.Range as Record<string, unknown>
      const lf2 = lf.ListFormat as Record<string, unknown>
      const rawType = lf2.ListType as number
      const type = rawType === 2 ? "bullet" : rawType === 1 ? "numbered" : "mixed"
      const listParas = list.ListParagraphs as { Count: number; Item: (i: number) => Record<string, unknown> }
      const paraCount = listParas.Count as number
      const items: { level: number; text: string; prefix: string }[] = []
      for (let j = 1; j <= paraCount; j++) {
        const p = listParas.Item(j)
        const r = p.Range as Record<string, unknown>
        const lf3 = r.ListFormat as Record<string, unknown>
        const level = (lf3.ListLevelNumber as number) ?? 1
        const prefix = (lf3.ListString as string) ?? ""
        const text = ((r.Text as string) ?? "").replace(/[\r\n]+$/, "")
        items.push({ level, text, prefix })
      }
      result.push({ type, items })
    }
    return { listCount, lists: result }
  }

  async getSections(): Promise<{
    count: number
    sections: {
      index: number; orientation: string; columnCount: number
      pageWidth: number; pageHeight: number
    }[]
  }> {
    const doc = this.getDoc()
    const sections = doc.Sections as { Count: number; Item: (i: number) => Record<string, unknown> } | undefined
    if (!sections) return { count: 0, sections: [] }
    const count = sections.Count as number
    const result: {
      index: number; orientation: string; columnCount: number
      pageWidth: number; pageHeight: number
    }[] = []
    for (let i = 1; i <= count; i++) {
      const s = sections.Item(i)
      const ps = s.PageSetup as Record<string, unknown> | undefined
      const cols = s.Columns as { Count: number } | undefined
      result.push({
        index: i,
        orientation: ps && (ps.Orientation as number) === 1 ? "landscape" : "portrait",
        columnCount: cols ? cols.Count : 1,
        pageWidth: ps ? ps.PageWidth as number : 0,
        pageHeight: ps ? ps.PageHeight as number : 0,
      })
    }
    return { count, sections: result }
  }

  async exportToPdf(outputPath: string): Promise<void> {
    const doc = this.getDoc()
    try {
      ;(doc.ExportAsFixedFormat as (path: string, format: number) => void)(outputPath, 17)
    } catch (e) {
      throw new WordMcpError(
        `Failed to export PDF: ${e instanceof Error ? e.message : String(e)}`,
        "PDF_EXPORT_FAILED",
        true,
        "Ensure the output path is writable and not open in another application.",
      )
    }
  }

  async getStructure(): Promise<{ headings: HeadingEntry[]; totalParagraphs: number }> {
    if (this.positionMap) {
      await this.positionMap.ensureFresh()
      const headings = this.positionMap.getHeadings()
      const doc = this.getDoc()
      const totalParagraphs = (doc.Paragraphs as { Count: number }).Count as number
      return { headings, totalParagraphs }
    }

    const doc = this.getDoc()
    const totalParagraphs = (doc.Paragraphs as { Count: number }).Count as number
    const fullText = (doc.Content as Record<string, unknown>).Text as string
    const rawTexts = fullText.split('\r')
    if (rawTexts.length > 0 && rawTexts[rawTexts.length - 1] === '') rawTexts.pop()
    const allTexts = rawTexts.slice(0, totalParagraphs)
    while (allTexts.length < totalParagraphs) allTexts.push('')

    const paraStarts = new Array(totalParagraphs + 2)
    let textPos = 0
    for (let i = 1; i <= totalParagraphs; i++) {
      paraStarts[i] = textPos
      textPos += allTexts[i - 1].length + 1
    }
    paraStarts[totalParagraphs + 1] = fullText.length

    const binarySearchPara = (startPos: number): number => {
      let lo = 1, hi = totalParagraphs
      while (lo <= hi) {
        const mid = Math.floor((lo + hi) / 2)
        if (startPos >= paraStarts[mid]) {
          if (mid >= totalParagraphs || startPos < paraStarts[mid + 1]) return mid
          lo = mid + 1
        } else {
          hi = mid - 1
        }
      }
      return Math.min(totalParagraphs, Math.max(1, lo))
    }

    const sel = (doc.Application as Record<string, unknown>).Selection as Record<string, unknown>
    const origStart = sel.Start as number
    const origEnd = sel.End as number
    const find = sel.Find as Record<string, unknown>

    const headingEntries: Array<{ pi: number; level: number }> = []
    const headingSet = new Set<number>()

    const styleFormats: Array<(l: number) => string> = [
      (l) => `Heading ${l}`,
      (l) => `标题 ${l}`,
    ]

    for (const fmt of styleFormats) {
      for (let level = 1; level <= 9; level++) {
        ;(find.ClearFormatting as () => void)()
        try {
          find.Style = fmt(level)
        } catch { continue }
        find.Text = ""
        find.Forward = true
        find.Wrap = 0

        ;(sel.HomeKey as (u: number) => void)(6)

        while (true) {
          const found = (find.Execute as (...a: unknown[]) => boolean)(
            "", false, false, false, false, false, true, 0, true, "", 0
          )
          if (!found) break
          const start = sel.Start as number
          const pi = binarySearchPara(start)
          if (!headingSet.has(pi) && pi >= 1 && pi <= totalParagraphs) {
            headingSet.add(pi)
            headingEntries.push({ pi, level })
          }
        }
      }
    }

    try {
      const restore = (doc.Range as (s: number, e: number) => Record<string, unknown>)(origStart, origEnd)
      ;(restore.Select as () => void)()
    } catch { /* ignore */ }

    headingEntries.sort((a, b) => a.pi - b.pi)
    const headings: HeadingEntry[] = []
    for (const e of headingEntries) {
      const text = (allTexts[e.pi - 1] ?? "").replace(/\r?\n$/, "")
      headings.push({ text, level: e.level, paragraphIndex: e.pi })
    }

    return { headings, totalParagraphs }
  }
}
