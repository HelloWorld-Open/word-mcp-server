import type { IWordSession } from "./session.js"
import { EXPORT_FORMAT_PDF, type DocumentInfo, type HeadingEntry } from "./types.js"
import type { PositionMap } from "./position-map.js"
import { WordMcpError } from "../security/errors.js"
import { ContextSanitizer } from "./context-sanitizer.js"
import type { IDocumentProxy } from "./com-proxy/types.js"

export class WordDocument {
  constructor(
    private session: IWordSession,
    private positionMap: PositionMap,
  ) {}

  private requireDoc(): IDocumentProxy {
    const doc = this.session.getDocProxy()
    if (!doc) throw new WordMcpError("No document is open", "NO_DOCUMENT", false, "Use word_document(path) to open a file, or word_stream_start to create a new document.")
    return doc
  }

  private getDoc(): IDocumentProxy {
    return this.requireDoc()
  }


  getInfo(): DocumentInfo {
    const doc = this.getDoc()
    const stat = (n: number): number => {
      try { return doc.computeStatistics(n) as number } catch { return 0 }
    }
    const countOf = (key: string): number => {
      try {
        const col = doc.raw[key] as { Count?: number } | undefined
        return col?.Count ?? 0
      } catch { return 0 }
    }
    return {
      name: doc.getName() ?? "",
      fullName: doc.getFullName() ?? "",
      path: doc.getPath() ?? "",
      wordCount: stat(0),
      paragraphCount: countOf("Paragraphs"),
      pageCount: stat(2),
      characterCount: stat(3),
      sectionCount: countOf("Sections"),
      saved: (() => { try { return doc.getSaved() } catch { return false } })(),
    }
  }

  getFullText(): string {
    const doc = this.getDoc()
    const content = doc.getContent()
    const text = content.getText() ?? ""
    return text
  }

  getParagraphText(index: number): string {
    const doc = this.getDoc()
    const paras = doc.getParagraphs()
    const total = paras.count
    if (index < 1 || index > total) {
      throw new WordMcpError(
        `Paragraph index ${index} out of range (1-${total})`,
        "PARAGRAPH_INDEX_OUT_OF_RANGE",
        false,
        "Use word_get_info() to check the paragraph count.",
      )
    }
    const p = paras.item(index)
    return ContextSanitizer.stripBel((p.Range as Record<string, unknown>).Text as string ?? "")
  }

  getTableData(tableIndex: number): { tableCount: number; rows: number; columns: number; data: string[][] } {
    const doc = this.getDoc()
    const tables = doc.getTables()
    const tableCount = tables.count
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
    const table = tables.item(tableIndex)
    const rows = (table.Rows as { Count: number }).Count
    const columns = (table.Columns as { Count: number }).Count
    const data: string[][] = []
    for (let r = 1; r <= rows; r++) {
      const rowData: string[] = []
      for (let c = 1; c <= columns; c++) {
        try {
          const cellText = ((table.Cell as (r: number, c: number) => Record<string, unknown>)(r, c).Range as Record<string, unknown>).Text as string ?? ""
          rowData.push(ContextSanitizer.stripBel(cellText))
        } catch {
          rowData.push("")
        }
      }
      data.push(rowData)
    }
    return { tableCount, rows, columns, data }
  }

  getComments(): { author: string; text: string; index: number }[] {
    const doc = this.getDoc()
    const comments = doc.getComments()
    const count = comments.count
    const result: { author: string; text: string; index: number }[] = []
    for (let i = 1; i <= count; i++) {
      const c = comments.item(i)
      result.push({
        author: c.Author as string,
        text: ((c.Range as Record<string, unknown>).Text as string ?? "").replace(/[\r\x07]+$/, ""),
        index: i,
      })
    }
    return result
  }

  getBookmarks(): { name: string; index: number }[] {
    const doc = this.getDoc()
    const bookmarks = doc.getBookmarks()
    const count = bookmarks.count
    const result: { name: string; index: number }[] = []
    for (let i = 1; i <= count; i++) {
      const b = bookmarks.item(i)
      result.push({ name: b.Name as string, index: i })
    }
    return result
  }

  getLists(): {
    listCount: number
    lists: { type: string; items: { level: number; text: string; prefix: string }[] }[]
  } {
    const doc = this.getDoc()
    const lists = doc.getLists() as { Count: number; Item: (i: number) => Record<string, unknown> }
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

  getSections(): {
    count: number
    sections: {
      index: number; orientation: string; columnCount: number
      pageWidth: number; pageHeight: number
    }[]
  } {
    const doc = this.getDoc()
    const sections = doc.getSections()
    if (!sections) return { count: 0, sections: [] }
    const count = sections.count
    const result: {
      index: number; orientation: string; columnCount: number
      pageWidth: number; pageHeight: number
    }[] = []
    for (let i = 1; i <= count; i++) {
      const s = sections.item(i)
      const ps = s.getPageSetup()
      const cols = s.raw.Columns as { Count: number } | undefined
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

  exportToPdf(outputPath: string): void {
    const doc = this.getDoc()
    try {
      doc.exportAsFixedFormat(outputPath, EXPORT_FORMAT_PDF)
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
    await this.positionMap.ensureFresh()
    const headings = this.positionMap.getHeadings()
    const doc = this.getDoc()
    const totalParagraphs = doc.getParagraphs().count
    return { headings, totalParagraphs }
  }
}
