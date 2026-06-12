import type { IWordSession } from "../word/session.js"
import { WordApplicationManager } from "../word/application.js"
import { WordDocument } from "../word/document.js"
import { WordTextEditor } from "../word/word-text-editor.js"
import { WordTableEditor } from "../word/word-table-editor.js"
import { WordMediaEditor } from "../word/word-media-editor.js"
import { WordDocumentStructure } from "../word/word-document-structure.js"
import { WordFormatting } from "../word/formatting.js"
import { WordMarkdown } from "../word/word-markdown.js"
import type {
  CreateDocumentParams, WriteContentParams, InsertTableParams,
  InsertChartParams, InsertImageParams, InsertListParams, InsertTextboxParams,
  SetHeaderParams, SetFooterParams, AddBookmarkParams, AddCommentParams,
  AddFootnoteParams, AddHyperlinkParams, InsertSectionBreakParams, FormatPageParams,
} from "./types.js"

export class WordDocumentManager {
  constructor(
    private session: IWordSession,
    private appManager: WordApplicationManager,
    private docOps: WordDocument,
    private textEditor: WordTextEditor,
    private tableEditor: WordTableEditor,
    private mediaEditor: WordMediaEditor,
    private documentStructure: WordDocumentStructure,
    private formatting: WordFormatting,
    private markdown: WordMarkdown,
  ) {}

  async createDocument(params: CreateDocumentParams): Promise<string> {
    if (this.session.activeDoc) {
      await this.appManager.closeDocument(false)
    }
    await this.session.start()
    await this.appManager.createDocument({ title: params.title, author: params.author })
    if (params.topMargin != null || params.bottomMargin != null ||
        params.leftMargin != null || params.rightMargin != null ||
        params.orientation != null) {
      await this.formatting.setPageSetup({
        topMargin: params.topMargin, bottomMargin: params.bottomMargin,
        leftMargin: params.leftMargin, rightMargin: params.rightMargin,
        orientation: params.orientation,
      })
    }
    return "Document created and initialized"
  }

  async writeContent(params: WriteContentParams): Promise<string> {
    const app = this.session.application as Record<string, unknown>
    let screenWasOn = false
    try { screenWasOn = (app.ScreenUpdating as boolean) ?? false } catch { /* ignore */ }
    if (screenWasOn) {
      try { app.ScreenUpdating = false } catch { /* ignore */ }
    }
    try {
      await this.textEditor.goTo("end")
      const result = await this.markdown.write(params.text)
      await this.textEditor.insertParagraph(1)
      return `Written ${result.chars} chars in ${result.blocks} blocks`
    } finally {
      if (screenWasOn) {
        try { app.ScreenUpdating = true } catch { /* ignore */ }
        const doc = this.session.activeDoc
        if (doc) {
          try { ;(((doc as Record<string, unknown>).ActiveWindow as Record<string, unknown>).Refresh as () => void)() } catch { /* ignore */ }
        }
      }
    }
  }

  async insertTable(params: InsertTableParams): Promise<string> {
    const result = await this.tableEditor.insertTable({
      rows: params.rows, columns: params.columns,
      data: params.data, autoFitBehavior: params.autoFitBehavior,
    })
    return `Table created: ${result.rows}x${result.columns}`
  }

  async insertChart(params: InsertChartParams): Promise<string> {
    const result = await this.mediaEditor.insertChart(params)
    return `Chart inserted (${result.type}, ${result.series} series)`
  }

  async insertImage(params: InsertImageParams): Promise<string> {
    await this.mediaEditor.insertImage(params)
    return "Image inserted"
  }

  async insertList(params: InsertListParams): Promise<string> {
    await this.textEditor.insertList(params.type, params.items)
    return `List inserted with ${params.items.length} items`
  }

  async insertTextbox(params: InsertTextboxParams): Promise<string> {
    const result = await this.mediaEditor.insertTextbox(params)
    return `Textbox inserted (${result.width}x${result.height})`
  }

  async setHeader(params: SetHeaderParams): Promise<string> {
    await this.documentStructure.setHeader(params.text, params.alignment)
    await this.textEditor.goTo("end")
    return "Header set"
  }

  async setFooter(params: SetFooterParams): Promise<string> {
    await this.documentStructure.setFooter(params.text, params.alignment)
    await this.textEditor.goTo("end")
    return "Footer set"
  }

  async setPageNumbers(target: "header" | "footer"): Promise<string> {
    await this.documentStructure.setPageNumbers(target)
    await this.textEditor.goTo("end")
    return "Page numbers added"
  }

  async setWatermark(text: string): Promise<string> {
    await this.documentStructure.setWatermark({ text })
    await this.textEditor.goTo("end")
    return "Watermark set"
  }

  async addBookmark(params: AddBookmarkParams): Promise<string> {
    await this.documentStructure.addBookmark(params.name)
    await this.textEditor.goTo("end")
    await this.textEditor.insertParagraph(1)
    return "Bookmark added"
  }

  async addComment(params: AddCommentParams): Promise<string> {
    await this.documentStructure.addComment(params.text)
    await this.textEditor.goTo("end")
    await this.textEditor.insertParagraph(1)
    return "Comment added"
  }

  async addFootnote(params: AddFootnoteParams): Promise<string> {
    await this.textEditor.addFootnote(params.text)
    await this.textEditor.goTo("end")
    await this.textEditor.insertParagraph(1)
    return "Footnote added"
  }

  async addHyperlink(params: AddHyperlinkParams): Promise<string> {
    await this.textEditor.addHyperlink(params.text, params.address, params.subAddress, params.screenTip)
    await this.textEditor.goTo("end")
    await this.textEditor.insertParagraph(1)
    return "Hyperlink added"
  }

  async insertSectionBreak(params: InsertSectionBreakParams): Promise<string> {
    await this.textEditor.insertSectionBreak(params.type)
    await this.textEditor.goTo("end")
    return "Section break inserted"
  }

  async formatPage(params: FormatPageParams): Promise<string> {
    await this.formatting.setPageSetup({
      topMargin: params.topMargin, bottomMargin: params.bottomMargin,
      leftMargin: params.leftMargin, rightMargin: params.rightMargin,
      orientation: params.orientation,
      pageWidth: params.pageWidth, pageHeight: params.pageHeight,
    })
    return "Page setup applied"
  }

  async applyHeading(text: string, level: number): Promise<string> {
    level = Math.max(1, Math.min(9, level))
    await this.textEditor.goTo("end")
    await this.textEditor.typeText(text, "instant")
    await this.textEditor.selectCurrentParagraph()
    await this.formatting.applyStyle(`Heading ${level}`)
    await this.textEditor.goTo("end")
    await this.textEditor.insertParagraph(1)
    return `Heading ${level} applied: "${text}"`
  }

  async saveAndExport(exportPath?: string): Promise<string> {
    await this.appManager.saveDocument()
    let result = "Document saved"
    if (exportPath) {
      await this.docOps.exportToPdf(exportPath)
      result += ` and exported to ${exportPath}`
    }
    return result
  }

  async closeDocument(save?: boolean): Promise<string> {
    await this.appManager.closeDocument(save ?? false)
    return "Document closed"
  }
}
