import type { IDocumentProxy, ISelectionProxy, IRangeProxy, IParagraphsProxy, ISectionsProxy, ITablesProxy, IBookmarksProxy, IStylesProxy, IInlineShapesProxy, ICommentsProxy, IFootnotesProxy, IRevisionsProxy, IHyperlinksProxy, IPropertiesProxy, IPageSetupProxy } from "./types.js"
import { CollectionProxy } from "./collection-proxy.js"
import { SectionProxy } from "./section-proxy.js"

export class DocumentProxy implements IDocumentProxy {
  constructor(
    private rawDoc: Record<string, unknown>,
    private session: {
      wrapRange: (raw: Record<string, unknown>) => IRangeProxy
      getSelectionProxy: () => ISelectionProxy
      log?: (level: string, msg: string) => void
    },
  ) {}

  get raw(): Record<string, unknown> {
    return this.rawDoc
  }

  getName(): string {
    try { return this.rawDoc.Name as string } catch { this.session.log?.("debug", "DocumentProxy.getName failed"); return "未命名文档" }
  }

  getFullName(): string | undefined {
    try { return this.rawDoc.FullName as string } catch { this.session.log?.("debug", "DocumentProxy.getFullName failed"); return undefined }
  }

  getPath(): string | undefined {
    try { return this.rawDoc.Path as string } catch { this.session.log?.("debug", "DocumentProxy.getPath failed"); return undefined }
  }

  getSaved(): boolean {
    try { return this.rawDoc.Saved as boolean } catch { this.session.log?.("debug", "DocumentProxy.getSaved failed"); return true }
  }

  getContent(): IRangeProxy {
    try {
      const content = this.rawDoc.Content as Record<string, unknown>
      return this.session.wrapRange(content)
    } catch { this.session.log?.("debug", "DocumentProxy.getContent failed")
      return this.session.wrapRange({} as Record<string, unknown>)
    }
  }

  getParagraphs(): IParagraphsProxy {
    try {
      return new CollectionProxy(this.rawDoc.Paragraphs as Record<string, unknown>, (raw) => raw)
    } catch { this.session.log?.("debug", "DocumentProxy.getParagraphs failed")
      return new CollectionProxy({} as Record<string, unknown>, (raw) => raw)
    }
  }

  getSections(): ISectionsProxy {
    const sd = {
      getSelectionProxy: () => this.session.getSelectionProxy(),
      wrapRange: (r: Record<string, unknown>) => this.session.wrapRange(r),
      log: (level: string, msg: string) => this.session.log?.(level, msg),
    }
    try {
      return new CollectionProxy(
        this.rawDoc.Sections as Record<string, unknown>,
        (raw) => new SectionProxy(raw, sd),
      )
    } catch {
      this.session.log?.("debug", "DocumentProxy.getSections failed")
      return new CollectionProxy(
        {} as Record<string, unknown>,
        (raw) => new SectionProxy(raw, sd),
      )
    }
  }

  getTables(): ITablesProxy {
    try {
      const raw = this.rawDoc.Tables as Record<string, unknown>
      const base = new CollectionProxy(raw, (r) => r)
      return Object.assign(base, {
        add: (range: unknown, rows: number, cols: number) => {
          try {
            return (raw.Add as (r: unknown, ro: number, co: number) => Record<string, unknown>)(range, rows, cols)
          } catch { this.session.log?.("debug", "DocumentProxy.Tables.add failed")
            return {}
          }
        },
      })
    } catch { this.session.log?.("debug", "DocumentProxy.getTables failed")
      const base = new CollectionProxy({} as Record<string, unknown>, (r) => r)
      return Object.assign(base, { add: () => ({}) })
    }
  }

  getBookmarks(): IBookmarksProxy {
    try {
      const raw = this.rawDoc.Bookmarks as Record<string, unknown>
      const base = new CollectionProxy(raw, (r) => r)
      return Object.assign(base, {
        add: (name: string, range?: Record<string, unknown>) => {
          try {
            return (raw.Add as (n: string, r?: Record<string, unknown>) => Record<string, unknown>)(name, range)
          } catch { this.session.log?.("debug", "DocumentProxy.Bookmarks.add failed")
            return {}
          }
        },
      })
    } catch { this.session.log?.("debug", "DocumentProxy.getBookmarks failed")
      const base = new CollectionProxy({} as Record<string, unknown>, (r) => r)
      return Object.assign(base, { add: () => ({}) })
    }
  }

  getStyles(): IStylesProxy {
    try {
      const raw = this.rawDoc.Styles as Record<string, unknown>
      const base = new CollectionProxy(raw, (r) => r)
      return Object.assign(base, {
        itemByName: (name: string) => {
          try { return (raw.Item as (n: string) => Record<string, unknown>)(name) } catch { this.session.log?.("debug", "DocumentProxy.Styles.itemByName failed"); return {} }
        },
      })
    } catch { this.session.log?.("debug", "DocumentProxy.getStyles failed")
      const base = new CollectionProxy({} as Record<string, unknown>, (r) => r)
      return Object.assign(base, { itemByName: () => ({}) })
    }
  }

  getInlineShapes(): IInlineShapesProxy {
    try {
      const raw = this.rawDoc.InlineShapes as Record<string, unknown>
      const base = new CollectionProxy(raw, (r) => r)
      return Object.assign(base, {
        addPicture: (path: string) => {
          try {
            return (raw.AddPicture as (p: string) => Record<string, unknown>)(path)
          } catch { this.session.log?.("debug", "DocumentProxy.InlineShapes.addPicture failed")
            return {}
          }
        },
        addChart2: (style: number, type: number, range: unknown) => {
          try {
            return (raw.AddChart2 as (s: number, t: number, r: unknown) => Record<string, unknown>)(style, type, range)
          } catch { this.session.log?.("debug", "DocumentProxy.InlineShapes.addChart2 failed")
            return {}
          }
        },
        addHorizontalLineStandard: () => {
          try { ;(raw.AddHorizontalLineStandard as () => void)() } catch { this.session.log?.("warn", "DocumentProxy op failed") }
        },
        addTextbox: (orientation: number, left: number, top: number, width: number, height: number) => {
          try {
            return (raw.AddTextbox as (o: number, l: number, t: number, w: number, h: number) => Record<string, unknown>)(orientation, left, top, width, height)
          } catch { this.session.log?.("debug", "DocumentProxy.InlineShapes.addTextbox failed")
            return {}
          }
        },
      })
    } catch { this.session.log?.("debug", "DocumentProxy.getInlineShapes failed")
      const base = new CollectionProxy({} as Record<string, unknown>, (r) => r)
      return Object.assign(base, { addPicture: () => ({}), addChart2: () => ({}), addHorizontalLineStandard: () => {}, addTextbox: () => ({}) })
    }
  }

  getComments(): ICommentsProxy {
    try {
      const raw = this.rawDoc.Comments as Record<string, unknown>
      const base = new CollectionProxy(raw, (r) => r)
      return Object.assign(base, {
        add: (range: unknown, text: string) => {
          try {
            return (raw.Add as (r: unknown, t: string) => unknown)(range, text)
          } catch { this.session.log?.("debug", "DocumentProxy.Comments.add failed"); return undefined }
        },
      })
    } catch { this.session.log?.("debug", "DocumentProxy.getComments failed")
      const base = new CollectionProxy({} as Record<string, unknown>, (r) => r)
      return Object.assign(base, { add: () => {} })
    }
  }

  getFootnotes(): IFootnotesProxy {
    try {
      const raw = this.rawDoc.Footnotes as Record<string, unknown>
      const base = new CollectionProxy(raw, (r) => r)
      return Object.assign(base, {
        add: (range: unknown, text: string) => {
          try {
            return (raw.Add as (r: unknown, t: string) => unknown)(range, text)
          } catch { this.session.log?.("debug", "DocumentProxy.Footnotes.add failed"); return undefined }
        },
      })
    } catch { this.session.log?.("debug", "DocumentProxy.getFootnotes failed")
      const base = new CollectionProxy({} as Record<string, unknown>, (r) => r)
      return Object.assign(base, { add: () => {} })
    }
  }

  getRevisions(): IRevisionsProxy {
    try {
      const raw = this.rawDoc.Revisions as Record<string, unknown>
      const base = new CollectionProxy(raw, (r) => r)
      return Object.assign(base, {
        acceptAll: () => {
          try { ;(raw.AcceptAll as () => void)() } catch { this.session.log?.("warn", "DocumentProxy op failed") }
        },
        rejectAll: () => {
          try { ;(raw.RejectAll as () => void)() } catch { this.session.log?.("warn", "DocumentProxy op failed") }
        },
      })
    } catch { this.session.log?.("debug", "DocumentProxy.getRevisions failed")
      const base = new CollectionProxy({} as Record<string, unknown>, (r) => r)
      return Object.assign(base, { acceptAll: () => {}, rejectAll: () => {} })
    }
  }

  getHyperlinks(): IHyperlinksProxy {
    try {
      const raw = this.rawDoc.Hyperlinks as Record<string, unknown>
      const base = new CollectionProxy(raw, (r) => r)
      return Object.assign(base, {
        add: (anchor: Record<string, unknown>, address: string, subAddress?: string, screenTip?: string, textToDisplay?: string) => {
          try {
            return (raw.Add as (a: Record<string, unknown>, addr: string, sub?: string, tip?: string, display?: string) => Record<string, unknown>)(anchor, address, subAddress, screenTip, textToDisplay)
          } catch { this.session.log?.("debug", "DocumentProxy.Hyperlinks.add failed")
            return {}
          }
        },
      })
    } catch { this.session.log?.("debug", "DocumentProxy.getHyperlinks failed")
      const base = new CollectionProxy({} as Record<string, unknown>, (r) => r)
      return Object.assign(base, { add: () => ({}) })
    }
  }

  getBuiltInDocumentProperties(): IPropertiesProxy {
    try {
      const raw = this.rawDoc.BuiltInDocumentProperties as Record<string, unknown>
      const base = new CollectionProxy(raw, (r) => r)
      return Object.assign(base, {
        itemByName: (name: string) => {
          try { return (raw.Item as (n: string) => Record<string, unknown>)(name) } catch { this.session.log?.("debug", "DocumentProxy.BuiltInDocumentProperties.itemByName failed"); return {} }
        },
      })
    } catch { this.session.log?.("debug", "DocumentProxy.getBuiltInDocumentProperties failed")
      const base = new CollectionProxy({} as Record<string, unknown>, (r) => r)
      return Object.assign(base, { itemByName: () => ({}) })
    }
  }

  getPageSetup(): IPageSetupProxy {
    try {
      return (this.rawDoc.PageSetup as Record<string, unknown>) ?? {}
    } catch { this.session.log?.("debug", "DocumentProxy.getPageSetup failed")
      return {}
    }
  }

  getTrackRevisions(): boolean {
    try { return this.rawDoc.TrackRevisions as boolean } catch { this.session.log?.("debug", "DocumentProxy.getTrackRevisions failed"); return false }
  }

  setTrackRevisions(v: boolean): void {
    try { this.rawDoc.TrackRevisions = v } catch { this.session.log?.("warn", "DocumentProxy op failed") }
  }

  getRange(start?: number, end?: number): IRangeProxy {
    try {
      const raw = (this.rawDoc.Range as (s?: number, e?: number) => Record<string, unknown>)(start, end)
      return this.session.wrapRange(raw)
    } catch { this.session.log?.("debug", "DocumentProxy.getRange failed")
      return this.session.wrapRange({} as Record<string, unknown>)
    }
  }

  select(): void {
    try { ;(this.rawDoc.Select as () => void)() } catch { this.session.log?.("warn", "DocumentProxy op failed") }
  }

  save(): void {
    try { ;(this.rawDoc.Save as () => void)() } catch { this.session.log?.("warn", "DocumentProxy op failed") }
  }

  saveAs(path: string, format?: number): void {
    try {
      if (format != null) {
        ;(this.rawDoc.SaveAs as (p: string, f: number) => void)(path, format)
      } else {
        ;(this.rawDoc.SaveAs as (p: string) => void)(path)
      }
    } catch { this.session.log?.("warn", "DocumentProxy op failed") }
  }

  close(): void {
    try { ;(this.rawDoc.Close as () => void)() } catch { this.session.log?.("warn", "DocumentProxy op failed") }
  }

  exportAsFixedFormat(path: string, format: number): void {
    try {
      ;(this.rawDoc.ExportAsFixedFormat as (p: string, f: number) => void)(path, format)
    } catch { this.session.log?.("warn", "DocumentProxy op failed") }
  }

  computeStatistics(type: number): number {
    try { return (this.rawDoc.ComputeStatistics as (t: number) => number)(type) } catch { this.session.log?.("debug", "DocumentProxy.computeStatistics failed"); return 0 }
  }

  getTablesOfContents(): Record<string, unknown> {
    try { return this.rawDoc.TablesOfContents as Record<string, unknown> } catch { this.session.log?.("debug", "DocumentProxy.getTablesOfContents failed"); return {} }
  }

  getShapes(): Record<string, unknown> {
    try { return this.rawDoc.Shapes as Record<string, unknown> } catch { this.session.log?.("debug", "DocumentProxy.getShapes failed"); return {} }
  }

  getLists(): Record<string, unknown> {
    try { return this.rawDoc.Lists as Record<string, unknown> } catch { this.session.log?.("debug", "DocumentProxy.getLists failed"); return {} }
  }

  undo(): void {
    try { ;(this.rawDoc.Undo as () => void)() } catch { this.session.log?.("warn", "DocumentProxy op failed") }
  }

  redo(): void {
    try { ;(this.rawDoc.Redo as () => void)() } catch { this.session.log?.("warn", "DocumentProxy op failed") }
  }
}
