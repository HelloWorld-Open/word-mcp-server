import { vi } from "vitest"
import type { IDocumentProxy, ISelectionProxy, IRangeProxy, IHeaderFooterProxy, ISectionProxy } from "./types.js"

function mockRangeProxy(raw?: Record<string, unknown>) {
  return {
    raw: (raw ?? {}) as Record<string, unknown>,
    getText: vi.fn(() => (raw?.Text as string) ?? ""),
    setText: vi.fn(),
    getStart: vi.fn(() => (raw?.Start as number) ?? 0),
    getEnd: vi.fn(() => (raw?.End as number) ?? 0),
    setStart: vi.fn(),
    setEnd: vi.fn(),
    setRange: vi.fn(),
    getBold: vi.fn(() => undefined),
    setBold: vi.fn(),
    getItalic: vi.fn(() => undefined),
    setItalic: vi.fn(),
    select: vi.fn(),
    duplicate: vi.fn(),
    getFind: vi.fn(() => ({})),
    getFont: vi.fn(() => ({})),
    getParagraphFormat: vi.fn(() => ({})),
    getShading: vi.fn(() => ({})),
    getListFormat: vi.fn(() => ({})),
    getHyperlinks: vi.fn(() => ({})),
    insertFile: vi.fn(),
    convertToTable: vi.fn(() => ({})),
    addField: vi.fn(),
  }
}

function mockHeaderFooterProxy(): IHeaderFooterProxy {
  return {
    select: vi.fn(),
    clearContent: vi.fn(),
    typeText: vi.fn(),
    setAlignment: vi.fn(),
    setContent: vi.fn(),
    getEnd: vi.fn(() => 0),
    getText: vi.fn(() => ""),
    getRange: vi.fn(() => mockRangeProxy()),
    getFields: vi.fn(() => ({ count: 0, item: vi.fn() })),
    getPageNumbersCount: vi.fn(() => 0),
    setPageNumbersAlignment: vi.fn(),
    raw: {},
  }
}

function mockSectionProxy(raw: Record<string, unknown>): ISectionProxy {
  return {
    getHeader: vi.fn(() => mockHeaderFooterProxy()),
    getFooter: vi.fn(() => mockHeaderFooterProxy()),
    getPageSetup: vi.fn(() => (raw.PageSetup as Record<string, unknown>) ?? {}),
    raw,
  }
}

function mockCollection(raw: Record<string, unknown> | undefined, key: string) {
  const coll = raw?.[key] as Record<string, unknown> | undefined
  return {
    count: (coll?.Count as number) ?? 0,
    item: vi.fn((i: number) => {
      const rawColl = raw?.[key] as Record<string, unknown> | undefined
      return (rawColl?.Item as ((i: number) => Record<string, unknown>) | undefined)?.(i) ?? {}
    }),
  }
}

export class MockDocumentProxy implements IDocumentProxy {
  private _raw: Record<string, unknown>

  constructor(raw?: Record<string, unknown>) {
    this._raw = raw ?? {}
  }

  get raw(): Record<string, unknown> { return this._raw }

  getName = vi.fn(() => (this._raw.Name as string) ?? "mock.docx")
  getFullName = vi.fn(() => (this._raw.FullName as string) ?? "C:\\mock.docx")
  getPath = vi.fn(() => (this._raw.Path as string) ?? "C:\\")
  getSaved = vi.fn(() => (this._raw.Saved as boolean) ?? true)
  getContent = vi.fn(() => new MockRangeProxy(this._raw.Content as Record<string, unknown>))
  getParagraphs = vi.fn(() => mockCollection(this._raw, "Paragraphs"))
  getSections = vi.fn(() => {
    const base = mockCollection(this._raw, "Sections")
    return {
      ...base,
      item: vi.fn((i: number) => {
        const rawColl = this._raw.Sections as Record<string, unknown> | undefined
        const raw = (rawColl?.Item as ((i: number) => Record<string, unknown>) | undefined)?.(i) ?? {}
        return mockSectionProxy(raw)
      }),
    }
  })
  getTables = vi.fn(() => ({ ...mockCollection(this._raw, "Tables"), add: vi.fn() }))
  getBookmarks = vi.fn(() => ({ ...mockCollection(this._raw, "Bookmarks"), add: vi.fn() }))
  getStyles = vi.fn(function (this: MockDocumentProxy) {
    const raw = this._raw
    return {
      ...mockCollection(raw, "Styles"),
      itemByName: vi.fn((name: string) => {
        const coll = raw.Styles as Record<string, unknown> | undefined
        return (coll?.Item as ((n: string) => Record<string, unknown>) | undefined)?.(name) ?? {}
      }),
    }
  })
  getInlineShapes = vi.fn(() => ({ ...mockCollection(this._raw, "InlineShapes"), addPicture: vi.fn(), addChart2: vi.fn(() => ({ Width: 0, Height: 0, Chart: { HasTitle: false, ChartTitle: { Text: "" }, ChartData: { Workbook: { Application: { Quit: vi.fn() } } } } } as Record<string, unknown>)), addHorizontalLineStandard: vi.fn(), addTextbox: vi.fn(() => ({ Width: 0, Height: 0 } as Record<string, unknown>)) }))
  getComments = vi.fn(() => ({ ...mockCollection(this._raw, "Comments"), add: vi.fn() }))
  getFootnotes = vi.fn(() => ({ ...mockCollection(this._raw, "Footnotes"), add: vi.fn() }))
  getRevisions = vi.fn(() => ({ ...mockCollection(this._raw, "Revisions"), acceptAll: vi.fn(), rejectAll: vi.fn() }))
  getHyperlinks = vi.fn(() => ({ ...mockCollection(this._raw, "Hyperlinks"), add: vi.fn() }))
  getBuiltInDocumentProperties = vi.fn(() => {
    const base = mockCollection(this._raw, "BuiltInDocumentProperties")
    return { ...base, itemByName: vi.fn(() => ({ Value: undefined })) }
  })
  getPageSetup = vi.fn(() => (this._raw.PageSetup as Record<string, unknown>) ?? {})
  getTrackRevisions = vi.fn(() => (this._raw.TrackRevisions as boolean) ?? false)
  setTrackRevisions = vi.fn()
  getRange = vi.fn((start?: number, end?: number) => {
    const raw = (this._raw.Range as ((s?: number, e?: number) => Record<string, unknown>) | undefined)?.(start, end)
    return new MockRangeProxy(raw ?? {})
  })
  select = vi.fn()
  save = vi.fn()
  saveAs = vi.fn()
  close = vi.fn()
  exportAsFixedFormat = vi.fn()
  computeStatistics = vi.fn(() => 0)
  getTablesOfContents = vi.fn(() => ({}))
  getShapes = vi.fn(() => ({ AddTextbox: vi.fn(() => ({ Width: 0, Height: 0, TextFrame: { TextRange: { Text: "" } }, Chart: { HasTitle: false, ChartData: { Workbook: { Application: { Quit: vi.fn() } } } } } as Record<string, unknown>)) } as Record<string, unknown>))
  getLists = vi.fn(() => ({}))
  undo = vi.fn()
  redo = vi.fn()
}

export class MockSelectionProxy implements ISelectionProxy {
  private _raw: Record<string, unknown>

  constructor(raw?: Record<string, unknown>) {
    this._raw = raw ?? {}
  }

  get raw(): Record<string, unknown> { return this._raw }

  getStart = vi.fn(() => (this._raw.Start as number) ?? 0)
  getEnd = vi.fn(() => (this._raw.End as number) ?? 0)
  getStoryType = vi.fn(() => (this._raw.StoryType as number) ?? 1)
  getStyle = vi.fn(() => (this._raw.Style as string | number) ?? "")
  setStyle = vi.fn()
  getType = vi.fn(() => (this._raw.Type as number) ?? 1)
  typeText = vi.fn()
  typeParagraph = vi.fn()
  typeBackspace = vi.fn()
  collapse = vi.fn()
  endKey = vi.fn()
  moveStart = vi.fn()
  homeKey = vi.fn()
  wholeStory = vi.fn()
  delete = vi.fn()
  copy = vi.fn()
  cut = vi.fn()
  paste = vi.fn()
  select = vi.fn()
  expand = vi.fn()
  goTo = vi.fn(() => ({}))
  insertBreak = vi.fn()
  getRange = vi.fn(() => new MockRangeProxy(this._raw.Range as Record<string, unknown>))
  getInformation = vi.fn(() => false)
  addHorizontalLine = vi.fn()
  getFind = vi.fn(() => {
    const raw = this._raw.Find as Record<string, unknown> | undefined
    return raw ?? { ClearFormatting: vi.fn(), Replacement: { ClearFormatting: vi.fn(), Text: "" }, Execute: vi.fn() }
  })
  getFont = vi.fn(() => ({}))
  getParagraphFormat = vi.fn(() => ({}))
  getInlineShapes = vi.fn(() => ({}))
  getTables = vi.fn(() => ({}))
  getShapeRange = vi.fn(() => ({}))
}

export class MockRangeProxy implements IRangeProxy {
  private _raw: Record<string, unknown>

  constructor(raw?: Record<string, unknown>) {
    this._raw = raw ?? {}
  }

  get raw(): Record<string, unknown> { return this._raw }

  getText = vi.fn(() => (this._raw.Text as string) ?? "")
  setText = vi.fn()
  getStart = vi.fn(() => (this._raw.Start as number) ?? 0)
  getEnd = vi.fn(() => (this._raw.End as number) ?? 0)
  setStart = vi.fn()
  setEnd = vi.fn()
  setRange = vi.fn()
  getBold = vi.fn(() => this._raw.Bold as boolean | undefined)
  setBold = vi.fn()
  getItalic = vi.fn(() => this._raw.Italic as boolean | undefined)
  setItalic = vi.fn()
  select = vi.fn()
  duplicate = vi.fn(() => {
    const rawDup = this._raw.Duplicate as Record<string, unknown> | undefined
    return new MockRangeProxy(rawDup ?? { ...this._raw })
  })
  getFind = vi.fn(() => (this._raw.Find as Record<string, unknown>) ?? {})
  getFont = vi.fn(() => (this._raw.Font as Record<string, unknown>) ?? {})
  getParagraphFormat = vi.fn(() => (this._raw.ParagraphFormat as Record<string, unknown>) ?? {})
  getShading = vi.fn(() => (this._raw.Shading as Record<string, unknown>) ?? {})
  getListFormat = vi.fn(() => (this._raw.ListFormat as Record<string, unknown>) ?? {})
  getHyperlinks = vi.fn(() => (this._raw.Hyperlinks as Record<string, unknown>) ?? {})
  insertFile = vi.fn()
  convertToTable = vi.fn(() => ({}))
  addField = vi.fn()
}

export function createMockDocProxy(raw?: Record<string, unknown>): MockDocumentProxy {
  return new MockDocumentProxy(raw)
}

export function createMockSelProxy(raw?: Record<string, unknown>): MockSelectionProxy {
  return new MockSelectionProxy(raw)
}

export function createMockRangeProxy(raw?: Record<string, unknown>): MockRangeProxy {
  return new MockRangeProxy(raw)
}
