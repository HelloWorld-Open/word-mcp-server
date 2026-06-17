import { describe, it, expect, vi } from "vitest"
import { DocumentProxy } from "../../../src/word/com-proxy/document-proxy.js"
import { SelectionProxy } from "../../../src/word/com-proxy/selection-proxy.js"
import { RangeProxy } from "../../../src/word/com-proxy/range-proxy.js"
import { CollectionProxy } from "../../../src/word/com-proxy/collection-proxy.js"
import type { IDocumentProxy, ISelectionProxy, IRangeProxy } from "../../../src/word/com-proxy/types.js"

describe("DocumentProxy", () => {
  function makeDoc(overrides?: Record<string, unknown>): Record<string, unknown> {
    return {
      Name: "test.docx",
      FullName: "C:\\docs\\test.docx",
      Path: "C:\\docs",
      Saved: true,
      Content: { Text: "hello", Start: 0, End: 5 },
      Paragraphs: { Count: 2, Item: vi.fn() },
      Sections: { Count: 1, Item: vi.fn() },
      Tables: { Count: 0, Item: vi.fn() },
      Bookmarks: { Count: 2, Item: vi.fn(), Add: vi.fn() },
      Styles: { Count: 5, Item: vi.fn() },
      InlineShapes: { Count: 0, Item: vi.fn(), AddPicture: vi.fn() },
      Comments: { Count: 0, Item: vi.fn() },
      Footnotes: { Count: 0, Item: vi.fn() },
      Revisions: { Count: 0, AcceptAll: vi.fn(), RejectAll: vi.fn() },
      Hyperlinks: { Count: 0, Add: vi.fn() },
      BuiltInDocumentProperties: { Count: 10, Item: vi.fn() },
      PageSetup: { TopMargin: 50 },
      TrackRevisions: false,
      Range: vi.fn(() => ({ Text: "range", Start: 0, End: 5 })),
      Select: vi.fn(),
      Save: vi.fn(),
      SaveAs: vi.fn(),
      Close: vi.fn(),
      ExportAsFixedFormat: vi.fn(),
      ...overrides,
    } as unknown as Record<string, unknown>
  }

  function createSession() {
    const rangeCache = new Map<Record<string, unknown>, IRangeProxy>()
    return {
      wrapRange: (raw: Record<string, unknown>): IRangeProxy => {
        let p = rangeCache.get(raw)
        if (!p) {
          p = new RangeProxy(raw, createSession())
          rangeCache.set(raw, p)
        }
        return p
      },
      getSelectionProxy: () => ({ raw: {} as Record<string, unknown> } as unknown as ISelectionProxy),
    }
  }

  it("getName returns the document name", () => {
    const doc = new DocumentProxy(makeDoc(), createSession())
    expect(doc.getName()).toBe("test.docx")
  })

  it("getName falls back when COM throws", () => {
    const raw = makeDoc()
    Object.defineProperty(raw, "Name", { get: () => { throw new Error("COM error") } })
    const doc = new DocumentProxy(raw, createSession())
    expect(doc.getName()).toBe("未命名文档")
  })

  it("getFullName returns full path", () => {
    const doc = new DocumentProxy(makeDoc(), createSession())
    expect(doc.getFullName()).toBe("C:\\docs\\test.docx")
  })

  it("getFullName returns undefined when COM throws", () => {
    const raw = makeDoc()
    Object.defineProperty(raw, "FullName", { get: () => { throw new Error("COM error") } })
    const doc = new DocumentProxy(raw, createSession())
    expect(doc.getFullName()).toBeUndefined()
  })

  it("getPath returns the path", () => {
    const doc = new DocumentProxy(makeDoc(), createSession())
    expect(doc.getPath()).toBe("C:\\docs")
  })

  it("getSaved returns saved state", () => {
    const doc = new DocumentProxy(makeDoc({ Saved: false }), createSession())
    expect(doc.getSaved()).toBe(false)
  })

  it("getSaved defaults true on error", () => {
    const raw = makeDoc()
    Object.defineProperty(raw, "Saved", { get: () => { throw new Error("COM error") } })
    const doc = new DocumentProxy(raw, createSession())
    expect(doc.getSaved()).toBe(true)
  })

  it("getContent returns a RangeProxy", () => {
    const doc = new DocumentProxy(makeDoc(), createSession())
    expect(doc.getContent().getText()).toBe("hello")
  })

  it("getParagraphs returns collection with count", () => {
    const doc = new DocumentProxy(makeDoc(), createSession())
    expect(doc.getParagraphs().count).toBe(2)
  })

  it("getSections returns collection", () => {
    const doc = new DocumentProxy(makeDoc(), createSession())
    expect(doc.getSections().count).toBe(1)
  })

  it("getTables returns collection", () => {
    const doc = new DocumentProxy(makeDoc(), createSession())
    expect(doc.getTables().count).toBe(0)
  })

  it("getBookmarks allows adding a bookmark", () => {
    const raw = makeDoc()
    const doc = new DocumentProxy(raw, createSession())
    doc.getBookmarks().add("myBookmark")
    expect(raw.Bookmarks.Add).toHaveBeenCalledWith("myBookmark", undefined)
  })

  it("getRange delegates to doc.Range", () => {
    const raw = makeDoc()
    const doc = new DocumentProxy(raw, createSession())
    const rng = doc.getRange(1, 10)
    expect(raw.Range).toHaveBeenCalledWith(1, 10)
    expect(rng.getText()).toBe("range")
  })

  it("save calls doc.Save", () => {
    const raw = makeDoc()
    const doc = new DocumentProxy(raw, createSession())
    doc.save()
    expect(raw.Save).toHaveBeenCalled()
  })

  it("exportAsFixedFormat calls the COM method", () => {
    const raw = makeDoc()
    const doc = new DocumentProxy(raw, createSession())
    doc.exportAsFixedFormat("out.pdf", 17)
    expect(raw.ExportAsFixedFormat).toHaveBeenCalledWith("out.pdf", 17)
  })

  it("select calls doc.Select", () => {
    const raw = makeDoc()
    const doc = new DocumentProxy(raw, createSession())
    doc.select()
    expect(raw.Select).toHaveBeenCalled()
  })

  it("getTrackRevisions reads the property", () => {
    const doc = new DocumentProxy(makeDoc({ TrackRevisions: true }), createSession())
    expect(doc.getTrackRevisions()).toBe(true)
  })

  it("setTrackRevisions writes the property", () => {
    const raw = makeDoc()
    const doc = new DocumentProxy(raw, createSession())
    doc.setTrackRevisions(true)
    expect(raw.TrackRevisions).toBe(true)
  })

  it("getPageSetup returns page setup object", () => {
    const doc = new DocumentProxy(makeDoc(), createSession())
    expect(doc.getPageSetup()).toEqual({ TopMargin: 50 })
  })
})

describe("SelectionProxy", () => {
  function createSession() {
    const rangeCache = new Map<Record<string, unknown>, IRangeProxy>()
    return {
      wrapRange: (raw: Record<string, unknown>): IRangeProxy => {
        let p = rangeCache.get(raw)
        if (!p) {
          p = new RangeProxy(raw, createSession())
          rangeCache.set(raw, p)
        }
        return p
      },
    }
  }

  function makeSel(overrides?: Record<string, unknown>): Record<string, unknown> {
    return {
      Start: 5, End: 10, StoryType: 1, Type: 1,
      Style: "Normal",
      TypeText: vi.fn(),
      TypeParagraph: vi.fn(),
      TypeBackspace: vi.fn(),
      Collapse: vi.fn(),
      EndKey: vi.fn(),
      MoveStart: vi.fn(),
      HomeKey: vi.fn(),
      WholeStory: vi.fn(),
      Delete: vi.fn(),
      Copy: vi.fn(),
      Cut: vi.fn(),
      Paste: vi.fn(),
      Select: vi.fn(),
      Expand: vi.fn(),
      GoTo: vi.fn(() => ({})),
      InsertBreak: vi.fn(),
      Range: { Text: "selected", Start: 5, End: 10, Font: {} },
      Information: vi.fn(() => 0),
      ...overrides,
    } as unknown as Record<string, unknown>
  }

  it("getStart returns selection start", () => {
    const sel = new SelectionProxy(makeSel(), createSession())
    expect(sel.getStart()).toBe(5)
  })

  it("getEnd returns selection end", () => {
    const sel = new SelectionProxy(makeSel(), createSession())
    expect(sel.getEnd()).toBe(10)
  })

  it("getStoryType returns story type", () => {
    const sel = new SelectionProxy(makeSel(), createSession())
    expect(sel.getStoryType()).toBe(1)
  })

  it("getStyle returns style name", () => {
    const sel = new SelectionProxy(makeSel(), createSession())
    expect(sel.getStyle()).toBe("Normal")
  })

  it("typeText calls the COM method", () => {
    const raw = makeSel()
    const sel = new SelectionProxy(raw, createSession())
    sel.typeText("hello")
    expect(raw.TypeText).toHaveBeenCalledWith("hello")
  })

  it("typeParagraph calls the COM method", () => {
    const raw = makeSel()
    const sel = new SelectionProxy(raw, createSession())
    sel.typeParagraph()
    expect(raw.TypeParagraph).toHaveBeenCalled()
  })

  it("collapse calls Collapse with default direction 0", () => {
    const raw = makeSel()
    const sel = new SelectionProxy(raw, createSession())
    sel.collapse()
    expect(raw.Collapse).toHaveBeenCalledWith(0)
  })

  it("endKey calls EndKey with unit", () => {
    const raw = makeSel()
    const sel = new SelectionProxy(raw, createSession())
    sel.endKey(6)
    expect(raw.EndKey).toHaveBeenCalledWith(6)
  })

  it("getRange returns a RangeProxy", () => {
    const sel = new SelectionProxy(makeSel(), createSession())
    expect(sel.getRange().getText()).toBe("selected")
  })

  it("getInformation delegates to COM", () => {
    const raw = makeSel()
    const sel = new SelectionProxy(raw, createSession())
    sel.getInformation(12)
    expect(raw.Information).toHaveBeenCalledWith(12)
  })

  it("delete calls Delete", () => {
    const raw = makeSel()
    const sel = new SelectionProxy(raw, createSession())
    sel.delete()
    expect(raw.Delete).toHaveBeenCalled()
  })

  it("copy/cut/paste delegate to COM", () => {
    const raw = makeSel()
    const sel = new SelectionProxy(raw, createSession())
    sel.copy()
    sel.cut()
    sel.paste()
    expect(raw.Copy).toHaveBeenCalled()
    expect(raw.Cut).toHaveBeenCalled()
    expect(raw.Paste).toHaveBeenCalled()
  })
})

describe("RangeProxy", () => {
  function createSession() {
    const rangeCache = new Map<Record<string, unknown>, IRangeProxy>()
    return {
      wrapRange: (raw: Record<string, unknown>): IRangeProxy => {
        let p = rangeCache.get(raw)
        if (!p) {
          p = new RangeProxy(raw, createSession())
          rangeCache.set(raw, p)
        }
        return p
      },
    }
  }

  function makeRange(overrides?: Record<string, unknown>): Record<string, unknown> {
    return {
      Text: "hello world",
      Start: 0, End: 11,
      Bold: true, Italic: false,
      Select: vi.fn(),
      Duplicate: { Text: "dupe", Start: 0, End: 3 },
      Find: {},
      Font: { Name: "Arial", Size: 12 },
      ParagraphFormat: {},
      Shading: { BackgroundPatternColor: 0xFFFFFF },
      ListFormat: {},
      Hyperlinks: {},
      InsertFile: vi.fn(),
      ConvertToTable: vi.fn(() => ({})),
      ...overrides,
    } as unknown as Record<string, unknown>
  }

  it("getText returns range text", () => {
    const rng = new RangeProxy(makeRange(), createSession())
    expect(rng.getText()).toBe("hello world")
  })

  it("getText returns empty on COM error", () => {
    const raw = makeRange()
    Object.defineProperty(raw, "Text", { get: () => { throw new Error("COM error") } })
    const rng = new RangeProxy(raw, createSession())
    expect(rng.getText()).toBe("")
  })

  it("setText writes the property", () => {
    const raw = makeRange()
    const rng = new RangeProxy(raw, createSession())
    rng.setText("new text")
    expect(raw.Text).toBe("new text")
  })

  it("getStart/End return positions", () => {
    const rng = new RangeProxy(makeRange(), createSession())
    expect(rng.getStart()).toBe(0)
    expect(rng.getEnd()).toBe(11)
  })

  it("getBold returns the bold state", () => {
    const rng = new RangeProxy(makeRange(), createSession())
    expect(rng.getBold()).toBe(true)
  })

  it("setBold writes the property", () => {
    const raw = makeRange()
    const rng = new RangeProxy(raw, createSession())
    rng.setBold(false)
    expect(raw.Bold).toBe(false)
  })

  it("select calls COM Select", () => {
    const raw = makeRange()
    const rng = new RangeProxy(raw, createSession())
    rng.select()
    expect(raw.Select).toHaveBeenCalled()
  })

  it("duplicate returns a new RangeProxy", () => {
    const rng = new RangeProxy(makeRange(), createSession())
    const dup = rng.duplicate()
    expect(dup.getText()).toBe("dupe")
  })

  it("getFont returns raw font object", () => {
    const rng = new RangeProxy(makeRange(), createSession())
    const font = rng.getFont() as Record<string, unknown>
    expect(font.Name).toBe("Arial")
  })
})

describe("CollectionProxy", () => {
  it("count reads from COM Count", () => {
    const raw = { Count: 5, Item: vi.fn(() => ({})) } as unknown as Record<string, unknown>
    const col = new CollectionProxy(raw, (r) => r)
    expect(col.count).toBe(5)
  })

  it("item calls COM Item and transforms", () => {
    const raw = { Count: 2, Item: vi.fn((i: number) => ({ index: i, value: `item${i}` })) } as unknown as Record<string, unknown>
    const col = new CollectionProxy(raw, (r) => r)
    const item = col.item(1)
    expect(item.value).toBe("item1")
    expect(raw.Item).toHaveBeenCalledWith(1)
  })
})

describe("Error resilience", () => {
  function makeThrowingRaw(): Record<string, unknown> {
    return new Proxy({} as Record<string, unknown>, {
      get: () => { throw new Error("COM disconnected") },
    })
  }

  it("all proxy methods return safe defaults on COM errors", () => {
    const badRaw = makeThrowingRaw()
    const session = {
      wrapRange: (r: Record<string, unknown>) => new RangeProxy(r, session),
      getSelectionProxy: () => ({ raw: {} as Record<string, unknown> } as unknown as ISelectionProxy),
    }

    const doc = new DocumentProxy(badRaw, session)
    expect(doc.getName()).toBe("未命名文档")
    expect(doc.getFullName()).toBeUndefined()
    expect(doc.getSaved()).toBe(true)
    expect(() => doc.getContent()).not.toThrow()
    expect(() => doc.getParagraphs()).not.toThrow()
    expect(() => doc.save()).not.toThrow()
    expect(() => doc.select()).not.toThrow()
  })

  it("SelectionProxy returns safe defaults on COM errors", () => {
    const session = {
      wrapRange: (r: Record<string, unknown>) => new RangeProxy(r, session),
    }
    const sel = new SelectionProxy(makeThrowingRaw(), session)
    expect(sel.getStart()).toBe(0)
    expect(sel.getStoryType()).toBe(1)
    expect(() => sel.typeText("hi")).not.toThrow()
    expect(() => sel.typeParagraph()).not.toThrow()
    expect(() => sel.getRange()).not.toThrow()
  })

  it("RangeProxy returns safe defaults on COM errors", () => {
    const session = {
      wrapRange: (r: Record<string, unknown>) => new RangeProxy(r, session),
    }
    const rng = new RangeProxy(makeThrowingRaw(), session)
    expect(rng.getText()).toBe("")
    expect(rng.getBold()).toBeUndefined()
    expect(() => rng.select()).not.toThrow()
    expect(() => rng.duplicate()).not.toThrow()
  })
})

describe("WeakMap range cache", () => {
  it("wrapRange returns same proxy for same raw object", () => {
    const session = {
      _cache: new WeakMap<Record<string, unknown>, IRangeProxy>(),
      wrapRange(this: { _cache: WeakMap<Record<string, unknown>, IRangeProxy> }, raw: Record<string, unknown>): IRangeProxy {
        let p = this._cache.get(raw)
        if (!p) {
          p = new RangeProxy(raw, this as unknown as { wrapRange: (r: Record<string, unknown>) => IRangeProxy })
          this._cache.set(raw, p)
        }
        return p
      },
    }

    const raw = { Text: "cached" } as Record<string, unknown>
    const a = session.wrapRange(raw)
    const b = session.wrapRange(raw)
    expect(a).toBe(b)
  })
})
