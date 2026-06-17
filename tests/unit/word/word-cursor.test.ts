import { describe, it, expect, vi } from "vitest"
import { WordCursor } from "../../../src/word/word-cursor.js"
import { createMockSession } from "../test-helpers.js"
import { MockSelectionProxy, MockDocumentProxy, MockRangeProxy } from "../../../src/word/com-proxy/com-proxy.mock.js"

function cursorWithSel(overrides?: Record<string, unknown>) {
  const sel = {
    Start: 0, End: 10, Type: 2,
    TypeText: vi.fn(),
    TypeParagraph: vi.fn(),
    TypeBackspace: vi.fn(),
    Collapse: vi.fn(),
    MoveStart: vi.fn(),
    EndKey: vi.fn(),
    Style: "",
    Font: {},
    ParagraphFormat: { LeftIndent: 0, Alignment: 0 },
    Range: {
      Text: "selected", Font: {}, Start: 0, End: 10,
      Hyperlinks: { Add: vi.fn() },
      Shading: { BackgroundPatternColor: 0 },
      Select: vi.fn(),
      ConvertToTable: vi.fn(),
      InsertFile: vi.fn(),
      ListFormat: {
        ApplyBulletDefault: vi.fn(),
        ApplyNumberDefault: vi.fn(),
        RemoveNumbers: vi.fn(),
        IncreaseIndent: vi.fn(),
      },
    },
    InlineShapes: { AddHorizontalLineStandard: vi.fn(), AddPicture: vi.fn() },
    Find: {
      ClearFormatting: vi.fn(),
      MatchCase: false, MatchWholeWord: false,
      Style: "", Text: "", Forward: true, Wrap: 0, Format: false,
      Execute: vi.fn(),
      Replacement: {
        ClearFormatting: vi.fn(),
        ParagraphFormat: { FirstLineIndent: 0 },
        Text: "",
      },
    },
    Information: vi.fn(() => false),
    Tables: { Item: vi.fn(), Count: 0 },
    ShapeRange: { Count: 0 },
    InsertBreak: vi.fn(),
    HomeKey: vi.fn(),
    WholeStory: vi.fn(),
    Delete: vi.fn(),
    Copy: vi.fn(),
    Cut: vi.fn(),
    Paste: vi.fn(),
    GoTo: vi.fn(),
    Select: vi.fn(),
    Expand: vi.fn(),
    ...overrides,
  } as unknown as Record<string, unknown>
  const doc = {
    Name: "test.docx",
    FullName: "C:\\test.docx",
    Content: { End: 100, Start: 0, Text: "" },
    Range: vi.fn(() => ({
      End: 100, Start: 0, Select: vi.fn(),
      Hyperlinks: { Add: vi.fn() },
      Shading: { BackgroundPatternColor: 0 },
    })),
    Paragraphs: { Count: 5, Item: vi.fn((i: number) => ({ Range: { Select: vi.fn(), Start: i * 20, End: i * 20 + 10 } })) },
    Tables: { Count: 1, Item: vi.fn() },
    Bookmarks: { Count: 0, Item: vi.fn(), Add: vi.fn() },
    Styles: { Count: 0, Item: vi.fn() },
    Sections: { Count: 1, Item: vi.fn() },
    Saved: true,
    ExportAsFixedFormat: vi.fn(),
    TrackRevisions: false,
    Revisions: { Count: 0, AcceptAll: vi.fn(), RejectAll: vi.fn() },
    Comments: { Count: 0, Item: vi.fn() },
    Hyperlinks: { Add: vi.fn() },
    Footnotes: { Add: vi.fn() },
    Select: vi.fn(),
    Undo: vi.fn(),
    Redo: vi.fn(),
  } as unknown as Record<string, unknown>
  const session = createMockSession({}, doc, sel)
  const selProxy = session.getSelectionProxy() as MockSelectionProxy
  const docProxy = session.getDocProxy() as MockDocumentProxy
  return { session, cursor: new WordCursor(session), sel, doc, selProxy, docProxy }
}

describe("WordCursor — findText", () => {
  it("returns position string when text is found", async () => {
    const { cursor, sel, selProxy } = cursorWithSel()
    sel.Type = 0
    sel.Range.Text = "found text"
    sel.Range.Start = 42
    const find = selProxy.raw.Find as Record<string, unknown>
    const executeSpy = find.Execute as ReturnType<typeof vi.fn>
    const result = await cursor.findText("keyword")
    expect(executeSpy).toHaveBeenCalledWith("keyword", false, false, false, false, false, true, 1, false, "", 0)
    expect(result).toContain("Found at position 42")
    expect(result).toContain("found text")
  })

  it("returns empty string when text is not found", async () => {
    const { cursor, selProxy } = cursorWithSel()
    const find = selProxy.raw.Find as ReturnType<typeof vi.fn>
    Object.defineProperty(selProxy.raw, "Type", { value: 1, configurable: true })
    const result = await cursor.findText("nonexistent")
    expect(result).toBe("")
  })

  it("forwards matchCase, matchWholeWord, direction backward, wrap false", async () => {
    const { cursor, selProxy, sel } = cursorWithSel()
    sel.Type = 0
    const find = selProxy.raw.Find as Record<string, unknown>
    const executeSpy = find.Execute as ReturnType<typeof vi.fn>
    await cursor.findText("keyword", {
      matchCase: true,
      matchWholeWord: true,
      direction: "backward",
      wrap: false,
    })
    expect(find.MatchCase).toBe(true)
    expect(find.MatchWholeWord).toBe(true)
    expect(executeSpy).toHaveBeenCalledWith("keyword", true, true, false, false, false, false, 0, false, "", 0)
  })
})

describe("WordCursor — findReplace", () => {
  it("uses replaceAll mode by default and sets Replacement.Text", async () => {
    const { cursor, selProxy } = cursorWithSel()
    const find = selProxy.raw.Find as Record<string, unknown>
    const executeSpy = find.Execute as ReturnType<typeof vi.fn>
    await cursor.findReplace("foo", "bar")
    expect((find.Replacement as Record<string, unknown>).Text).toBe("bar")
    expect(executeSpy).toHaveBeenCalledWith("foo", false, false, false, false, false, true, 1, false, "bar", 2)
  })

  it("uses replaceOne mode when replaceAll:false", async () => {
    const { cursor, selProxy } = cursorWithSel()
    const find = selProxy.raw.Find as Record<string, unknown>
    const executeSpy = find.Execute as ReturnType<typeof vi.fn>
    await cursor.findReplace("foo", "bar", { replaceAll: false })
    expect(executeSpy).toHaveBeenCalledWith("foo", false, false, false, false, false, true, 1, false, "bar", 1)
  })

  it("forwards matchCase and matchWholeWord", async () => {
    const { cursor, selProxy } = cursorWithSel()
    const find = selProxy.raw.Find as Record<string, unknown>
    const executeSpy = find.Execute as ReturnType<typeof vi.fn>
    await cursor.findReplace("Foo", "Bar", { matchCase: true, matchWholeWord: true })
    expect(executeSpy).toHaveBeenCalledWith("Foo", true, true, false, false, false, true, 1, false, "Bar", 2)
  })
})

describe("WordCursor — goTo", () => {
  it("calls goTo with page/first by default", async () => {
    const { cursor, selProxy } = cursorWithSel()
    await cursor.goTo()
    expect(selProxy.goTo).toHaveBeenCalledWith(1, 1)
  })

  it("handles 'end' special case", async () => {
    const { cursor, docProxy } = cursorWithSel()
    const rangeSelect = vi.fn()
    vi.mocked(docProxy.getRange).mockReturnValue({ select: rangeSelect } as unknown as MockRangeProxy)
    await cursor.goTo("end")
    expect(docProxy.getRange).toHaveBeenCalled()
    expect(rangeSelect).toHaveBeenCalled()
  })

  it("routes section/next correctly", async () => {
    const { cursor, selProxy } = cursorWithSel()
    await cursor.goTo("section", "next")
    expect(selProxy.goTo).toHaveBeenCalledWith(2, 2)
  })
})

describe("WordCursor — goToParagraph", () => {
  it("selects the paragraph range", async () => {
    const { cursor, docProxy } = cursorWithSel()
    const range = { Select: vi.fn() }
    const rawParas = docProxy.raw.Paragraphs as Record<string, unknown>
    ;(rawParas.Item as ReturnType<typeof vi.fn>).mockReturnValue({ Range: range })
    await cursor.goToParagraph(1)
    expect(range.Select).toHaveBeenCalled()
  })

  it("throws on out-of-range index", async () => {
    const { cursor } = cursorWithSel()
    await expect(cursor.goToParagraph(999)).rejects.toThrow("Paragraph index 999 out of range")
    await expect(cursor.goToParagraph(0)).rejects.toThrow("Paragraph index 0 out of range")
  })
})

describe("WordCursor — selectAll / selectCurrentWord / selectCurrentParagraph", () => {
  it("selectAll calls wholeStory", async () => {
    const { cursor, selProxy } = cursorWithSel()
    await cursor.selectAll()
    expect(selProxy.wholeStory).toHaveBeenCalled()
  })

  it("selectCurrentWord calls Expand(2)", async () => {
    const { cursor, selProxy } = cursorWithSel()
    await cursor.selectCurrentWord()
    expect(selProxy.expand).toHaveBeenCalledWith(2)
  })

  it("selectCurrentParagraph calls Expand(4)", async () => {
    const { cursor, selProxy } = cursorWithSel()
    await cursor.selectCurrentParagraph()
    expect(selProxy.expand).toHaveBeenCalledWith(4)
  })
})

describe("WordCursor — selectText", () => {
  it("calls doc.getRange then range.select", async () => {
    const { cursor, docProxy } = cursorWithSel()
    const mockRange = new MockRangeProxy()
    vi.mocked(docProxy.getRange).mockReturnValue(mockRange)
    await cursor.selectText(5, 20)
    expect(docProxy.getRange).toHaveBeenCalledWith(5, 25)
    expect(mockRange.select).toHaveBeenCalled()
  })
})

describe("WordCursor — deleteSelection", () => {
  it("calls delete when selection exists", async () => {
    const { cursor, selProxy } = cursorWithSel()
    await cursor.deleteSelection()
    expect(selProxy.delete).toHaveBeenCalled()
  })

  it("throws when nothing selected", async () => {
    const { cursor } = cursorWithSel({ Start: 0, End: 0, Type: 1 })
    await expect(cursor.deleteSelection()).rejects.toThrow("No text is selected")
  })
})

describe("WordCursor — backspace", () => {
  it("calls typeBackspace once by default", async () => {
    const { cursor, selProxy } = cursorWithSel()
    await cursor.backspace()
    expect(selProxy.typeBackspace).toHaveBeenCalledTimes(1)
  })

  it("calls typeBackspace N times with count", async () => {
    const { cursor, selProxy } = cursorWithSel()
    await cursor.backspace(5)
    expect(selProxy.typeBackspace).toHaveBeenCalledTimes(5)
  })
})

describe("WordCursor — copy / cut / paste", () => {
  it("copy calls Copy on selection", async () => {
    const { cursor, selProxy } = cursorWithSel()
    await cursor.copy()
    expect(selProxy.copy).toHaveBeenCalled()
  })

  it("copy throws when nothing selected", async () => {
    const { cursor } = cursorWithSel({ Start: 0, End: 0, Type: 1 })
    await expect(cursor.copy()).rejects.toThrow("No text is selected")
  })

  it("cut calls Cut on selection", async () => {
    const { cursor, selProxy } = cursorWithSel()
    await cursor.cut()
    expect(selProxy.cut).toHaveBeenCalled()
  })

  it("paste calls Paste then goToEnd", async () => {
    const { cursor, selProxy, docProxy } = cursorWithSel()
    const rangeSelect = vi.fn()
    vi.mocked(docProxy.getRange).mockReturnValue({ select: rangeSelect } as unknown as MockRangeProxy)
    await cursor.paste()
    expect(selProxy.paste).toHaveBeenCalled()
    expect(docProxy.getRange).toHaveBeenCalled()
    expect(rangeSelect).toHaveBeenCalled()
  })
})

describe("WordCursor — undo / redo", () => {
  it("undo calls doc.Undo once by default", async () => {
    const { cursor, docProxy } = cursorWithSel()
    await cursor.undo()
    expect(docProxy.undo).toHaveBeenCalledTimes(1)
  })

  it("undo calls doc.Undo N times", async () => {
    const { cursor, docProxy } = cursorWithSel()
    await cursor.undo(3)
    expect(docProxy.undo).toHaveBeenCalledTimes(3)
  })

  it("redo calls doc.Redo once by default", async () => {
    const { cursor, docProxy } = cursorWithSel()
    await cursor.redo()
    expect(docProxy.redo).toHaveBeenCalledTimes(1)
  })

  it("redo calls doc.Redo N times", async () => {
    const { cursor, docProxy } = cursorWithSel()
    await cursor.redo(3)
    expect(docProxy.redo).toHaveBeenCalledTimes(3)
  })
})

describe("WordCursor — getCursorInfo", () => {
  it("returns no selection when Type=1", async () => {
    const { cursor } = cursorWithSel({ Start: 0, End: 0, Type: 1 })
    const info = await cursor.getCursorInfo()
    expect(info.hasSelection).toBe(false)
    expect(info.selectedText).toBe("")
  })

  it("returns selection info when Type≠1", async () => {
    const { cursor } = cursorWithSel()
    const info = await cursor.getCursorInfo()
    expect(info.hasSelection).toBe(true)
    expect(info.selectedText).toBe("selected")
    expect(info.start).toBe(0)
    expect(info.end).toBe(10)
  })
})

describe("WordCursor — insertFile", () => {
  it("calls Range.InsertFile", async () => {
    const { cursor, selProxy } = cursorWithSel()
    const mockRange = new MockRangeProxy()
    vi.mocked(selProxy.getRange).mockReturnValue(mockRange)
    await cursor.insertFile("C:\\test.docx")
    expect(mockRange.insertFile).toHaveBeenCalledWith("C:\\test.docx")
  })
})
