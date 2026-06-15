import { describe, it, expect, vi } from "vitest"
import { WordTextEditor } from "../../../src/word/word-text-editor.js"
import { createMockSession } from "../test-helpers.js"

function makeSession() {
  return createMockSession()
}

function editorWithSel(overrides?: Record<string, unknown>) {
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
  const textColumnsMock = { SetCount: vi.fn(), Spacing: 0 }
  const doc = {
    Name: "test.docx",
    FullName: "C:\\test.docx",
    Content: { End: 100, Start: 0, Text: "" },
    Range: vi.fn(() => ({
      End: 100, Start: 0, Select: vi.fn(),
      Hyperlinks: { Add: vi.fn() },
      Shading: { BackgroundPatternColor: 0 },
      InsertFile: vi.fn(),
    })),
    Paragraphs: { Count: 5, Item: vi.fn((i: number) => ({ Range: { Select: vi.fn(), Start: i * 20, End: i * 20 + 10 } })) },
    Tables: { Count: 1, Item: vi.fn() },
    Bookmarks: { Count: 0, Item: vi.fn(), Add: vi.fn() },
    Styles: { Count: 0, Item: vi.fn() },
    Sections: { Count: 1, Item: vi.fn(() => ({ PageSetup: { TextColumns: textColumnsMock } })) },
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
  const app = {
    Selection: sel,
    ScreenUpdating: true,
    ActiveDocument: doc,
    ScreenRefresh: vi.fn(),
  } as Record<string, unknown>
  const session = createMockSession({ application: app, activeDoc: doc })
  return { session, editor: new WordTextEditor(session), sel, doc, app }
}

describe("WordTextEditor — insertParagraph", () => {
  it("calls TypeParagraph N times", async () => {
    const { editor, sel } = editorWithSel()
    await editor.insertParagraph(3)
    expect(sel.TypeParagraph).toHaveBeenCalledTimes(3)
  })

  it("defaults to 1", async () => {
    const { editor, sel } = editorWithSel()
    await editor.insertParagraph()
    expect(sel.TypeParagraph).toHaveBeenCalledTimes(1)
  })
})

describe("WordTextEditor — insertPageBreak", () => {
  it("calls InsertBreak(7)", async () => {
    const { editor, sel } = editorWithSel()
    await editor.insertPageBreak()
    expect(sel.InsertBreak).toHaveBeenCalledWith(7)
  })
})

describe("WordTextEditor — insertHorizontalLine", () => {
  it("calls AddHorizontalLineStandard", async () => {
    const { editor, sel } = editorWithSel()
    await editor.insertHorizontalLine()
    const spy = (sel.InlineShapes as Record<string, unknown>).AddHorizontalLineStandard as ReturnType<typeof vi.fn>
    expect(spy).toHaveBeenCalledTimes(1)
  })
})

describe("WordTextEditor — findText", () => {
  it("calls Execute with correct parameters", async () => {
    const { editor, sel } = editorWithSel()
    const executeSpy = sel.Find.Execute as ReturnType<typeof vi.fn>
    executeSpy.mockReturnValue(true)
    Object.defineProperty(sel, "Type", { value: 2, configurable: true })
    const result = await editor.findText("keyword", { matchCase: true, matchWholeWord: true, direction: "forward", wrap: false })
    expect(executeSpy).toHaveBeenCalledWith("keyword", true, true, false, false, false, true, 0, false, "", 0)
    expect(result).toContain("Found")
  })

  it("returns empty string when not found", async () => {
    const { editor, sel } = editorWithSel()
    const executeSpy = sel.Find.Execute as ReturnType<typeof vi.fn>
    Object.defineProperty(sel, "Type", { value: 1, configurable: true })
    executeSpy.mockReturnValue(false)
    const result = await editor.findText("nonexistent")
    expect(result).toBe("")
  })
})

describe("WordTextEditor — findReplace", () => {
  function setup() {
    const { editor, sel } = editorWithSel()
    const find = sel.Find as Record<string, unknown>
    const executeSpy = find.Execute as ReturnType<typeof vi.fn>
    return { editor, find, executeSpy }
  }

  it("calls Execute with replaceAll mode by default", async () => {
    const { editor, executeSpy } = setup()
    await editor.findReplace("foo", "bar")
    expect(executeSpy).toHaveBeenCalledWith("foo", false, false, false, false, false, true, 1, false, "bar", 2)
  })

  it("uses replaceOne mode when replaceAll:false", async () => {
    const { editor, executeSpy } = setup()
    await editor.findReplace("foo", "bar", { replaceAll: false })
    expect(executeSpy).toHaveBeenCalledWith("foo", false, false, false, false, false, true, 1, false, "bar", 1)
  })

  it("passes matchCase and matchWholeWord", async () => {
    const { editor, executeSpy } = setup()
    await editor.findReplace("Foo", "Bar", { matchCase: true, matchWholeWord: true })
    expect(executeSpy).toHaveBeenCalledWith("Foo", true, true, false, false, false, true, 1, false, "Bar", 2)
  })

  it("uses wrap=0 when wrap:false", async () => {
    const { editor, executeSpy } = setup()
    await editor.findReplace("foo", "bar", { wrap: false })
    expect(executeSpy).toHaveBeenCalledWith("foo", false, false, false, false, false, true, 0, false, "bar", 2)
  })
})

describe("WordTextEditor — goTo", () => {
  it("calls GoTo with page/first by default", async () => {
    const { editor, sel } = editorWithSel()
    await editor.goTo()
    expect(sel.GoTo).toHaveBeenCalledWith(1, 1)
  })

  it("handles 'end' special case", async () => {
    const { editor, sel } = editorWithSel()
    await editor.goTo("end")
    expect(sel.EndKey).toHaveBeenCalledWith(6)
    expect(sel.Collapse).toHaveBeenCalledWith(0)
  })

  it("routes section/next correctly", async () => {
    const { editor, sel } = editorWithSel()
    await editor.goTo("section", "next")
    expect(sel.GoTo).toHaveBeenCalledWith(2, 2)
  })
})

describe("WordTextEditor — goToParagraph", () => {
  it("selects the paragraph range", async () => {
    const { editor, doc } = editorWithSel()
    const itemSpy = doc.Paragraphs.Item as ReturnType<typeof vi.fn>
    const range = { Select: vi.fn() }
    itemSpy.mockReturnValue({ Range: range })
    await editor.goToParagraph(1)
    expect(range.Select).toHaveBeenCalled()
  })

  it("throws on out-of-range index", async () => {
    const { editor } = editorWithSel()
    await expect(editor.goToParagraph(999)).rejects.toThrow("Paragraph index 999 out of range")
    await expect(editor.goToParagraph(0)).rejects.toThrow("Paragraph index 0 out of range")
  })
})

describe("WordTextEditor — selectAll/selectText/selectWord/selectParagraph", () => {
  it("selectAll calls WholeStory", async () => {
    const { editor, sel } = editorWithSel()
    await editor.selectAll()
    expect(sel.WholeStory).toHaveBeenCalled()
  })

  it("selectText selects range in doc", async () => {
    const { editor, doc } = editorWithSel()
    const range = { Select: vi.fn() }
    ;(doc.Range as ReturnType<typeof vi.fn>).mockReturnValue(range)
    await editor.selectText(5, 20)
    expect(doc.Range).toHaveBeenCalledWith(5, 25)
    expect(range.Select).toHaveBeenCalled()
  })

  it("selectCurrentWord calls Expand(2)", async () => {
    const { editor, sel } = editorWithSel()
    await editor.selectCurrentWord()
    expect(sel.Expand).toHaveBeenCalledWith(2)
  })

  it("selectCurrentParagraph calls Expand(4)", async () => {
    const { editor, sel } = editorWithSel()
    await editor.selectCurrentParagraph()
    expect(sel.Expand).toHaveBeenCalledWith(4)
  })
})

describe("WordTextEditor — deleteSelection / backspace / copy / cut / paste", () => {
  it("deleteSelection calls Delete on selection", async () => {
    const { editor, sel } = editorWithSel()
    await editor.deleteSelection()
    expect(sel.Delete).toHaveBeenCalled()
  })

  it("deleteSelection throws when nothing selected", async () => {
    const { editor } = editorWithSel({ Start: 0, End: 0, Type: 1 })
    await expect(editor.deleteSelection()).rejects.toThrow("No text is selected")
  })

  it("backspace calls TypeBackspace once by default", async () => {
    const { editor, sel } = editorWithSel()
    await editor.backspace()
    expect(sel.TypeBackspace).toHaveBeenCalledTimes(1)
  })

  it("backspace calls TypeBackspace N times", async () => {
    const { editor, sel } = editorWithSel()
    await editor.backspace(5)
    expect(sel.TypeBackspace).toHaveBeenCalledTimes(5)
  })

  it("copy calls Copy on selection", async () => {
    const { editor, sel } = editorWithSel()
    await editor.copy()
    expect(sel.Copy).toHaveBeenCalled()
  })

  it("copy throws when nothing selected", async () => {
    const { editor } = editorWithSel({ Start: 0, End: 0, Type: 1 })
    await expect(editor.copy()).rejects.toThrow("No text is selected")
  })

  it("cut calls Cut on selection", async () => {
    const { editor, sel } = editorWithSel()
    await editor.cut()
    expect(sel.Cut).toHaveBeenCalled()
  })

  it("paste calls Paste then EndKey", async () => {
    const { editor, sel } = editorWithSel()
    await editor.paste()
    expect(sel.Paste).toHaveBeenCalled()
    expect(sel.EndKey).toHaveBeenCalledWith(6)
  })
})

describe("WordTextEditor — undo / redo", () => {
  it("undo calls doc.Undo once by default", async () => {
    const { editor, doc } = editorWithSel()
    await editor.undo()
    expect(doc.Undo).toHaveBeenCalledTimes(1)
  })

  it("undo calls doc.Undo N times", async () => {
    const { editor, doc } = editorWithSel()
    await editor.undo(3)
    expect(doc.Undo).toHaveBeenCalledTimes(3)
  })

  it("redo calls doc.Redo once by default", async () => {
    const { editor, doc } = editorWithSel()
    await editor.redo()
    expect(doc.Redo).toHaveBeenCalledTimes(1)
  })
})

describe("WordTextEditor — getCursorInfo", () => {
  it("returns no selection when Type=1", async () => {
    const { editor } = editorWithSel({ Start: 0, End: 0, Type: 1 })
    const info = await editor.getCursorInfo()
    expect(info.hasSelection).toBe(false)
    expect(info.selectedText).toBe("")
  })

  it("returns selection text when Type≠1", async () => {
    const { editor } = editorWithSel()
    const info = await editor.getCursorInfo()
    expect(info.hasSelection).toBe(true)
    expect(info.selectedText).toBe("selected")
    expect(info.start).toBe(0)
    expect(info.end).toBe(10)
  })
})

describe("WordTextEditor — insertList", () => {
  function listSetup() {
    const { editor, sel, app } = editorWithSel()
    const lf = (sel.Range as Record<string, unknown>).ListFormat as Record<string, unknown>
    return { editor, sel, lf, app }
  }

  it("bullet list calls ApplyBulletDefault + RemoveNumbers", async () => {
    const { editor, lf } = listSetup()
    const spy = lf.ApplyBulletDefault as ReturnType<typeof vi.fn>
    await editor.insertList("bullet", ["a", "b"])
    expect(spy).toHaveBeenCalled()
    expect(lf.RemoveNumbers).toHaveBeenCalled()
  })

  it("number list calls ApplyNumberDefault + RemoveNumbers", async () => {
    const { editor, lf } = listSetup()
    const spy = lf.ApplyNumberDefault as ReturnType<typeof vi.fn>
    await editor.insertList("number", ["1", "2"])
    expect(spy).toHaveBeenCalled()
    expect(lf.RemoveNumbers).toHaveBeenCalled()
  })

  it("types each item with TypeParagraph between", async () => {
    const { editor, sel } = listSetup()
    await editor.insertList("bullet", ["x", "y", "z"])
    expect(sel.TypeText).toHaveBeenCalledTimes(3)
    expect(sel.TypeParagraph).toHaveBeenCalledTimes(3)
  })
})

describe("WordTextEditor — addHyperlink", () => {
  it("adds hyperlink with existing selection", async () => {
    const { editor, doc } = editorWithSel()
    await editor.addHyperlink("text", "https://example.com")
    expect(doc.Hyperlinks.Add).toHaveBeenCalled()
  })

  it("adds hyperlink without selection uses collapsed position", async () => {
    const { editor, doc, sel } = editorWithSel({ Start: 0, End: 0, Type: 1 })
    await editor.addHyperlink("link", "https://example.com", "sub", "tip")
    expect(sel.EndKey).toHaveBeenCalledWith(6)
    expect(doc.Hyperlinks.Add).toHaveBeenCalledWith(
      expect.anything(), "https://example.com", "sub", "tip", "link"
    )
  })
})

describe("WordTextEditor — addFootnote", () => {
  it("adds footnote via COM when successful", async () => {
    const { editor, doc, sel } = editorWithSel()
    await editor.addFootnote("note text")
    expect(doc.Footnotes.Add).toHaveBeenCalled()
  })

  it("falls back to inline text when COM fails", async () => {
    const { editor, doc, sel } = editorWithSel()
    ;(doc.Footnotes.Add as ReturnType<typeof vi.fn>).mockImplementation(() => { throw new Error("COM fail") })
    await editor.addFootnote("note text")
    expect(sel.TypeText).toHaveBeenCalledWith(expect.stringContaining("note text"))
  })
})

describe("WordTextEditor — insertSectionBreak", () => {
  it("defaults to nextPage (8)", async () => {
    const { editor, sel } = editorWithSel()
    await editor.insertSectionBreak()
    expect(sel.InsertBreak).toHaveBeenCalledWith(8)
  })

  it("maps continuous to 9", async () => {
    const { editor, sel } = editorWithSel()
    await editor.insertSectionBreak("continuous")
    expect(sel.InsertBreak).toHaveBeenCalledWith(9)
  })
})

describe("WordTextEditor — setColumns", () => {
  it("calls TextColumns.SetCount with count", async () => {
    const tc = { SetCount: vi.fn(), Spacing: 0 }
    const { editor } = editorWithSel()
    const doc = (editor as unknown as { session: { activeDoc: Record<string, unknown> } }).session.activeDoc!
    const si = doc.Sections as unknown as { Count: number; Item: (i: number) => Record<string, unknown> }
    si.Item = vi.fn(() => ({ PageSetup: { TextColumns: tc } })) as unknown as (i: number) => Record<string, unknown>
    await editor.setColumns(2)
    expect(tc.SetCount).toHaveBeenCalledWith(2)
  })

  it("sets spacing when provided", async () => {
    const tc = { SetCount: vi.fn(), Spacing: 0 }
    const { editor } = editorWithSel()
    const doc = (editor as unknown as { session: { activeDoc: Record<string, unknown> } }).session.activeDoc!
    const si = doc.Sections as unknown as { Count: number; Item: (i: number) => Record<string, unknown> }
    si.Item = vi.fn(() => ({ PageSetup: { TextColumns: tc } })) as unknown as (i: number) => Record<string, unknown>
    await editor.setColumns(2, 1.5)
    expect(tc.Spacing).toBe(43)
  })
})

describe("WordTextEditor — insertFile", () => {
  it("calls Range.InsertFile", async () => {
    const { editor, sel } = editorWithSel()
    const spy = (sel.Range as Record<string, unknown>).InsertFile as ReturnType<typeof vi.fn>
    await editor.insertFile("C:\\test.docx")
    expect(spy).toHaveBeenCalledWith("C:\\test.docx")
  })
})
