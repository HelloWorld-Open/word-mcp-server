import { describe, it, expect, vi, beforeEach } from "vitest"
import { WordTextEditor } from "../../../src/word/word-text-editor.js"
import type { IWordSession } from "../../../src/word/session.js"

function createMockDoc(): Record<string, unknown> {
  return {
    Content: { End: 0, Text: "" },
    Range: vi.fn(() => ({ End: 0, Select: vi.fn() })),
    Paragraphs: { Count: 0, Item: vi.fn() },
    Application: { Selection: {} },
    Tables: { Count: 0, Item: vi.fn() },
    Bookmarks: { Count: 0, Item: vi.fn() },
  } as unknown as Record<string, unknown>
}

function createMockSession(mockDoc?: Record<string, unknown>): IWordSession {
  const sel = {
    Start: 0, End: 0, StoryType: 1,
    TypeText: vi.fn(),
    TypeParagraph: vi.fn(),
    Collapse: vi.fn(),
    MoveStart: vi.fn(),
    Style: "",
    Font: {},
    ParagraphFormat: { LeftIndent: 0 },
    Range: { Text: "", Font: {}, Hyperlinks: { Add: vi.fn() }, Select: vi.fn() },
    InlineShapes: { AddHorizontalLineStandard: vi.fn() },
    Find: {
      ClearFormatting: vi.fn(),
      MatchCase: false,
      MatchWholeWord: false,
      Style: "",
      Text: "",
      Forward: true,
      Wrap: 0,
      Execute: vi.fn(),
    },
    Information: vi.fn(() => false),
    Tables: { Item: vi.fn() },
    ShapeRange: { Count: 0 },
    InsertBreak: vi.fn(),
    HomeKey: vi.fn(),
  } as unknown as Record<string, unknown>

  const doc = mockDoc ?? createMockDoc()

  return {
    application: { Selection: sel, ScreenUpdating: true, ActiveDocument: doc } as Record<string, unknown>,
    activeDoc: doc,
    activeDocPath: null,
    wasInNonBody: false,
    setActiveDoc: vi.fn(),
    setActiveDocPath: vi.fn(),
    ensureAlive: vi.fn(),
    isAlive: () => true,
    start: vi.fn(),
    quit: vi.fn(),
    setOnLog: vi.fn(),
    setScreenUpdating: vi.fn(),
    healthCheck: () => true,
    recover: async () => {},
    comCall: <T>(fn: () => T) => fn(),
    markHealthy: vi.fn(),
    markUnhealthy: vi.fn(),
    isUnhealthy: () => false,
  }
}

describe("WordTextEditor.splitIntoBatches", () => {
  let editor: WordTextEditor

  beforeEach(() => {
    editor = new WordTextEditor(createMockSession())
  })

  it("returns single batch for short text without sentence enders", () => {
    const batches = (editor as any).splitIntoBatches("hello world")
    expect(batches).toEqual(["hello world"])
  })

  it("splits on Chinese period", () => {
    const batches = (editor as any).splitIntoBatches("第一句。第二句。第三句")
    expect(batches).toEqual(["第一句。", "第二句。", "第三句"])
  })

  it("splits on Chinese exclamation", () => {
    const batches = (editor as any).splitIntoBatches("你好！再见！")
    expect(batches).toEqual(["你好！", "再见！"])
  })

  it("splits on English period", () => {
    const batches = (editor as any).splitIntoBatches("Hello. World. Done.")
    expect(batches).toEqual(["Hello.", " World.", " Done."])
  })

  it("splits on English question mark", () => {
    const batches = (editor as any).splitIntoBatches("A? B? C?")
    expect(batches).toEqual(["A?", " B?", " C?"])
  })

  it("splits on newline", () => {
    const batches = (editor as any).splitIntoBatches("line1\nline2\nline3")
    expect(batches).toEqual(["line1\n", "line2\n", "line3"])
  })

  it("forces split at 500+ chars even without sentence enders", () => {
    const longText = "a".repeat(503)
    const batches = (editor as any).splitIntoBatches(longText)
    expect(batches).toHaveLength(2)
    expect(batches[0]).toHaveLength(501)
    expect(batches[1]).toHaveLength(2)
  })

  it("handles empty string", () => {
    const batches = (editor as any).splitIntoBatches("")
    expect(batches).toEqual([""])
  })
})

describe("WordTextEditor COM operations", () => {
  it("typeText instant mode calls TypeText once", async () => {
    const session = createMockSession()
    const editor = new WordTextEditor(session)
    const sel = (session.application as Record<string, unknown>).Selection as Record<string, unknown>
    const typeTextSpy = sel.TypeText as ReturnType<typeof vi.fn>

    await editor.typeText("hello world", "instant")

    expect(typeTextSpy).toHaveBeenCalledTimes(1)
    expect(typeTextSpy).toHaveBeenCalledWith("hello world")
  })

  it("typeText smooth mode splits into batches with delays", async () => {
    const session = createMockSession()
    const editor = new WordTextEditor(session)
    const sel = (session.application as Record<string, unknown>).Selection as Record<string, unknown>
    const typeTextSpy = sel.TypeText as ReturnType<typeof vi.fn>
    const sleepSpy = vi.spyOn(editor as any, "sleep").mockResolvedValue(undefined)

    await editor.typeText("A。B。C", "smooth")

    expect(typeTextSpy).toHaveBeenCalledTimes(3)
    expect(typeTextSpy.mock.calls[0][0]).toBe("A。")
    expect(typeTextSpy.mock.calls[1][0]).toBe("B。")
    expect(typeTextSpy.mock.calls[2][0]).toBe("C")
    expect(sleepSpy).toHaveBeenCalledTimes(2)
  })

  it("insertParagraph calls TypeParagraph N times", async () => {
    const session = createMockSession()
    const editor = new WordTextEditor(session)
    const sel = (session.application as Record<string, unknown>).Selection as Record<string, unknown>
    const typeParaSpy = sel.TypeParagraph as ReturnType<typeof vi.fn>

    await editor.insertParagraph(3)

    expect(typeParaSpy).toHaveBeenCalledTimes(3)
  })

  it("insertParagraph defaults to 1", async () => {
    const session = createMockSession()
    const editor = new WordTextEditor(session)
    const sel = (session.application as Record<string, unknown>).Selection as Record<string, unknown>
    const typeParaSpy = sel.TypeParagraph as ReturnType<typeof vi.fn>

    await editor.insertParagraph()

    expect(typeParaSpy).toHaveBeenCalledTimes(1)
  })

  it("insertPageBreak calls InsertBreak(7)", async () => {
    const session = createMockSession()
    const editor = new WordTextEditor(session)
    const sel = (session.application as Record<string, unknown>).Selection as Record<string, unknown>
    const insertBreakSpy = sel.InsertBreak as ReturnType<typeof vi.fn>

    await editor.insertPageBreak()

    expect(insertBreakSpy).toHaveBeenCalledWith(7)
  })

  it("insertHorizontalLine calls AddHorizontalLineStandard", async () => {
    const session = createMockSession()
    const editor = new WordTextEditor(session)
    const sel = (session.application as Record<string, unknown>).Selection as Record<string, unknown>
    const shapes = sel.InlineShapes as { AddHorizontalLineStandard: ReturnType<typeof vi.fn> }
    const spy = shapes.AddHorizontalLineStandard as ReturnType<typeof vi.fn>

    await editor.insertHorizontalLine()

    expect(spy).toHaveBeenCalledTimes(1)
  })

  it("findText calls Execute with correct parameters", async () => {
    const session = createMockSession()
    const editor = new WordTextEditor(session)
    const sel = (session.application as Record<string, unknown>).Selection as Record<string, unknown>
    const find = sel.Find as Record<string, unknown>
    const executeSpy = find.Execute as ReturnType<typeof vi.fn>

    executeSpy.mockReturnValue(true)
    const result = await editor.findText("keyword", { matchCase: true, matchWholeWord: true, direction: "forward", wrap: false })

    expect(executeSpy).toHaveBeenCalledWith(
      "keyword", true, true, false, false, false, true, 0, false, "", 0
    )
  })

  it("findText returns empty string when not found", async () => {
    const session = createMockSession()
    const editor = new WordTextEditor(session)
    const sel = (session.application as Record<string, unknown>).Selection as Record<string, unknown>
    const find = sel.Find as Record<string, unknown>
    const executeSpy = find.Execute as ReturnType<typeof vi.fn>

    Object.defineProperty(sel, "Type", { value: 1, configurable: true })
    executeSpy.mockReturnValue(false)
    const result = await editor.findText("nonexistent")
    expect(result).toBe("")
  })
})
