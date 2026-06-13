import { describe, it, expect, vi } from "vitest"
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

describe("WordTextEditor COM operations", () => {
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
