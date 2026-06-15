import { describe, it, expect, vi } from "vitest"
import { WordFormatting, type StyleProfile } from "../../../src/word/formatting.js"
import { createMockSession, createMockDoc, createMockSel } from "../test-helpers.js"

function createFormatting() {
  const doc = createMockDoc()
  const sel = createMockSel()
  const session = createMockSession({}, doc, sel)
  const fmt = new WordFormatting(session)
  return { fmt, doc, sel, session }
}

function mockStyle(fontProps: Record<string, unknown> = {}, pfProps: Record<string, unknown> = {}) {
  const borders = {
    Item: vi.fn(() => ({ LineStyle: 0, ColorIndex: 0, LineWidth: 0 })),
  }
  return {
    Font: {
      Name: "", Size: 11, Bold: false, Italic: false,
      ColorIndex: 0, Underline: 0, Strikethrough: 0, HighlightColorIndex: 0,
      Shading: { BackgroundPatternColor: 0 },
      ...fontProps,
    },
    ParagraphFormat: {
      Alignment: 0, LeftIndent: 0, RightIndent: 0,
      FirstLineIndent: 0, SpaceBefore: 0, SpaceAfter: 0,
      LineSpacing: 0, LineSpacingRule: 0,
      KeepWithNext: false, PageBreakBefore: false,
      Borders: borders,
      ...pfProps,
    } as Record<string, unknown>,
  }
}

describe("WordFormatting.modifyStyle", () => {
  it("sets font underline via enum", async () => {
    const { fmt, doc } = createFormatting()
    const st = mockStyle()
    doc.Styles = { Item: vi.fn(() => st), Count: 1 }
    await fmt.modifyStyle("Normal", { font: { underline: "double" } })
    expect(st.Font.Underline).toBe(3)
  })

  it("sets font strikethrough", async () => {
    const { fmt, doc } = createFormatting()
    const st = mockStyle()
    doc.Styles = { Item: vi.fn(() => st), Count: 1 }
    await fmt.modifyStyle("Normal", { font: { strikethrough: true } })
    expect(st.Font.Strikethrough).toBe(1)
  })

  it("sets font highlight via enum name", async () => {
    const { fmt, doc } = createFormatting()
    const st = mockStyle()
    doc.Styles = { Item: vi.fn(() => st), Count: 1 }
    await fmt.modifyStyle("Normal", { font: { highlight: "yellow" } })
    expect(st.Font.HighlightColorIndex).toBe(7)
  })

  it("sets font highlight via hex to shading", async () => {
    const { fmt, doc } = createFormatting()
    const st = mockStyle()
    doc.Styles = { Item: vi.fn(() => st), Count: 1 }
    await fmt.modifyStyle("Normal", { font: { highlight: "#FFF0E0" } })
    expect(st.Font.Shading.BackgroundPatternColor).toBe(14741759)
  })

  it("sets paragraph keepWithNext", async () => {
    const { fmt, doc } = createFormatting()
    const st = mockStyle()
    doc.Styles = { Item: vi.fn(() => st), Count: 1 }
    await fmt.modifyStyle("Normal", { paragraph: { keepWithNext: true } })
    expect((st.ParagraphFormat as Record<string, unknown>).KeepWithNext).toBe(true)
  })

  it("sets paragraph pageBreakBefore", async () => {
    const { fmt, doc } = createFormatting()
    const st = mockStyle()
    doc.Styles = { Item: vi.fn(() => st), Count: 1 }
    await fmt.modifyStyle("Heading 1", { paragraph: { pageBreakBefore: true } })
    expect((st.ParagraphFormat as Record<string, unknown>).PageBreakBefore).toBe(true)
  })

  it("applies borders in correct order: LineStyle before ColorIndex", async () => {
    const { fmt, doc } = createFormatting()
    const borderMock = { LineStyle: 0, ColorIndex: 0, LineWidth: 0 }
    const borders = { Item: vi.fn(() => borderMock) }
    const pf = { Borders: borders, KeepWithNext: false, PageBreakBefore: false, Alignment: 0 } as Record<string, unknown>
    const st = mockStyle({}, pf)
    doc.Styles = { Item: vi.fn(() => st), Count: 1 }

    await fmt.modifyStyle("Normal", {
      paragraph: { borders: { style: "single", color: "red", size: 8, sides: ["top"] } },
    })

    expect(borders.Item).toHaveBeenCalledWith(-1) // top
    expect(borderMock.LineStyle).toBe(1) // single
    expect(borderMock.ColorIndex).toBe(6) // red
    expect(borderMock.LineWidth).toBe(8)
  })

  it("applies borders to all four sides by default", async () => {
    const { fmt, doc } = createFormatting()
    const borderMock = { LineStyle: 0, ColorIndex: 0, LineWidth: 0 }
    const borders = { Item: vi.fn(() => borderMock) }
    const pf = { Borders: borders, KeepWithNext: false, PageBreakBefore: false, Alignment: 0 } as Record<string, unknown>
    const st = mockStyle({}, pf)
    doc.Styles = { Item: vi.fn(() => st), Count: 1 }

    await fmt.modifyStyle("Normal", {
      paragraph: { borders: { style: "single" } },
    })

    expect(borders.Item).toHaveBeenCalledWith(-1) // top
    expect(borders.Item).toHaveBeenCalledWith(-2) // left
    expect(borders.Item).toHaveBeenCalledWith(-3) // bottom
    expect(borders.Item).toHaveBeenCalledWith(-4) // right
  })

  it("handles borders gracefully when per-side COM fails", async () => {
    const { fmt, doc } = createFormatting()
    const borders = { Item: vi.fn(() => { throw new Error("COM fail") }) }
    const pf = { Borders: borders, KeepWithNext: false, PageBreakBefore: false, Alignment: 0 } as Record<string, unknown>
    const st = mockStyle({}, pf)
    doc.Styles = { Item: vi.fn(() => st), Count: 1 }

    await expect(fmt.modifyStyle("Normal", {
      paragraph: { borders: { style: "single" } },
    })).resolves.toBeUndefined()
  })

  it("existing font fields still work alongside new fields", async () => {
    const { fmt, doc } = createFormatting()
    const st = mockStyle()
    doc.Styles = { Item: vi.fn(() => st), Count: 1 }
    await fmt.modifyStyle("Normal", {
      font: { name: "Calibri", size: 14, bold: true, underline: "wavy" },
    })
    expect(st.Font.Name).toBe("Calibri")
    expect(st.Font.Size).toBe(14)
    expect(st.Font.Bold).toBe(true)
    expect(st.Font.Underline).toBe(11)
  })

  it("existing paragraph fields still work alongside new fields", async () => {
    const { fmt, doc } = createFormatting()
    const st = mockStyle()
    doc.Styles = { Item: vi.fn(() => st), Count: 1 }
    await fmt.modifyStyle("Normal", {
      paragraph: { alignment: "center", spaceAfter: 12, keepWithNext: true },
    })
    expect((st.ParagraphFormat as Record<string, unknown>).Alignment).toBe(1)
    expect((st.ParagraphFormat as Record<string, unknown>).SpaceAfter).toBe(12)
    expect((st.ParagraphFormat as Record<string, unknown>).KeepWithNext).toBe(true)
  })
})
