import { describe, it, expect, vi } from "vitest"
import { WordContentWriter } from "../../../src/word/word-content-writer.js"
import type { IChartDataBridge } from "../../../src/word/chart-data-bridge.js"
import { createMockSession, createMockDoc, createMockSel } from "../test-helpers.js"
import { MockRangeProxy, MockDocumentProxy, MockSelectionProxy } from "../../../src/word/com-proxy/com-proxy.mock.js"
import type { IWordSession } from "../../../src/word/session.js"

function makeWriter(options?: {
  customDoc?: Record<string, unknown>
  customSel?: Record<string, unknown>
  sessionOverrides?: Partial<Record<string, unknown>>
  chartBridgeOverrides?: Partial<IChartDataBridge>
}) {
  const chartBridge: IChartDataBridge = {
    setChartData: vi.fn().mockResolvedValue({ ok: true, series: 3 }),
    dispose: vi.fn(),
    ...options?.chartBridgeOverrides,
  }
  const session = createMockSession(
    options?.sessionOverrides as unknown as Partial<Record<string, unknown>> | undefined,
    options?.customDoc,
    options?.customSel,
  )
  const writer = new WordContentWriter(session, chartBridge)
  const selProxy = session.getSelectionProxy() as MockSelectionProxy
  const docProxy = session.getDocProxy() as MockDocumentProxy
  const rawSel = session.application.Selection as Record<string, unknown>
  const rawDoc = session.application.ActiveDocument as Record<string, unknown>
  return { writer, session, chartBridge, selProxy, docProxy, rawSel, rawDoc }
}

// ---------------------------------------------------------------------------
// writeBlocks
// ---------------------------------------------------------------------------
describe("writeBlocks", () => {
  it("returns {blocks:0, chars:0} for empty string", async () => {
    const { writer } = makeWriter()
    expect(await writer.writeBlocks("")).toEqual({ blocks: 0, chars: 0 })
  })

  it("returns {blocks:0, chars:0} for whitespace", async () => {
    const { writer } = makeWriter()
    expect(await writer.writeBlocks("   \n\n  ")).toEqual({ blocks: 0, chars: 0 })
  })

  it("writes a heading block with style and trailing typeParagraph", async () => {
    const sel = createMockSel()
    const { writer, selProxy, rawSel } = makeWriter({ customSel: sel })
    const result = await writer.writeBlocks("## Section Title")
    expect(selProxy.typeText).toHaveBeenCalledWith("Section Title")
    expect(selProxy.moveStart).toHaveBeenCalledWith(1, -13)
    expect((rawSel.Style as string)).toContain("Heading 2")
    expect(selProxy.collapse).toHaveBeenCalledWith(0)
    expect(result).toEqual({ blocks: 1, chars: 13 })
  })

  it("writes a paragraph block and returns correct counts", async () => {
    const sel = createMockSel()
    const { writer, selProxy } = makeWriter({ customSel: sel })
    const result = await writer.writeBlocks("Hello paragraph")
    expect(selProxy.typeText).toHaveBeenCalledWith("Hello paragraph")
    expect(result).toEqual({ blocks: 1, chars: 15 })
  })

  it("writes multiple blocks and returns correct counts", async () => {
    const sel = createMockSel()
    const { writer, selProxy } = makeWriter({ customSel: sel })
    const result = await writer.writeBlocks("## A\n\nB\n\n### C")
    expect(selProxy.typeText).toHaveBeenCalledWith("A")
    expect(selProxy.typeText).toHaveBeenCalledWith("B")
    expect(selProxy.typeText).toHaveBeenCalledWith("C")
    expect(result.blocks).toBe(3)
  })
})

// ---------------------------------------------------------------------------
// insertAtCursor
// ---------------------------------------------------------------------------
describe("insertAtCursor", () => {
  it("returns {blocks:0, chars:0} for empty input", async () => {
    const { writer } = makeWriter()
    expect(await writer.insertAtCursor("")).toEqual({ blocks: 0, chars: 0 })
  })

  it("inserts blocks at cursor without goToEnd", async () => {
    const sel = createMockSel()
    const { writer, selProxy } = makeWriter({ customSel: sel })
    await writer.insertAtCursor("Hello cursor")
    expect(selProxy.typeText).toHaveBeenCalledWith("Hello cursor")
  })

  it("returns correct block/char counts when inserting at cursor", async () => {
    const sel = createMockSel()
    const { writer } = makeWriter({ customSel: sel })
    const result = await writer.insertAtCursor("## Title\n\nBody")
    expect(result.blocks).toBeGreaterThan(0)
    expect(result.chars).toBeGreaterThan(0)
  })
})

// ---------------------------------------------------------------------------
// insertParagraph
// ---------------------------------------------------------------------------
describe("insertParagraph", () => {
  it("calls proxy.typeParagraph once by default", async () => {
    const { writer, selProxy } = makeWriter()
    await writer.insertParagraph()
    expect(selProxy.typeParagraph).toHaveBeenCalledTimes(1)
  })

  it("calls proxy.typeParagraph N times for given count", async () => {
    const { writer, selProxy } = makeWriter()
    await writer.insertParagraph(3)
    expect(selProxy.typeParagraph).toHaveBeenCalledTimes(3)
  })
})

// ---------------------------------------------------------------------------
// insertPageBreak
// ---------------------------------------------------------------------------
describe("insertPageBreak", () => {
  it("calls proxy.insertBreak(7)", async () => {
    const { writer, selProxy } = makeWriter()
    await writer.insertPageBreak()
    expect(selProxy.insertBreak).toHaveBeenCalledWith(7)
  })
})

// ---------------------------------------------------------------------------
// insertHorizontalLine
// ---------------------------------------------------------------------------
describe("insertHorizontalLine", () => {
  it("calls addHorizontalLine on selection proxy", async () => {
    const { writer, selProxy } = makeWriter()
    await writer.insertHorizontalLine()
    expect(selProxy.addHorizontalLine).toHaveBeenCalled()
  })
})

// ---------------------------------------------------------------------------
// insertList
// ---------------------------------------------------------------------------
describe("insertList", () => {
  function listSel() {
    const sel = createMockSel()
    const lf = (sel.Range as Record<string, unknown>).ListFormat as Record<string, ReturnType<typeof vi.fn>>
    lf.ApplyBulletDefault = vi.fn()
    lf.ApplyNumberDefault = vi.fn()
    lf.RemoveNumbers = vi.fn()
    return { sel, lf }
  }

  it("applies bullet list, types items, removes numbers", async () => {
    const { sel, lf } = listSel()
    const { writer, selProxy } = makeWriter({ customSel: sel })
    await writer.insertList("bullet", ["alpha", "beta"])
    expect(lf.ApplyBulletDefault).toHaveBeenCalled()
    expect(selProxy.typeText).toHaveBeenCalledWith("alpha")
    expect(selProxy.typeText).toHaveBeenCalledWith("beta")
    expect(lf.RemoveNumbers).toHaveBeenCalled()
  })

  it("applies number list", async () => {
    const { sel, lf } = listSel()
    const { writer } = makeWriter({ customSel: sel })
    await writer.insertList("number", ["one", "two"])
    expect(lf.ApplyNumberDefault).toHaveBeenCalled()
    expect(lf.RemoveNumbers).toHaveBeenCalled()
  })

  it("handles single item list", async () => {
    const { sel, lf } = listSel()
    const { writer, selProxy } = makeWriter({ customSel: sel })
    await writer.insertList("bullet", ["only"])
    expect(selProxy.typeText).toHaveBeenCalledWith("only")
    expect(lf.RemoveNumbers).toHaveBeenCalled()
  })
})

// ---------------------------------------------------------------------------
// addHyperlink
// ---------------------------------------------------------------------------
describe("addHyperlink", () => {
  function linkWriter(sel: Record<string, unknown>) {
    const hyperlinksAdd = vi.fn()
    const session = createMockSession(undefined, undefined, sel)
    const docProxy = session.getDocProxy() as MockDocumentProxy
    docProxy.getHyperlinks = vi.fn(() => ({ add: hyperlinksAdd } as unknown as Record<string, unknown>))
    const chartBridge: IChartDataBridge = { setChartData: vi.fn().mockResolvedValue({ ok: true, series: 3 }), dispose: vi.fn() }
    const writer = new WordContentWriter(session, chartBridge)
    const selProxy = session.getSelectionProxy() as MockSelectionProxy
    return { writer, selProxy, hyperlinksAdd }
  }

  it("adds hyperlink when selection exists (type !== 1)", async () => {
    const sel = createMockSel()
    sel.Type = 2
    const { writer, selProxy, hyperlinksAdd } = linkWriter(sel)
    await writer.addHyperlink("click here", "https://example.com")
    expect(hyperlinksAdd).toHaveBeenCalled()
    // goToEnd() uses doc.getRange(end,end).select() instead of unreliable endKey(wdStory)
    expect(selProxy.collapse).toHaveBeenCalledWith(0)
  })

  it("adds hyperlink when no selection via collapsed range", async () => {
    const sel = createMockSel()
    sel.Type = 1
    const { writer, hyperlinksAdd, selProxy } = linkWriter(sel)
    await writer.addHyperlink("link", "https://example.com")
    expect(hyperlinksAdd).toHaveBeenCalled()
    // goToEnd() uses doc.getRange(end,end).select() instead of unreliable endKey(wdStory)
    expect(selProxy.collapse).toHaveBeenCalledWith(0)
  })

  it("passes subAddress and screenTip to hyperlinks.add", async () => {
    const sel = createMockSel()
    sel.Type = 2
    const { writer, hyperlinksAdd } = linkWriter(sel)
    await writer.addHyperlink("link", "https://example.com", "sub", "tooltip")
    expect(hyperlinksAdd).toHaveBeenCalledWith(
      expect.anything(),
      "https://example.com",
      "sub",
      "tooltip",
      "link",
    )
  })
})

// ---------------------------------------------------------------------------
// addFootnote
// ---------------------------------------------------------------------------
describe("addFootnote", () => {
  it("adds footnote via Footnotes.Add on success", async () => {
    const sel = createMockSel()
    const session = createMockSession(undefined, undefined, sel)
    const docProxy = session.getDocProxy() as MockDocumentProxy
    const footnotesAdd = vi.fn()
    docProxy.getFootnotes = vi.fn().mockReturnValue({ add: footnotesAdd } as unknown as Record<string, unknown>)
    const chartBridge: IChartDataBridge = { setChartData: vi.fn().mockResolvedValue({ ok: true, series: 3 }), dispose: vi.fn() }
    const writer = new WordContentWriter(session, chartBridge)
    await writer.addFootnote("source note")
    expect(footnotesAdd).toHaveBeenCalled()
  })

  it("falls back to text when Footnotes.Add throws", async () => {
    const sel = createMockSel()
    const session = createMockSession(undefined, undefined, sel)
    const docProxy = session.getDocProxy() as MockDocumentProxy
    const footnotesAdd = vi.fn()
    docProxy.getFootnotes = vi.fn().mockReturnValue({ add: footnotesAdd } as unknown as Record<string, unknown>)
    footnotesAdd.mockImplementationOnce(() => { throw new Error("COM fail") })
    const chartBridge: IChartDataBridge = { setChartData: vi.fn().mockResolvedValue({ ok: true, series: 3 }), dispose: vi.fn() }
    const writer = new WordContentWriter(session, chartBridge)
    const selProxy = session.getSelectionProxy() as MockSelectionProxy
    await writer.addFootnote("source note")
    expect(selProxy.typeText).toHaveBeenCalled()
  })

  it("does not throw when session.withScreenOff fails", async () => {
    const { writer } = makeWriter({
      sessionOverrides: {
        withScreenOff: vi.fn(() => Promise.reject(new Error("session error"))),
      },
    })
    await expect(writer.addFootnote("note")).resolves.toBeUndefined()
  })
})

// ---------------------------------------------------------------------------
// insertFile
// ---------------------------------------------------------------------------
describe("insertFile", () => {
  it("calls range.insertFile with the path", async () => {
    const sel = createMockSel()
    const rangeMock = new MockRangeProxy(sel.Range as Record<string, unknown>)
    rangeMock.insertFile = vi.fn()
    const { writer, selProxy } = makeWriter({ customSel: sel })
    selProxy.getRange = vi.fn(() => rangeMock)
    await writer.insertFile("C:\\template.docx")
    expect(rangeMock.insertFile).toHaveBeenCalledWith("C:\\template.docx")
  })
})

// ---------------------------------------------------------------------------
// insertSectionBreak
// ---------------------------------------------------------------------------
describe("insertSectionBreak", () => {
  it("inserts nextPage section break by default (type 8)", async () => {
    const { writer, selProxy } = makeWriter()
    await writer.insertSectionBreak()
    expect(selProxy.collapse).toHaveBeenCalledWith(1)
    expect(selProxy.insertBreak).toHaveBeenCalledWith(8)
  })

  it("inserts continuous section break (type 9)", async () => {
    const { writer, selProxy } = makeWriter()
    await writer.insertSectionBreak("continuous")
    expect(selProxy.insertBreak).toHaveBeenCalledWith(9)
  })
})

// ---------------------------------------------------------------------------
// setColumns
// ---------------------------------------------------------------------------
describe("setColumns", () => {
  it("calls SetCount on TextColumns", async () => {
    const doc = createMockDoc()
    const setCount = vi.fn()
    const textColumns = { SetCount: setCount }
    doc.Sections = {
      Count: 1,
      Item: vi.fn(() => ({ PageSetup: { TextColumns: textColumns } })),
    }
    const { writer } = makeWriter({ customDoc: doc })
    await writer.setColumns(2)
    expect(setCount).toHaveBeenCalledWith(2)
  })

  it("sets spacing when provided", async () => {
    const doc = createMockDoc()
    const textColumns = { SetCount: vi.fn(), Spacing: 0 }
    doc.Sections = {
      Count: 1,
      Item: vi.fn(() => ({ PageSetup: { TextColumns: textColumns } })),
    }
    const { writer } = makeWriter({ customDoc: doc })
    await writer.setColumns(3, 0.5)
    expect(textColumns.SetCount).toHaveBeenCalledWith(3)
    // 0.5 cm ≈ 14 points
    expect(textColumns.Spacing).toBeGreaterThan(13)
    expect(textColumns.Spacing).toBeLessThan(15)
  })
})

// ---------------------------------------------------------------------------
// insertImage
// ---------------------------------------------------------------------------
describe("insertImage", () => {
  function imageWriter(shape: Record<string, unknown>, customSel?: Record<string, unknown>) {
    const sel = customSel ?? createMockSel()
    const session = createMockSession(undefined, undefined, sel)
    const docProxy = session.getDocProxy() as MockDocumentProxy
    const addPictureMock = vi.fn(() => shape)
    docProxy.getInlineShapes = vi.fn(() => ({ addPicture: addPictureMock } as unknown as Record<string, unknown>))
    const chartBridge: IChartDataBridge = { setChartData: vi.fn().mockResolvedValue({ ok: true, series: 3 }), dispose: vi.fn() }
    const writer = new WordContentWriter(session, chartBridge)
    const selProxy = session.getSelectionProxy() as MockSelectionProxy
    const rawSel = session.application.Selection as Record<string, unknown>
    return { writer, selProxy, rawSel, addPictureMock }
  }

  it("inserts image, sets width/height, and adds paragraph separators", async () => {
    const shape = { Width: 800, Height: 600 }
    const { writer, selProxy, addPictureMock } = imageWriter(shape)
    await writer.insertImage({ imagePath: "C:\\pic.png", width: 200, height: 150 })
    expect(selProxy.typeParagraph).toHaveBeenCalled()
    expect(addPictureMock).toHaveBeenCalledWith("C:\\pic.png")
    expect(shape.Width).toBe(200)
    expect(shape.Height).toBe(150)
    // goToEnd() uses doc.getRange(end,end).select() instead of unreliable endKey(wdStory)
    expect(selProxy.collapse).toHaveBeenCalledWith(0)
    expect(selProxy.typeParagraph).toHaveBeenCalled()
  })

  it("auto-resizes image when width/height not given and exceeds maxWidth", async () => {
    const shape = { Width: 920, Height: 600 }
    const { writer } = imageWriter(shape)
    await writer.insertImage({ imagePath: "C:\\large.png" })
    const ratio = 460 / 920
    expect(shape.Width).toBe(460)
    expect(shape.Height).toBeCloseTo(600 * ratio, 1)
  })

  it("does not auto-resize when image fits within maxWidth", async () => {
    const shape = { Width: 400, Height: 300 }
    const { writer } = imageWriter(shape)
    await writer.insertImage({ imagePath: "C:\\small.png" })
    expect(shape.Width).toBe(400)
    expect(shape.Height).toBe(300)
  })
})

// ---------------------------------------------------------------------------
// insertChart
// ---------------------------------------------------------------------------
describe("insertChart", () => {
  function chartWriter() {
    const sel = createMockSel()
    const session = createMockSession(undefined, undefined, sel)
    const docProxy = session.getDocProxy() as MockDocumentProxy
    const addChart2Result: Record<string, unknown> = {
      Width: 0,
      Height: 0,
      Chart: {
        HasTitle: false,
        ChartTitle: { Text: "" },
        ChartData: { Workbook: { Close: vi.fn() } },
      },
    }
    const inlineShapesMock: Record<string, unknown> = {
      count: 2,
      addChart2: vi.fn(() => addChart2Result),
      item: vi.fn((i: number) => ({ Chart: addChart2Result.Chart as Record<string, unknown> })),
      addPicture: vi.fn(),
    }
    docProxy.getInlineShapes = vi.fn(() => inlineShapesMock as unknown as Record<string, unknown>)
    const chartBridge: IChartDataBridge = {
      setChartData: vi.fn().mockResolvedValue({ ok: true, series: 3 }),
      dispose: vi.fn(),
    }
    const writer = new WordContentWriter(session, chartBridge)
    const selProxy = session.getSelectionProxy() as MockSelectionProxy
    return { writer, selProxy, chartBridge, addChart2Result, inlineShapesMock }
  }

  it("inserts chart, calls chartBridge, closes workbook, adds separators", async () => {
    const { writer, selProxy, chartBridge, addChart2Result } = chartWriter()
    const result = await writer.insertChart({
      type: "bar",
      data: [["Category", "Value"], ["A", 1]],
      title: "My Chart",
      width: 400,
      height: 300,
    })
    expect(addChart2Result.Width).toBe(400)
    expect(addChart2Result.Height).toBe(300)
    expect((addChart2Result as Record<string, unknown>).Chart).toBeDefined()
    expect(chartBridge.setChartData).toHaveBeenCalledWith("test.docx", 3, [["Category", "Value"], ["A", 1]])
    const wb = ((addChart2Result.Chart as Record<string, unknown>).ChartData as Record<string, unknown>).Workbook as Record<string, unknown>
    expect(wb.Close).toHaveBeenCalledWith(false)
    // goToEnd() uses doc.getRange(end,end).select() instead of unreliable endKey(wdStory)
    expect(selProxy.collapse).toHaveBeenCalledWith(0)
    expect(selProxy.typeParagraph).toHaveBeenCalled()
    expect(result).toEqual({ type: "bar", series: 3 })
  })

  it("handles line chart type mapping", async () => {
    const { writer, inlineShapesMock } = chartWriter()
    await writer.insertChart({ type: "line", data: [["x", "y"], [1, 2]] })
    const addChart2Mock = (inlineShapesMock.addChart2 as ReturnType<typeof vi.fn>)
    expect(addChart2Mock).toHaveBeenCalled()
    expect(addChart2Mock.mock.calls[0][1]).toBe(4)
  })

  it("handles no workbook gracefully", async () => {
    const sel = createMockSel()
    const session = createMockSession(undefined, undefined, sel)
    const docProxy = session.getDocProxy() as MockDocumentProxy
    const inlineShapesMock: Record<string, unknown> = {
      count: 2,
      addChart2: vi.fn(() => ({ Width: 0, Height: 0, Chart: {} } as Record<string, unknown>)),
      item: vi.fn(() => ({ Chart: {} })),
      addPicture: vi.fn(),
    }
    docProxy.getInlineShapes = vi.fn(() => inlineShapesMock as unknown as Record<string, unknown>)
    const chartBridge: IChartDataBridge = {
      setChartData: vi.fn().mockResolvedValue({ ok: true, series: 3 }),
      dispose: vi.fn(),
    }
    const writer = new WordContentWriter(session, chartBridge)
    await expect(writer.insertChart({ type: "column", data: [["A", 1]] })).resolves.toBeDefined()
  })
})

// ---------------------------------------------------------------------------
// insertTextbox
// ---------------------------------------------------------------------------
describe("insertTextbox", () => {
  function textboxWriter(addTextboxResult?: Record<string, unknown>) {
    const sel = createMockSel()
    const session = createMockSession(undefined, undefined, sel)
    const docProxy = session.getDocProxy() as MockDocumentProxy
    const result = addTextboxResult ?? { TextFrame: { TextRange: { Text: "" } } }
    const addTextboxMock = vi.fn(() => result)
    docProxy.getShapes = vi.fn().mockReturnValue({ AddTextbox: addTextboxMock } as unknown as Record<string, unknown>)
    const chartBridge: IChartDataBridge = { setChartData: vi.fn().mockResolvedValue({ ok: true, series: 3 }), dispose: vi.fn() }
    const writer = new WordContentWriter(session, chartBridge)
    const selProxy = session.getSelectionProxy() as MockSelectionProxy
    return { writer, selProxy, docProxy, addTextboxMock, addTextboxResult: result }
  }

  it("inserts textbox with default values", async () => {
    const addTextboxResult: Record<string, unknown> = {
      TextFrame: { TextRange: { Text: "" } },
    }
    const { writer, selProxy, addTextboxMock } = textboxWriter(addTextboxResult)
    const result = await writer.insertTextbox({ text: "Hello" })
    expect(addTextboxMock).toHaveBeenCalledWith(1, 50, 50, 200, 100)
    expect((addTextboxResult.TextFrame as Record<string, unknown>).TextRange as Record<string, unknown>).toHaveProperty("Text", "Hello")
    // goToEnd() uses doc.getRange(end,end).select() instead of unreliable endKey(wdStory)
    expect(selProxy.collapse).toHaveBeenCalledWith(0)
    expect(selProxy.typeParagraph).toHaveBeenCalled()
    expect(result).toEqual({ width: 200, height: 100 })
  })

  it("uses custom position, orientation, and dimensions", async () => {
    const { writer, addTextboxMock } = textboxWriter()
    await writer.insertTextbox({
      text: "V", orientation: "vertical", positionLeft: 100,
      positionTop: 200, width: 150, height: 300,
    })
    expect(addTextboxMock).toHaveBeenCalledWith(5, 100, 200, 150, 300)
  })
})

// ---------------------------------------------------------------------------
// insertTable
// ---------------------------------------------------------------------------
describe("insertTable", () => {
  function tableWriter(data?: { withStyle?: boolean; cellFailOn?: number; styleName?: string }) {
    const sel = createMockSel()
    const session = createMockSession(undefined, undefined, sel)
    const docProxy = session.getDocProxy() as MockDocumentProxy
    let cellCallCount = 0
    const tableResult: Record<string, unknown> = {
      Cell: vi.fn(() => {
        cellCallCount++
        if (data?.cellFailOn != null && cellCallCount === data.cellFailOn) {
          throw new Error("cell error")
        }
        return { Range: { Text: "" } }
      }),
      Rows: { Count: 2, Item: vi.fn(() => ({ Range: { Font: { Bold: false } }, Shading: { BackgroundPatternColor: 0 } })) },
      Columns: { Count: 2 },
      Style: "",
      AutoFitBehavior: vi.fn(),
      Borders: {
        Item: vi.fn(() => ({ LineStyle: 0, ColorIndex: 0, LineWidth: 0 })),
      },
    }
    const tablesMock: Record<string, unknown> = {
      count: 0,
      add: vi.fn(() => tableResult),
      item: vi.fn(),
    }
    const stylesMock: Record<string, unknown> = {
      count: data?.withStyle ? 1 : 0,
      item: vi.fn(() => ({ NameLocal: data?.styleName ?? "Grid Table 4 - Accent 1", InUse: false, Type: 0, BuiltIn: false })),
      itemByName: vi.fn(),
    }
    docProxy.getTables = vi.fn(() => tablesMock as unknown as Record<string, unknown>)
    docProxy.getStyles = vi.fn(() => stylesMock as unknown as Record<string, unknown>)
    const chartBridge: IChartDataBridge = { setChartData: vi.fn().mockResolvedValue({ ok: true, series: 3 }), dispose: vi.fn() }
    const writer = new WordContentWriter(session, chartBridge)
    const selProxy = session.getSelectionProxy() as MockSelectionProxy
    return { writer, selProxy, docProxy, tableResult, tablesMock }
  }

  it("creates empty table without data", async () => {
    const { writer, selProxy, tablesMock } = tableWriter()
    const result = await writer.insertTable({ rows: 3, columns: 2 })
    expect(tablesMock.add).toHaveBeenCalled()
    // insertTable now uses doc.getRange(tableEnd,docEnd).select() instead of unreliable endKey(wdStory)
    expect(selProxy.collapse).toHaveBeenCalledWith(0)
    expect(selProxy.typeParagraph).toHaveBeenCalled()
    expect(result).toEqual({ rows: 3, columns: 2, failCount: 0 })
  })

  it("fills cell data and counts failures", async () => {
    const { writer } = tableWriter({ cellFailOn: 2 })
    const result = await writer.insertTable({
      rows: 2, columns: 2,
      data: [["A", "B"], ["C", "D"]],
    })
    expect(result.failCount).toBe(1)
  })

  it("applies built-in style when available", async () => {
    const { writer, tableResult } = tableWriter({ withStyle: true })
    await writer.insertTable({ rows: 2, columns: 2, data: [["H1", "H2"], ["D1", "D2"]] })
    expect(tableResult.Style).toBe("Grid Table 4 - Accent 1")
  })

  it("falls back to manual header formatting when no style match", async () => {
    const sel = createMockSel()
    const session = createMockSession(undefined, undefined, sel)
    const docProxy = session.getDocProxy() as MockDocumentProxy
    const firstRowFont = { Bold: false }
    const firstRowShading = { BackgroundPatternColor: 0 }
    const tableResult: Record<string, unknown> = {
      Cell: vi.fn(() => ({ Range: { Text: "" } })),
      Rows: {
        Count: 2,
        Item: vi.fn(() => ({ Range: { Font: firstRowFont }, Shading: firstRowShading })),
      },
      Columns: { Count: 2 },
      Style: "",
      AutoFitBehavior: vi.fn(),
      Borders: {
        Item: vi.fn(() => ({ LineStyle: 0, ColorIndex: 0, LineWidth: 0 })),
      },
    }
    const tablesMock: Record<string, unknown> = {
      count: 0,
      add: vi.fn(() => tableResult),
      item: vi.fn(),
    }
    const stylesMock: Record<string, unknown> = {
      count: 0,
      item: vi.fn(),
      itemByName: vi.fn(),
    }
    docProxy.getTables = vi.fn(() => tablesMock as unknown as Record<string, unknown>)
    docProxy.getStyles = vi.fn(() => stylesMock as unknown as Record<string, unknown>)
    const chartBridge: IChartDataBridge = { setChartData: vi.fn().mockResolvedValue({ ok: true, series: 3 }), dispose: vi.fn() }
    const writer = new WordContentWriter(session, chartBridge)
    await writer.insertTable({ rows: 2, columns: 2, data: [["H1", "H2"], ["D1", "D2"]] })
    expect(firstRowFont.Bold).toBe(true)
    expect(firstRowShading.BackgroundPatternColor).toBe(0xF3E2D9)
  })

  it("applies autoFitBehavior", async () => {
    const sel = createMockSel()
    const session = createMockSession(undefined, undefined, sel)
    const docProxy = session.getDocProxy() as MockDocumentProxy
    const autoFitFn = vi.fn()
    const tableResult: Record<string, unknown> = {
      Cell: vi.fn(),
      Rows: { Count: 1, Item: vi.fn() },
      Columns: { Count: 1 },
      Style: "",
      AutoFitBehavior: autoFitFn,
      Borders: {
        Item: vi.fn(() => ({ LineStyle: 0, ColorIndex: 0, LineWidth: 0 })),
      },
    }
    const tablesMock: Record<string, unknown> = {
      count: 0,
      add: vi.fn(() => tableResult),
      item: vi.fn(),
    }
    docProxy.getTables = vi.fn(() => tablesMock as unknown as Record<string, unknown>)
    docProxy.getStyles = vi.fn(() => ({ count: 0, item: vi.fn(), itemByName: vi.fn() }) as unknown as Record<string, unknown>)
    const chartBridge: IChartDataBridge = { setChartData: vi.fn().mockResolvedValue({ ok: true, series: 3 }), dispose: vi.fn() }
    const writer = new WordContentWriter(session, chartBridge)
    await writer.insertTable({ rows: 1, columns: 1, autoFitBehavior: "window" })
    expect(autoFitFn).toHaveBeenCalledWith(2)
  })
})

// ---------------------------------------------------------------------------
// textToTable
// ---------------------------------------------------------------------------
describe("textToTable", () => {
  function makeTableWriter(convertResult: Record<string, unknown>) {
    const sel = createMockSel()
    const rangeMock = new MockRangeProxy(sel.Range as Record<string, unknown>)
    rangeMock.convertToTable = vi.fn(() => convertResult)
    const session = createMockSession(undefined, undefined, sel)
    const selProxy = session.getSelectionProxy() as MockSelectionProxy
    selProxy.getRange = vi.fn(() => rangeMock)
    const chartBridge: IChartDataBridge = { setChartData: vi.fn().mockResolvedValue({ ok: true, series: 3 }), dispose: vi.fn() }
    const writer = new WordContentWriter(session, chartBridge)
    return { writer, rangeMock }
  }

  it("converts selection range to table with default separator", async () => {
    const convertResult = { Rows: { Count: 3 }, Columns: { Count: 2 } }
    const { writer, rangeMock } = makeTableWriter(convertResult)
    const result = await writer.textToTable({})
    expect(rangeMock.convertToTable).toHaveBeenCalledWith("\t")
    expect(result).toEqual({ rows: 3, columns: 2 })
  })

  it("uses custom separator and autoFitBehavior", async () => {
    const convertResult = { Rows: { Count: 2 }, Columns: { Count: 3 }, AutoFitBehavior: 0 }
    const { writer, rangeMock } = makeTableWriter(convertResult)
    const result = await writer.textToTable({ separator: ",", autoFitBehavior: "fixed" })
    expect(rangeMock.convertToTable).toHaveBeenCalledWith(",")
    expect((convertResult as Record<string, unknown>).AutoFitBehavior).toBe(0)
    expect(result).toEqual({ rows: 2, columns: 3 })
  })
})

// ---------------------------------------------------------------------------
// replaceVariables
// ---------------------------------------------------------------------------
describe("replaceVariables", () => {
  function findDoc() {
    const doc = createMockDoc()
    doc.Content.End = 100
    const executeMock = vi.fn().mockReturnValue(false)
    const findMock = {
      ClearFormatting: vi.fn(),
      Text: "",
      Forward: true,
      Wrap: 0,
      Format: false,
      MatchCase: false,
      MatchWholeWord: false,
      MatchWildcards: false,
      Execute: executeMock,
      Replacement: {
        ClearFormatting: vi.fn(),
        Text: "",
      },
    }
    doc.Range = vi.fn(() => ({
      End: 100,
      Start: 0,
      Select: vi.fn(),
      Hyperlinks: { Add: vi.fn() },
      Shading: { BackgroundPatternColor: 0 },
      Find: findMock,
    }))
    return { doc, findMock, executeMock }
  }

  it("replaces single variable found once", async () => {
    const { doc, executeMock } = findDoc()
    executeMock.mockReturnValueOnce(true).mockReturnValue(false)
    const sel = createMockSel()
    sel.Start = 0
    sel.End = 50
    const { writer } = makeWriter({ customDoc: doc, customSel: sel })
    const results = await writer.replaceVariables({ name: "Alice" })
    expect(results).toEqual([{ key: "name", count: 1 }])
  })

  it("returns zero count for missing variable", async () => {
    const { doc } = findDoc()
    const sel = createMockSel()
    const { writer } = makeWriter({ customDoc: doc, customSel: sel })
    const results = await writer.replaceVariables({ missing: "value" })
    expect(results).toEqual([{ key: "missing", count: 0 }])
  })

  it("replaces multiple variables", async () => {
    const { doc } = findDoc()
    const sel = createMockSel()
    const { writer } = makeWriter({ customDoc: doc, customSel: sel })
    const results = await writer.replaceVariables({ a: "1", b: "2" })
    expect(results[0]).toEqual({ key: "a", count: 0 })
    expect(results[1]).toEqual({ key: "b", count: 0 })
  })

  it("restores selection after replacement", async () => {
    const doc = createMockDoc()
    doc.Content.End = 100
    const executeMock = vi.fn().mockReturnValue(false)
    const findMock = {
      ClearFormatting: vi.fn(),
      Text: "",
      Forward: true,
      Wrap: 0,
      Format: false,
      MatchCase: false,
      MatchWholeWord: false,
      MatchWildcards: false,
      Execute: executeMock,
      Replacement: { ClearFormatting: vi.fn(), Text: "" },
    }
    doc.Range = vi.fn((start?: number, end?: number) => ({
      End: end ?? 100,
      Start: start ?? 0,
      Select: vi.fn(),
      Hyperlinks: { Add: vi.fn() },
      Shading: { BackgroundPatternColor: 0 },
      Find: findMock,
    }))
    const sel = createMockSel()
    sel.Start = 5
    sel.End = 15
    const { writer, rawDoc } = makeWriter({ customDoc: doc, customSel: sel })
    await writer.replaceVariables({ x: "y" })
    // The finally block calls doc.getRange(5, 15) for selection restore
    expect(rawDoc.Range).toHaveBeenCalledWith(5, 15)
  })
})
