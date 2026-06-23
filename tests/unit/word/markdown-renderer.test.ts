import { describe, it, expect, vi } from "vitest"
import { MarkdownRenderer, type RenderContext } from "../../../src/word/markdown-renderer.js"
import { createMockDoc, createMockSel } from "../test-helpers.js"
import { MockSelectionProxy, MockDocumentProxy } from "../../../src/word/com-proxy/com-proxy.mock.js"
import type { Block } from "../../../src/word/markdown-parser.js"

function createRenderContext(
  customSel?: Record<string, unknown>,
  customDoc?: Record<string, unknown>,
): {
  ctx: RenderContext
  sel: MockSelectionProxy
  doc: MockDocumentProxy
  selRaw: Record<string, unknown>
  docRaw: Record<string, unknown>
  mockInlineShapes: { addPicture: ReturnType<typeof vi.fn> }
  mockTables: { add: ReturnType<typeof vi.fn> }
  goToEnd: ReturnType<typeof vi.fn>
} {
  const rawSel = customSel ?? createMockSel()
  const rawDoc = customDoc ?? createMockDoc()
  const selProxy = new MockSelectionProxy(rawSel)
  const docProxy = new MockDocumentProxy(rawDoc)
  const goToEnd = vi.fn()

  // delegate to raw records like real proxy does
  selProxy.getParagraphFormat = vi.fn(() => rawSel.ParagraphFormat as Record<string, unknown>)
  selProxy.getFont = vi.fn(() => rawSel.Font as Record<string, unknown>)

  // cached collections so proxy method assertions work
  const mockInlineShapes = {
    count: 0, item: vi.fn(), add: vi.fn(),
    addPicture: vi.fn((...args: unknown[]) => {
      const rawAdd = (rawDoc.InlineShapes as Record<string, unknown> | undefined)?.AddPicture as ((...a: unknown[]) => unknown) | undefined
      return rawAdd ? rawAdd(...args) : undefined
    }),
    addChart2: vi.fn(), addHorizontalLineStandard: vi.fn(), addTextbox: vi.fn(),
    itemByName: vi.fn(),
  }
  vi.spyOn(docProxy, 'getInlineShapes').mockReturnValue(mockInlineShapes as never)

  const mockTables = {
    count: 0, item: vi.fn(),
    add: vi.fn((...args: unknown[]) => {
      const rawAdd = (rawDoc.Tables as Record<string, unknown> | undefined)?.Add as ((...a: unknown[]) => unknown) | undefined
      return rawAdd ? rawAdd(...args) : {}
    }),
  }
  vi.spyOn(docProxy, 'getTables').mockReturnValue(mockTables as never)

  return {
    ctx: {
      getSelection: () => selProxy,
      requireDoc: () => docProxy,
      goToEnd,
      withScreenOff: async <T>(fn: () => Promise<T>) => fn(),
    },
    sel: selProxy,
    doc: docProxy,
    selRaw: rawSel,
    docRaw: rawDoc,
    mockInlineShapes,
    mockTables,
    goToEnd,
  }
}

describe("MarkdownRenderer.renderBlock", () => {
  it("renders hr via addHorizontalLine", async () => {
    const { ctx, sel } = createRenderContext()
    const r = new MarkdownRenderer(ctx)
    await r.renderBlock({ type: "hr" } as Block, 0, 1, "end")
    expect(sel.addHorizontalLine).toHaveBeenCalled()
    expect(sel.typeParagraph).toHaveBeenCalled()
  })

  it("renders hr at cursor without goToEnd", async () => {
    const { ctx, goToEnd } = createRenderContext()
    const r = new MarkdownRenderer(ctx)
    await r.renderBlock({ type: "hr" } as Block, 0, 1, "cursor")
    expect(goToEnd).not.toHaveBeenCalled()
  })

  it("renders pagebreak via insertBreak(7)", async () => {
    const { ctx, sel } = createRenderContext()
    const r = new MarkdownRenderer(ctx)
    const chars = await r.renderBlock({ type: "pagebreak" } as Block, 0, 1, "end")
    expect(sel.insertBreak).toHaveBeenCalledWith(7)
    expect(chars).toBe(0)
  })

  it("renders image via proxy addPicture", async () => {
    const { ctx, mockInlineShapes } = createRenderContext()
    const r = new MarkdownRenderer(ctx)
    const chars = await r.renderBlock({ type: "image", alt: "pic", url: "C:\\img.png" } as Block, 0, 1, "end")
    expect(mockInlineShapes.addPicture).toHaveBeenCalledWith("C:\\img.png")
    expect(chars).toBe(3 + 10)
  })

  it("renders image with typeParagraph", async () => {
    const { ctx, sel, mockInlineShapes } = createRenderContext()
    const r = new MarkdownRenderer(ctx)
    const chars = await r.renderBlock({ type: "image", alt: "pic", url: "C:\\img.png" } as Block, 0, 1, "end")
    expect(mockInlineShapes.addPicture).toHaveBeenCalledWith("C:\\img.png")
    expect(sel.typeParagraph).toHaveBeenCalled()
    expect(chars).toBe(3 + 10)
  })

  it("renders heading with style and paragraph after", async () => {
    const { ctx, sel, selRaw } = createRenderContext()
    const r = new MarkdownRenderer(ctx)
    const chars = await r.renderBlock({ type: "heading", level: 2, text: "Section" } as Block, 0, 2, "end")
    expect(sel.typeText).toHaveBeenCalledWith("Section")
    expect(selRaw.Style).toContain("Heading 2")
    expect(sel.typeParagraph).toHaveBeenCalled()
    expect(chars).toBe(7)
  })

  it("renders paragraph without heading style", async () => {
    const { ctx, sel, selRaw } = createRenderContext()
    const r = new MarkdownRenderer(ctx)
    await r.renderBlock({ type: "paragraph", text: "Body" } as Block, 0, 2, "end")
    expect(sel.typeText).toHaveBeenCalledWith("Body")
    expect(selRaw.Style).not.toContain("Heading")
    expect(sel.typeParagraph).toHaveBeenCalled()
  })

  it("renders paragraph without trailing paragraph break for last block", async () => {
    const { ctx, sel } = createRenderContext()
    const r = new MarkdownRenderer(ctx)
    await r.renderBlock({ type: "paragraph", text: "Last" } as Block, 0, 1, "end")
    expect(sel.typeParagraph).not.toHaveBeenCalled()
  })

  it("renders heading with inline bold via Range.Font", async () => {
    const doc = createMockDoc()
    const rangeMock = { Font: {}, Shading: { BackgroundPatternColor: 0 }, Hyperlinks: { Add: vi.fn() } }
    const rangeFn = vi.fn(() => rangeMock)
    doc.Range = rangeFn
    const sel = createMockSel()
    sel.Range.Start = 0
    const { ctx } = createRenderContext(sel, doc)
    const r = new MarkdownRenderer(ctx)
    await r.renderBlock({ type: "paragraph", text: "**bold**" } as Block, 0, 1, "end")
    expect(rangeFn).toHaveBeenCalled()
  })

  it("renders heading with link via Hyperlinks.Add", async () => {
    const doc = createMockDoc()
    const rangeMock = { Font: {}, Shading: { BackgroundPatternColor: 0 }, Hyperlinks: { Add: vi.fn() } }
    const rangeFn = vi.fn(() => rangeMock)
    doc.Range = rangeFn
    const sel = createMockSel()
    sel.Range.Start = 0
    const { ctx } = createRenderContext(sel, doc)
    const r = new MarkdownRenderer(ctx)
    await r.renderBlock({ type: "paragraph", text: "a [link](https://x.com)" } as Block, 0, 1, "end")
    expect(rangeMock.Hyperlinks.Add).toHaveBeenCalled()
  })

  it("renders bullet list with ApplyBulletDefault", async () => {
    const { ctx, sel, selRaw } = createRenderContext()
    const r = new MarkdownRenderer(ctx)
    await r.renderBlock({ type: "bullet_list", items: [{ text: "a", indent: 0 }, { text: "b", indent: 0 }] } as Block, 0, 1, "end")
    expect(sel.getRange().getListFormat().ApplyBulletDefault).toHaveBeenCalled()
    expect(selRaw.TypeText).toHaveBeenNthCalledWith(1, "a")
    expect(selRaw.TypeText).toHaveBeenNthCalledWith(2, "b")
  })

  it("renders bullet list with indentation", async () => {
    const { ctx, sel } = createRenderContext()
    const r = new MarkdownRenderer(ctx)
    await r.renderBlock({ type: "bullet_list", items: [{ text: "deep", indent: 2 }] } as Block, 0, 1, "end")
    expect(sel.getRange().getListFormat().IncreaseIndent).toHaveBeenCalledTimes(2)
  })

  it("renders numbered list with ApplyNumberDefault", async () => {
    const { ctx, sel } = createRenderContext()
    const r = new MarkdownRenderer(ctx)
    await r.renderBlock({ type: "numbered_list", items: [{ text: "1", indent: 0 }] } as Block, 0, 1, "end")
    expect(sel.getRange().getListFormat().ApplyNumberDefault).toHaveBeenCalled()
  })

  it("renders table with Tables.Add", async () => {
    const doc = createMockDoc()
    const tableMock = {
      Cell: vi.fn(() => ({ Range: { Text: "" } })),
      Rows: { Item: vi.fn(() => ({ Range: { Font: {} }, Shading: {} })), Alignment: 0 },
      AutoFitBehavior: vi.fn(),
    }
    doc.Tables = { Add: vi.fn(() => tableMock), Count: 1, Item: vi.fn() }
    const { ctx, docRaw } = createRenderContext(createMockSel(), doc)
    const r = new MarkdownRenderer(ctx)
    await r.renderBlock({ type: "table", rows: [["A", "B"], ["1", "2"]] } as Block, 0, 1, "end")
    expect(docRaw.Tables.Add).toHaveBeenCalledWith(expect.anything(), 2, 2)
  })

  it("renders blockquote with indent, italic, and paragraph break", async () => {
    const rawSel = createMockSel()
    rawSel.Range.Borders = { Item: vi.fn(() => ({})) }
    const { ctx, sel, selRaw } = createRenderContext(rawSel)
    const r = new MarkdownRenderer(ctx)
    await r.renderBlock({ type: "blockquote", text: "quote" } as Block, 0, 2, "end")
    expect(selRaw.ParagraphFormat.LeftIndent).toBe(0)
    expect(selRaw.Font.Italic).toBe(false)
    expect(sel.typeParagraph).toHaveBeenCalled()
  })

  it("renders blockquote without trailing paragraph for last block", async () => {
    const { ctx, sel } = createRenderContext()
    const r = new MarkdownRenderer(ctx)
    await r.renderBlock({ type: "blockquote", text: "last" } as Block, 0, 1, "end")
    expect(sel.typeParagraph).not.toHaveBeenCalled()
  })

  it("renders codeblock with Consolas font", async () => {
    const doc = createMockDoc()
    const rangeMock = { Font: {}, Shading: { BackgroundPatternColor: 0 }, Select: vi.fn() }
    doc.Range = vi.fn(() => rangeMock)
    doc.Content.End = 10
    const sel = createMockSel()
    const { ctx, sel: ctxSel } = createRenderContext(sel, doc)
    const r = new MarkdownRenderer(ctx)
    await r.renderBlock({ type: "codeblock", text: "code" } as Block, 0, 1, "end")
    expect(ctxSel.typeText).toHaveBeenCalledWith("code")
  })

  it("returns 0 for unknown block type", async () => {
    const { ctx } = createRenderContext()
    const r = new MarkdownRenderer(ctx)
    const chars = await r.renderBlock({ type: "image", alt: "", url: "" } as Block, 0, 1, "end")
    expect(chars).toBe(0)
  })
})

describe("MarkdownRenderer.blockCost", () => {
  it("returns 1 for heading/paragraph/blockquote/hr/pagebreak", () => {
    for (const t of ["heading", "paragraph", "blockquote", "hr", "pagebreak"] as const) {
      expect(MarkdownRenderer.blockCost({ type: t } as Block)).toBe(1)
    }
  })

  it("returns 2 for image", () => {
    expect(MarkdownRenderer.blockCost({ type: "image" } as Block)).toBe(2)
  })

  it("returns cost based on items count for bullet_list", () => {
    const c = MarkdownRenderer.blockCost({ type: "bullet_list", items: [{ text: "a", indent: 0 }, { text: "b", indent: 1 }] } as Block)
    expect(c).toBeGreaterThan(1)
  })

  it("returns cost based on rows*cols for table", () => {
    const c = MarkdownRenderer.blockCost({ type: "table", rows: [["a", "b"], ["c", "d"]] } as Block)
    expect(c).toBeGreaterThan(1)
  })

  it("returns cost based on lines for codeblock", () => {
    const c = MarkdownRenderer.blockCost({ type: "codeblock", text: "a\nb\nc" } as Block)
    expect(c).toBeGreaterThan(1)
  })

  it("handles empty table gracefully", () => {
    const c = MarkdownRenderer.blockCost({ type: "table", rows: [] } as Block)
    expect(c).toBeGreaterThanOrEqual(1)
  })
})

describe("MarkdownRenderer.applyHeadingStyle", () => {
  it("sets English style name", () => {
    const sel = { Style: "" }
    MarkdownRenderer.applyHeadingStyle(sel, 1)
    expect(sel.Style).toBe("Heading 1")
  })

  it("falls back to Chinese style if English fails", () => {
    const sel: Record<string, unknown> = {}
    let first = true
    Object.defineProperty(sel, "Style", {
      set(v: unknown) {
        if (first) { first = false; throw new Error("not found") }
        Object.defineProperty(sel, "Style", { value: v, configurable: true, writable: true })
      },
      get() { return undefined },
      configurable: true,
    })
    MarkdownRenderer.applyHeadingStyle(sel, 2)
    expect(sel.Style).toBe("标题 2")
  })
})

describe("MarkdownRenderer.typeSeg", () => {
  it("sets font bold", () => {
    const sel = createMockSel()
    MarkdownRenderer.typeSeg(sel, { text: "b", bold: true, italic: false, code: false, strikethrough: false })
    expect(sel.Font.Bold).toBe(true)
  })

  it("sets font italic", () => {
    const sel = createMockSel()
    MarkdownRenderer.typeSeg(sel, { text: "i", bold: false, italic: true, code: false, strikethrough: false })
    expect(sel.Font.Italic).toBe(true)
  })

  it("sets font strikethrough", () => {
    const sel = createMockSel()
    MarkdownRenderer.typeSeg(sel, { text: "s", bold: false, italic: false, code: false, strikethrough: true })
    expect(sel.Font.Strikethrough).toBe(true)
  })

  it("sets Consolas for code segment", () => {
    const sel = createMockSel()
    MarkdownRenderer.typeSeg(sel, { text: "code", bold: false, italic: false, code: true, strikethrough: false })
    expect(sel.Font.Name).toBe("Consolas")
    expect(sel.Font.Size).toBe(10.5)
  })

  it("types text for code segment", () => {
    const sel = createMockSel()
    MarkdownRenderer.typeSeg(sel, { text: "code", bold: false, italic: false, code: true, strikethrough: false })
    expect(sel.TypeText).toHaveBeenCalled()
  })

  it("adds hyperlink for link segment", () => {
    const sel = createMockSel()
    MarkdownRenderer.typeSeg(sel, { text: "click", bold: false, italic: false, code: false, strikethrough: false, link: "https://x.com" })
    expect(sel.Range.Hyperlinks.Add).toHaveBeenCalled()
  })

  it("types text for plain segment", () => {
    const sel = createMockSel()
    MarkdownRenderer.typeSeg(sel, { text: "plain", bold: false, italic: false, code: false, strikethrough: false })
    expect(sel.TypeText).toHaveBeenCalledWith("plain")
  })
})

// ---------------------------------------------------------------------------
// Consecutive table rendering stability (regression test for table hang bug)
// ---------------------------------------------------------------------------
describe("MarkdownRenderer.renderBlock — consecutive tables", () => {
  function multiTableContext() {
    const sel = createMockSel()
    const doc = createMockDoc()
    doc.Content = { End: 0, Start: 0, Text: "" } as Record<string, unknown>
    const selectMock = vi.fn()
    doc.Range = vi.fn((_s?: number, _e?: number) => ({
      End: _e ?? 0,
      Start: _s ?? 0,
      Select: selectMock,
      Hyperlinks: { Add: vi.fn() },
      Shading: { BackgroundPatternColor: 0 },
    }))

    const selProxy = new MockSelectionProxy(sel)
    const docProxy = new MockDocumentProxy(doc)
    const goToEnd = vi.fn()
    selProxy.getParagraphFormat = vi.fn(() => sel.ParagraphFormat as Record<string, unknown>)
    selProxy.getFont = vi.fn(() => sel.Font as Record<string, unknown>)

    // Override getTables to return a controllable mock
    let tablesAdded = 0
    const tableAddMock = vi.fn((..._args: unknown[]) => {
      tablesAdded++
      const tableEnd = tablesAdded * 20
      ;(doc.Content as Record<string, unknown>).End = tableEnd
      return {
        Cell: vi.fn(() => ({ Range: { Text: "" } })),
        Rows: { Item: vi.fn(() => ({ Range: { Font: {} }, Shading: {} })), Alignment: 0 },
        AutoFitBehavior: vi.fn(),
        Range: { End: tableEnd, Start: tableEnd - 20 },
        Borders: { Item: vi.fn(() => ({ LineStyle: 0, ColorIndex: 0, LineWidth: 0 })) },
      }
    })
    vi.spyOn(docProxy, 'getTables').mockReturnValue({
      add: tableAddMock,
      count: 0,
      item: vi.fn(),
      acceptAll: vi.fn(),
      rejectAll: vi.fn(),
      addPicture: vi.fn(),
      addChart2: vi.fn(),
      addHorizontalLineStandard: vi.fn(),
      addTextbox: vi.fn(),
      itemByName: vi.fn(),
    } as never)

    const ctx: RenderContext = {
      getSelection: () => selProxy,
      requireDoc: () => docProxy,
      goToEnd,
      withScreenOff: async <T>(fn: () => Promise<T>) => fn(),
    }
    return { ctx, sel: selProxy, doc: docProxy, tableAddMock, selectMock }
  }

  it("renders 10 tables consecutively without calling endKey", async () => {
    const { ctx, tableAddMock, sel } = multiTableContext()
    const r = new MarkdownRenderer(ctx)

    for (let i = 0; i < 10; i++) {
      await r.renderBlock(
        { type: "table", rows: [[`T${i}-A`, `T${i}-B`], [`${i}`, `${i + 1}`]] } as Block,
        i,
        10,
        "end",
      )
    }

    // All 10 tables should be created
    expect(tableAddMock).toHaveBeenCalledTimes(10)
    // endKey should never be called (old buggy approach)
    expect(sel.endKey).not.toHaveBeenCalled()
  })

  it("uses doc.getRange for cursor positioning after each table", async () => {
    const { ctx, tableAddMock, doc } = multiTableContext()
    const r = new MarkdownRenderer(ctx)

    for (let i = 0; i < 3; i++) {
      await r.renderBlock(
        { type: "table", rows: [["A", "B"], ["1", "2"]] } as Block,
        i,
        3,
        "end",
      )
    }

    expect(tableAddMock).toHaveBeenCalledTimes(3)
    // doc.getRange should be called once per table for cursor positioning
    expect(doc.getRange).toHaveBeenCalledTimes(3)
  })

  it("renders table blocks interspersed with paragraph blocks", async () => {
    const { ctx, tableAddMock } = multiTableContext()
    const r = new MarkdownRenderer(ctx)

    const blocks: Block[] = [
      { type: "paragraph", text: "Intro" } as Block,
      { type: "table", rows: [["H1", "H2"], ["A", "B"]] } as Block,
      { type: "paragraph", text: "Middle" } as Block,
      { type: "table", rows: [["X", "Y"], ["1", "2"]] } as Block,
      { type: "paragraph", text: "End" } as Block,
    ]

    for (let i = 0; i < blocks.length; i++) {
      await r.renderBlock(blocks[i], i, blocks.length, "end")
    }

    expect(tableAddMock).toHaveBeenCalledTimes(2)
  })
})
