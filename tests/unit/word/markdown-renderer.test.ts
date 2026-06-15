import { describe, it, expect, vi } from "vitest"
import { MarkdownRenderer, type RenderContext } from "../../../src/word/markdown-renderer.js"
import { createMockDoc, createMockSel } from "../test-helpers.js"
import type { Block } from "../../../src/word/markdown-parser.js"

function createRenderContext(
  customSel?: Record<string, unknown>,
  customDoc?: Record<string, unknown>,
): { ctx: RenderContext; sel: Record<string, unknown>; doc: Record<string, unknown>; goToEnd: ReturnType<typeof vi.fn> } {
  const sel = customSel ?? createMockSel()
  const doc = customDoc ?? createMockDoc()
  const goToEnd = vi.fn()
  return {
    ctx: {
      getSelection: () => sel,
      requireDoc: () => doc,
      goToEnd,
    },
    sel, doc, goToEnd,
  }
}

describe("MarkdownRenderer.renderBlock", () => {
  it("renders hr via AddHorizontalLineStandard", async () => {
    const { ctx, sel } = createRenderContext()
    const r = new MarkdownRenderer(ctx)
    await r.renderBlock({ type: "hr" } as Block, 0, 1, "end")
    expect(sel.InlineShapes.AddHorizontalLineStandard).toHaveBeenCalled()
    expect(sel.TypeParagraph).toHaveBeenCalled()
  })

  it("renders hr at cursor without goToEnd", async () => {
    const { ctx, goToEnd } = createRenderContext()
    const r = new MarkdownRenderer(ctx)
    await r.renderBlock({ type: "hr" } as Block, 0, 1, "cursor")
    expect(goToEnd).not.toHaveBeenCalled()
  })

  it("renders pagebreak via InsertBreak(7)", async () => {
    const { ctx, sel } = createRenderContext()
    const r = new MarkdownRenderer(ctx)
    const chars = await r.renderBlock({ type: "pagebreak" } as Block, 0, 1, "end")
    expect(sel.InsertBreak).toHaveBeenCalledWith(7)
    expect(chars).toBe(0)
  })

  it("renders image via AddPicture", async () => {
    const { ctx, doc } = createRenderContext()
    const r = new MarkdownRenderer(ctx)
    const chars = await r.renderBlock({ type: "image", alt: "pic", url: "C:\\img.png" } as Block, 0, 1, "end")
    expect(doc.InlineShapes.AddPicture).toHaveBeenCalledWith("C:\\img.png")
    expect(chars).toBe(3 + 10)
  })

  it("fallback image to text when AddPicture throws", async () => {
    const doc = createMockDoc()
    doc.InlineShapes.AddPicture = vi.fn(() => { throw new Error("COM error") })
    const { ctx } = createRenderContext(createMockSel(), doc)
    const r = new MarkdownRenderer(ctx)
    await r.renderBlock({ type: "image", alt: "err", url: "x.png" } as Block, 0, 1, "end")
    expect(ctx.getSelection().TypeText).toHaveBeenCalledWith("[图片: err]")
  })

  it("renders heading with style and paragraph after", async () => {
    const { ctx, sel } = createRenderContext()
    const r = new MarkdownRenderer(ctx)
    const chars = await r.renderBlock({ type: "heading", level: 2, text: "Section" } as Block, 0, 2, "end")
    expect(sel.TypeText).toHaveBeenCalledWith("Section")
    expect(sel.Style).toContain("Heading 2")
    expect(sel.TypeParagraph).toHaveBeenCalled()
    expect(chars).toBe(7)
  })

  it("renders paragraph without heading style", async () => {
    const { ctx, sel } = createRenderContext()
    const r = new MarkdownRenderer(ctx)
    await r.renderBlock({ type: "paragraph", text: "Body" } as Block, 0, 2, "end")
    expect(sel.TypeText).toHaveBeenCalledWith("Body")
    expect(sel.Style).not.toContain("Heading")
    expect(sel.TypeParagraph).toHaveBeenCalled()
  })

  it("renders paragraph without trailing paragraph break for last block", async () => {
    const sel = createMockSel()
    const { ctx } = createRenderContext(sel)
    const r = new MarkdownRenderer(ctx)
    await r.renderBlock({ type: "paragraph", text: "Last" } as Block, 0, 1, "end")
    expect(sel.TypeParagraph).not.toHaveBeenCalled()
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
    const { ctx, sel } = createRenderContext()
    const r = new MarkdownRenderer(ctx)
    await r.renderBlock({ type: "bullet_list", items: [{ text: "a", indent: 0 }, { text: "b", indent: 0 }] } as Block, 0, 1, "end")
    expect(sel.Range.ListFormat.ApplyBulletDefault).toHaveBeenCalled()
    expect(sel.TypeText).toHaveBeenNthCalledWith(1, "a")
    expect(sel.TypeText).toHaveBeenNthCalledWith(2, "b")
  })

  it("renders bullet list with indentation", async () => {
    const { ctx, sel } = createRenderContext()
    const r = new MarkdownRenderer(ctx)
    await r.renderBlock({ type: "bullet_list", items: [{ text: "deep", indent: 2 }] } as Block, 0, 1, "end")
    expect(sel.Range.ListFormat.IncreaseIndent).toHaveBeenCalledTimes(2)
  })

  it("renders numbered list with ApplyNumberDefault", async () => {
    const { ctx, sel } = createRenderContext()
    const r = new MarkdownRenderer(ctx)
    await r.renderBlock({ type: "numbered_list", items: [{ text: "1", indent: 0 }] } as Block, 0, 1, "end")
    expect(sel.Range.ListFormat.ApplyNumberDefault).toHaveBeenCalled()
  })

  it("renders table with Tables.Add", async () => {
    const doc = createMockDoc()
    const tableMock = {
      Cell: vi.fn(() => ({ Range: { Text: "" } })),
      Rows: { Item: vi.fn(() => ({ Range: { Font: {} }, Shading: {} })), Alignment: 0 },
      AutoFitBehavior: vi.fn(),
    }
    doc.Tables = { Add: vi.fn(() => tableMock), Count: 1, Item: vi.fn() }
    const { ctx } = createRenderContext(createMockSel(), doc)
    const r = new MarkdownRenderer(ctx)
    await r.renderBlock({ type: "table", rows: [["A", "B"], ["1", "2"]] } as Block, 0, 1, "end")
    expect(doc.Tables.Add).toHaveBeenCalledWith(expect.anything(), 2, 2)
  })

  it("renders blockquote with indent, italic, and paragraph break", async () => {
    const sel = createMockSel()
    sel.Range.Borders = { Item: vi.fn(() => ({})) }
    const { ctx } = createRenderContext(sel)
    const r = new MarkdownRenderer(ctx)
    await r.renderBlock({ type: "blockquote", text: "quote" } as Block, 0, 2, "end")
    expect(sel.ParagraphFormat.LeftIndent).toBe(0)
    expect(sel.Font.Italic).toBe(false)
    expect(sel.TypeParagraph).toHaveBeenCalled()
  })

  it("renders blockquote without trailing paragraph for last block", async () => {
    const sel = createMockSel()
    const { ctx } = createRenderContext(sel)
    const r = new MarkdownRenderer(ctx)
    await r.renderBlock({ type: "blockquote", text: "last" } as Block, 0, 1, "end")
    expect(sel.TypeParagraph).not.toHaveBeenCalled()
  })

  it("renders codeblock with Consolas font", async () => {
    const doc = createMockDoc()
    const rangeMock = { Font: {}, Shading: { BackgroundPatternColor: 0 }, Select: vi.fn() }
    doc.Range = vi.fn(() => rangeMock)
    doc.Content.End = 10
    const sel = createMockSel()
    const { ctx } = createRenderContext(sel, doc)
    const r = new MarkdownRenderer(ctx)
    await r.renderBlock({ type: "codeblock", text: "code" } as Block, 0, 1, "end")
    expect(sel.TypeText).toHaveBeenCalledWith("code")
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
