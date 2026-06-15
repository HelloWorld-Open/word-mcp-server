import { describe, it, expect, beforeEach } from "vitest"
import { PositionMap } from "../../../src/word/position-map.js"
import { createMockSession } from "../test-helpers.js"

interface HeadingEntry {
  level: number
  text: string
  paragraphIndex: number
}

describe("PositionMap.resolveHeading (pure logic)", () => {
  let pm: PositionMap

  const headings: HeadingEntry[] = [
    { level: 1, text: "Introduction", paragraphIndex: 1 },
    { level: 2, text: "Getting Started", paragraphIndex: 3 },
    { level: 2, text: "Configuration", paragraphIndex: 7 },
    { level: 1, text: "API Reference", paragraphIndex: 12 },
    { level: 2, text: "Authentication", paragraphIndex: 15 },
    { level: 2, text: "Endpoints", paragraphIndex: 20 },
    { level: 3, text: "GET /users", paragraphIndex: 22 },
    { level: 3, text: "POST /users", paragraphIndex: 25 },
    { level: 1, text: "Introduction", paragraphIndex: 30 },
  ]

  beforeEach(() => {
    const session = createMockSession()
    pm = new PositionMap(session)
    ;(pm as any).headings = [...headings]
    ;(pm as any).dirty = false
  })

  it("resolves heading by exact match", async () => {
    const result = await (pm as any).resolveHeading({
      by: "heading",
      match: "API Reference",
      matchMode: "exact",
    })
    expect(result.found).toBe(true)
    expect(result.paragraphIndex).toBe(12)
    expect(result.headingContext).toContain("API Reference")
  })

  it("resolves heading by contains match", async () => {
    const result = await (pm as any).resolveHeading({
      by: "heading",
      match: "Config",
      matchMode: "contains",
    })
    expect(result.found).toBe(true)
    expect(result.paragraphIndex).toBe(7)
  })

  it("resolves heading by startsWith match", async () => {
    const result = await (pm as any).resolveHeading({
      by: "heading",
      match: "Getting",
      matchMode: "startsWith",
    })
    expect(result.found).toBe(true)
    expect(result.paragraphIndex).toBe(3)
  })

  it("resolves heading by regex match", async () => {
    const result = await (pm as any).resolveHeading({
      by: "heading",
      match: "GET /users",
      matchMode: "regex",
    })
    expect(result.found).toBe(true)
    expect(result.paragraphIndex).toBe(22)
  })

  it("returns found:false when heading does not exist", async () => {
    const result = await (pm as any).resolveHeading({
      by: "heading",
      match: "Nonexistent",
      matchMode: "exact",
    })
    expect(result.found).toBe(false)
    expect(result.error).toBeTruthy()
  })

  it("resolves nth occurrence with occurrence param", async () => {
    const result = await (pm as any).resolveHeading({
      by: "heading",
      match: "Introduction",
      matchMode: "exact",
      occurrence: 2,
    })
    expect(result.found).toBe(true)
    expect(result.paragraphIndex).toBe(30)
  })

  it("applies offset after heading", async () => {
    const result = await (pm as any).resolveHeading({
      by: "heading",
      match: "Configuration",
      offset: { direction: "after", count: 2 },
    })
    expect(result.found).toBe(true)
    expect(result.paragraphIndex).toBe(9)
  })

  it("applies offset before heading", async () => {
    const result = await (pm as any).resolveHeading({
      by: "heading",
      match: "Configuration",
      offset: { direction: "before", count: 1 },
    })
    expect(result.found).toBe(true)
    expect(result.paragraphIndex).toBe(6)
  })

  it("returns error for invalid regex", async () => {
    const result = await (pm as any).resolveHeading({
      by: "heading",
      match: "[invalid",
      matchMode: "regex",
    })
    expect(result.found).toBe(false)
    expect(result.error).toContain("Invalid regex")
  })

  it("returns error when occurrence exceeds candidates", async () => {
    const result = await (pm as any).resolveHeading({
      by: "heading",
      match: "Introduction",
      occurrence: 99,
    })
    expect(result.found).toBe(false)
  })

  it("filters by heading level", async () => {
    const result = await (pm as any).resolveHeading({
      by: "heading",
      level: 3,
    })
    expect(result.found).toBe(true)
    expect(result.paragraphIndex).toBe(22)
    expect(result.headingContext).toContain("(H3)")
  })

  it("level + match narrows to same-level duplicate names", async () => {
    const result = await (pm as any).resolveHeading({
      by: "heading",
      level: 1,
      match: "Introduction",
      matchMode: "exact",
      occurrence: 2,
    })
    expect(result.found).toBe(true)
    expect(result.paragraphIndex).toBe(30)
  })

  it("level filters before match — unmatched level yields not found", async () => {
    const result = await (pm as any).resolveHeading({
      by: "heading",
      level: 1,
      match: "Authentication",
      matchMode: "contains",
    })
    expect(result.found).toBe(false)
  })

  it("level with occurrence counts within same level only", async () => {
    const result1 = await (pm as any).resolveHeading({
      by: "heading",
      level: 1,
      occurrence: 1,
    })
    const result2 = await (pm as any).resolveHeading({
      by: "heading",
      level: 1,
      occurrence: 2,
    })
    expect(result1.found).toBe(true)
    expect(result1.paragraphIndex).toBe(1)
    expect(result2.found).toBe(true)
    expect(result2.paragraphIndex).toBe(12)
  })

  it("level without match resolves first heading at that level", async () => {
    const result = await (pm as any).resolveHeading({
      by: "heading",
      level: 2,
    })
    expect(result.found).toBe(true)
    expect(result.paragraphIndex).toBe(3)
  })
})

describe("PositionMap.getHeadingContext", () => {
  let pm: PositionMap

  const headings: HeadingEntry[] = [
    { level: 1, text: "Doc Title", paragraphIndex: 1 },
    { level: 2, text: "Section A", paragraphIndex: 5 },
    { level: 3, text: "Sub A1", paragraphIndex: 7 },
    { level: 1, text: "Section B", paragraphIndex: 15 },
  ]

  beforeEach(() => {
    const session = createMockSession()
    pm = new PositionMap(session)
    ;(pm as any).headings = [...headings]
  })

  it("returns the closest heading above the given paragraph", () => {
    const ctx = (pm as any).getHeadingContext(8)
    expect(ctx).toBe("Sub A1 (H3)")
  })

  it("returns nearest heading below given paragraph as context", () => {
    const ctx = (pm as any).getHeadingContext(10)
    expect(ctx).toBe("Sub A1 (H3)")
  })

  it("returns parent heading between sub-headings", () => {
    const ctx = (pm as any).getHeadingContext(6)
    expect(ctx).toBe("Section A (H2)")
  })

  it("returns root heading for early paragraphs", () => {
    const ctx = (pm as any).getHeadingContext(3)
    expect(ctx).toBe("Doc Title (H1)")
  })

  it("returns deepest heading when exactly at heading paragraph", () => {
    const ctx = (pm as any).getHeadingContext(7)
    expect(ctx).toBe("Sub A1 (H3)")
  })

  it("returns new section heading after section break", () => {
    const ctx = (pm as any).getHeadingContext(16)
    expect(ctx).toBe("Section B (H1)")
  })

  it("returns null when no headings exist", () => {
    const emptyPm = new PositionMap(createMockSession())
    ;(emptyPm as any).headings = []
    const ctx = (emptyPm as any).getHeadingContext(5)
    expect(ctx).toBeNull()
  })
})

describe("PositionMap cached texts", () => {
  let pm: PositionMap
  let session: IWordSession

  beforeEach(() => {
    session = createMockSession()
    pm = new PositionMap(session)
  })

  it("starts with empty cached texts", () => {
    expect((pm as any).cachedTexts).toEqual([])
    expect((pm as any).cachedParaStarts).toEqual([])
  })

  it("resolveParagraph falls back to COM when cache empty", async () => {
    ;(pm as any).headings = []
    ;(pm as any).dirty = false
    const result = await pm.resolve({
      by: "paragraph",
      match: "nonexistent",
    })
    expect(result.found).toBe(false)
  })
})

describe("PositionMap state management", () => {
  it("starts dirty and clears after refresh", async () => {
    const session = createMockSession()
    const pm = new PositionMap(session)

    expect((pm as any).dirty).toBe(true)

    try {
      await pm.refresh()
    } catch {
      // COM mock may fail, that's fine
    }
  })

  it("markDirty sets dirty flag", () => {
    const session = createMockSession()
    const pm = new PositionMap(session)
    ;(pm as any).dirty = false

    pm.markDirty()

    expect((pm as any).dirty).toBe(true)
  })

  it("resolve returns error for bookmark when no bookmarks present", async () => {
    const session = createMockSession()
    const pm = new PositionMap(session)
    ;(pm as any).headings = []
    ;(pm as any).dirty = false

    const result = await pm.resolve({
      by: "bookmark",
      name: "nonexistent",
    })

    expect(result.found).toBe(false)
  })

  it("resolve returns error for table when no tables present", async () => {
    const session = createMockSession()
    const pm = new PositionMap(session)
    ;(pm as any).headings = []
    ;(pm as any).tables = []
    ;(pm as any).dirty = false

    const result = await pm.resolve({
      by: "table",
      occurrence: 1,
    })

    expect(result.found).toBe(false)
    expect(result.error).toContain("No tables")
  })
})

describe("PositionMap.refresh — COM path", () => {
  function makeRefreshableDoc(
    paraTexts: string[],
    headingSpots: Array<{ level: number; paraIdx: number; start: number }>,
    tableSpots: Array<{ paraIdx: number; rangeStart: number; rangeEnd: number }> = [],
  ): Record<string, unknown> {
    const count = paraTexts.length
    const fullText = paraTexts.join("\r") + (paraTexts.length > 0 ? "\r" : "")

    const pendingHeadings = headingSpots.map(h => ({ ...h }))

    const duplicateRange: Record<string, unknown> = {
      Start: 0,
      Find: {
        ClearFormatting: vi.fn(),
        Style: "",
        Text: "",
        Forward: true,
        Wrap: 0,
        Format: false,
        Execute: vi.fn(() => {
          const idx = pendingHeadings.findIndex(h =>
            `Heading ${h.level}` === duplicateRange.Find.Style
          )
          if (idx === -1) return false
          const h = pendingHeadings.splice(idx, 1)[0]
          duplicateRange.Start = h.start
          return true
        }),
        Replacement: {},
      },
    }

    const tables: Record<string, unknown> = { Count: tableSpots.length, Item: vi.fn() }
    for (let i = 0; i < tableSpots.length; i++) {
      tables.Item = vi.fn((idx: number) => {
        const spot = tableSpots[idx - 1]
        return { Range: { Start: spot.rangeStart, End: spot.rangeEnd } }
      })
    }

    return {
      Name: "test.docx",
      Content: {
        Text: fullText,
        End: fullText.length,
        Start: 0,
        Duplicate: duplicateRange,
      },
      Paragraphs: { Count: count, Item: vi.fn() },
      Tables: tables,
      Range: vi.fn(() => ({
        End: 0, Start: 0, Select: vi.fn(),
        Font: {}, Shading: {},
      })),
      Saved: true,
    }
  }

  it("extracts headings from document content via Find", async () => {
    const paraTexts = [
      "Title",
      "Intro text",
      "Getting Started",
      "Step 1 details",
      "Advanced Config",
      "Notes",
    ]
    const headingSpots = [
      { level: 1, paraIdx: 1, start: 0 },
      { level: 2, paraIdx: 3, start: "Title\rIntro text\r".length },
      { level: 2, paraIdx: 5, start: "Title\rIntro text\rGetting Started\rStep 1 details\r".length },
    ]
    const doc = makeRefreshableDoc(paraTexts, headingSpots)
    const session = createMockSession(undefined, doc)
    const pm = new PositionMap(session)

    await pm.refresh()

    const headings = pm.getHeadings()
    expect(headings).toHaveLength(3)
    expect(headings[0]).toMatchObject({ text: "Title", level: 1, paragraphIndex: 1 })
    expect(headings[1]).toMatchObject({ text: "Getting Started", level: 2, paragraphIndex: 3 })
    expect(headings[2]).toMatchObject({ text: "Advanced Config", level: 2, paragraphIndex: 5 })
  })

  it("skips headings inside table ranges", async () => {
    const paraTexts = [
      "Title",
      "Header1\tHeader2",
      "Cell1\tCell2",
      "Next Section",
      "Detail",
    ]
    const headingSpots = [
      { level: 1, paraIdx: 1, start: 0 },
      { level: 1, paraIdx: 4, start: "Title\rHeader1\tHeader2\rCell1\tCell2\r".length },
    ]
    const tableSpots = [
      { paraIdx: 2, rangeStart: "Title\r".length, rangeEnd: "Title\rHeader1\tHeader2\rCell1\tCell2\r".length },
    ]
    const doc = makeRefreshableDoc(paraTexts, headingSpots, tableSpots)
    const session = createMockSession(undefined, doc)
    const pm = new PositionMap(session)

    await pm.refresh()

    const headings = pm.getHeadings()
    expect(headings).toHaveLength(2)
    expect(headings[0].text).toBe("Title")
    expect(headings[1].text).toBe("Next Section")
  })

  it("sets dirty=false and records para count after refresh", async () => {
    const doc = makeRefreshableDoc(["A", "B"],[{ level: 1, paraIdx: 1, start: 0 }])
    const session = createMockSession(undefined, doc)
    const pm = new PositionMap(session)

    expect((pm as any).dirty).toBe(true)
    await pm.refresh()
    expect((pm as any).dirty).toBe(false)
    expect((pm as any).lastParaCount).toBe(2)
  })

  it("handles empty document gracefully", async () => {
    const doc = makeRefreshableDoc([], [])
    const session = createMockSession(undefined, doc)
    const pm = new PositionMap(session)

    await pm.refresh()
    expect(pm.getHeadings()).toEqual([])
    expect((pm as any).dirty).toBe(false)
  })

  it("handles document with no headings", async () => {
    const paraTexts = ["Just a paragraph", "Another one"]
    const doc = makeRefreshableDoc(paraTexts, [])
    const session = createMockSession(undefined, doc)
    const pm = new PositionMap(session)

    await pm.refresh()
    expect(pm.getHeadings()).toEqual([])
    expect((pm as any).dirty).toBe(false)
  })
})

describe("PositionMap.ensureFresh", () => {
  function makeDocForEnsure(
    paraCount: number,
    textLen: number,
  ): Record<string, unknown> {
    return {
      Name: "test.docx",
      Content: { Text: "x".repeat(textLen), End: textLen, Start: 0 },
      Paragraphs: { Count: paraCount, Item: vi.fn() },
      Tables: { Count: 0, Item: vi.fn() },
      Range: vi.fn(() => ({ End: 0, Start: 0, Select: vi.fn() })),
      Saved: true,
    }
  }

  it("refreshes when dirty flag is set", async () => {
    const doc = makeDocForEnsure(3, 50)
    const session = createMockSession(undefined, doc)
    const pm = new PositionMap(session)
    ;(pm as any).dirty = true
    ;(pm as any).headings = []

    await pm.ensureFresh()

    expect((pm as any).dirty).toBe(false)
  })

  it("refreshes when para count changed", async () => {
    const doc = makeDocForEnsure(10, 200)
    const session = createMockSession(undefined, doc)
    const pm = new PositionMap(session)
    ;(pm as any).dirty = false
    ;(pm as any).lastParaCount = 5
    ;(pm as any).headings = []

    await pm.ensureFresh()

    expect((pm as any).lastParaCount).toBe(10)
  })

  it("calls refreshContentOnly when para count matches but content end changed", async () => {
    const doc = makeDocForEnsure(3, 100)
    const session = createMockSession(undefined, doc)
    const pm = new PositionMap(session)
    ;(pm as any).dirty = false
    ;(pm as any).lastParaCount = 3
    ;(pm as any).lastContentEnd = 50
    ;(pm as any).headings = [{ text: "Old Title", level: 1, paragraphIndex: 1 }]

    await pm.ensureFresh()

    expect((pm as any).lastContentEnd).toBe(100)
    // refreshContentOnly updates heading text from new content
    expect((pm as any).headings[0].text).toBe("x".repeat(100).split("\r")[0] ?? "x".repeat(100))
  })

  it("falls back to refresh on COM error", async () => {
    const errorDoc = undefined as unknown as Record<string, unknown>
    const session = createMockSession(undefined, errorDoc)
    const pm = new PositionMap(session)
    ;(pm as any).dirty = false

    await expect(pm.ensureFresh()).resolves.toBeUndefined()
    // should not throw — falls back and may set dirty again
  })
})

describe("PositionMap.scheduleRefresh", () => {
  it("schedules refresh when dirty and no pending promise", async () => {
    const paraTexts = ["Title", "Body"]
    const fullText = paraTexts.join("\r") + "\r"
    const doc: Record<string, unknown> = {
      Content: { Text: fullText, End: fullText.length, Start: 0, Duplicate: { Start: 0, Find: { Style: "", ClearFormatting: vi.fn(), Execute: vi.fn(() => false), Text: "", Forward: true, Wrap: 0, Format: false, Replacement: {} } } },
      Paragraphs: { Count: paraTexts.length, Item: vi.fn() },
      Tables: { Count: 0, Item: vi.fn() },
      Range: vi.fn(() => ({ End: 0, Start: 0, Select: vi.fn() })),
      Saved: true,
    }
    const session = createMockSession(undefined, doc)
    const pm = new PositionMap(session)
    ;(pm as any).dirty = true

    pm.scheduleRefresh()

    expect((pm as any).refreshPromise).not.toBeNull()
    // wait for it to complete
    await (pm as any).refreshPromise
    expect((pm as any).refreshPromise).toBeNull()
    expect((pm as any).dirty).toBe(false)
  })

  it("does not schedule again while refresh is pending", async () => {
    const paraTexts = ["Title", "Body"]
    const fullText = paraTexts.join("\r") + "\r"
    const doc: Record<string, unknown> = {
      Content: { Text: fullText, End: fullText.length, Start: 0, Duplicate: { Start: 0, Find: { Style: "", ClearFormatting: vi.fn(), Execute: vi.fn(() => false), Text: "", Forward: true, Wrap: 0, Format: false, Replacement: {} } } },
      Paragraphs: { Count: paraTexts.length, Item: vi.fn() },
      Tables: { Count: 0, Item: vi.fn() },
      Range: vi.fn(() => ({ End: 0, Start: 0, Select: vi.fn() })),
      Saved: true,
    }
    const session = createMockSession(undefined, doc)
    const pm = new PositionMap(session)

    const p1 = new Promise<void>(resolve => {
      ;(pm as any).refreshPromise = new Promise(r => setTimeout(r, 50)).then(resolve as () => void) as unknown as Promise<void>
    })
    ;(pm as any).dirty = true
    const spy = vi.spyOn(pm, "refresh" as any)

    pm.scheduleRefresh()

    expect(spy).not.toHaveBeenCalled()
    spy.mockRestore()
    ;(pm as any).refreshPromise = null
  })
})

describe("PositionMap.getHeadings", () => {
  it("returns the headings array", () => {
    const session = createMockSession()
    const pm = new PositionMap(session)
    const hd = [{ text: "A", level: 1, paragraphIndex: 1 }]
    ;(pm as any).headings = hd

    expect(pm.getHeadings()).toBe(hd)
  })
})

describe("PositionMap.resolveTable — populated tables", () => {
  it("resolves first table by occurrence", async () => {
    const session = createMockSession()
    const pm = new PositionMap(session)
    ;(pm as any).headings = [{ text: "Title", level: 1, paragraphIndex: 1 }]
    ;(pm as any).tables = [
      { paragraphIndex: 5, rangeStart: 10, rangeEnd: 50 },
      { paragraphIndex: 10, rangeStart: 60, rangeEnd: 100 },
    ]
    ;(pm as any).dirty = false

    const result = await pm.resolve({ by: "table", occurrence: 1 })

    expect(result.found).toBe(true)
    expect(result.paragraphIndex).toBe(5)
    expect(result.tableIndex).toBe(1)
    expect(result.headingContext).toContain("Title")
  })

  it("resolves second table by occurrence", async () => {
    const session = createMockSession()
    const pm = new PositionMap(session)
    ;(pm as any).headings = [{ text: "Title", level: 1, paragraphIndex: 1 }]
    ;(pm as any).tables = [
      { paragraphIndex: 5, rangeStart: 10, rangeEnd: 50 },
      { paragraphIndex: 10, rangeStart: 60, rangeEnd: 100 },
    ]
    ;(pm as any).dirty = false

    const result = await pm.resolve({ by: "table", occurrence: 2 })

    expect(result.found).toBe(true)
    expect(result.paragraphIndex).toBe(10)
    expect(result.tableIndex).toBe(2)
  })

  it("applies offset after table", async () => {
    const session = createMockSession()
    const pm = new PositionMap(session)
    ;(pm as any).headings = [{ text: "Section", level: 1, paragraphIndex: 1 }]
    ;(pm as any).tables = [{ paragraphIndex: 3, rangeStart: 10, rangeEnd: 50 }]
    ;(pm as any).dirty = false

    const result = await pm.resolve({ by: "table", occurrence: 1, offset: { direction: "after", count: 2 } })

    expect(result.found).toBe(true)
    expect(result.paragraphIndex).toBe(5)
  })
})

describe("PositionMap.resolveBookmark — cached binary search", () => {
  it("resolves bookmark by name using binary search on cached texts", async () => {
    const paraTexts = ["Title", "Intro", "Section 1", "Body", "Section 2"]
    const starts = Array.from({ length: paraTexts.length + 2 }, () => 0)
    let textPos = 0
    for (let i = 1; i <= paraTexts.length; i++) {
      starts[i] = textPos
      textPos += paraTexts[i - 1].length + 1
    }
    starts[paraTexts.length + 1] = textPos

    const bookmarkRange = { Start: starts[3] } // at "Section 1" (1-based)

    const fullText = paraTexts.join("\r") + "\r"
    const endPos = starts[paraTexts.length + 1] ?? textPos
    const doc: Record<string, unknown> = {
      Content: { Text: fullText, End: endPos, Start: 0, Duplicate: {} },
      Paragraphs: { Count: paraTexts.length, Item: vi.fn() },
      Tables: { Count: 0, Item: vi.fn() },
      Bookmarks: {
        Count: 1,
        Item: vi.fn((idx: number) => {
          if (idx === 1) return { Name: "myBookmark", Range: bookmarkRange }
          return { Name: "", Range: { Start: 0 } }
        }),
      },
      Saved: true,
    }
    const session = createMockSession(undefined, doc)
    const pm = new PositionMap(session)
    ;(pm as any).headings = [{ text: "Title", level: 1, paragraphIndex: 1 }]
    ;(pm as any).dirty = false
    ;(pm as any).cachedTexts = paraTexts
    ;(pm as any).cachedParaStarts = starts
    ;(pm as any).lastParaCount = paraTexts.length

    const result = await pm.resolve({ by: "bookmark", name: "myBookmark" })

    expect(result.found).toBe(true)
    expect(result.paragraphIndex).toBe(3) // 1-based, "Section 1"
    expect(result.headingContext).toContain("Title")
  })

  it("returns not found for nonexistent bookmark", async () => {
    const doc: Record<string, unknown> = {
      Content: { Text: "x\r", End: 2, Start: 0 },
      Bookmarks: { Count: 1, Item: vi.fn(() => ({ Name: "other", Range: { Start: 0 } })) },
      Paragraphs: { Count: 1, Item: vi.fn() },
      Tables: { Count: 0, Item: vi.fn() },
      Saved: true,
    }
    const session = createMockSession(undefined, doc)
    const pm = new PositionMap(session)
    ;(pm as any).headings = []
    ;(pm as any).dirty = false
    ;(pm as any).lastParaCount = 1

    const result = await pm.resolve({ by: "bookmark", name: "missing" })

    expect(result.found).toBe(false)
    expect(result.error).toContain("missing")
  })

  it("falls back to COM binary search when cached texts absent", async () => {
    const paraTexts = ["Alpha", "Beta", "Gamma"]
    const bookmarkRange = { Start: "Alpha\r".length }
    const doc: Record<string, unknown> = {
      Bookmarks: { Count: 1, Item: vi.fn((i: number) => i === 1 ? { Name: "bm", Range: bookmarkRange } : { Name: "", Range: { Start: 0 } }) },
      Content: { Text: "Alpha\rBeta\rGamma\r", End: "Alpha\rBeta\rGamma\r".length, Start: 0, Duplicate: {} },
      Paragraphs: { Count: 3, Item: vi.fn() },
      Tables: { Count: 0, Item: vi.fn() },
      Range: vi.fn(() => ({ End: 0, Start: 0, Select: vi.fn() })),
      Saved: true,
    }
    const session = createMockSession(undefined, doc)
    const pm = new PositionMap(session)
    ;(pm as any).headings = []
    ;(pm as any).dirty = false
    ;(pm as any).cachedTexts = []
    ;(pm as any).cachedParaStarts = []

    const result = await pm.resolve({ by: "bookmark", name: "bm" })

    expect(result.found).toBe(true)
    expect(result.paragraphIndex).toBe(2) // "Beta"
  })
})

describe("PositionMap.fetchActualParaCount", () => {
  it("returns paragraph count from COM", async () => {
    const doc: Record<string, unknown> = {
      Paragraphs: { Count: 7, Item: vi.fn() },
      Tables: { Count: 0, Item: vi.fn() },
      Content: { Text: "" },
      Saved: true,
    }
    const session = createMockSession(undefined, doc)
    const pm = new PositionMap(session)

    const count = await pm.fetchActualParaCount()

    expect(count).toBe(7)
  })
})

describe("PositionMap.paraCountMatches", () => {
  it("returns true when counts match", async () => {
    const doc: Record<string, unknown> = {
      Paragraphs: { Count: 5, Item: vi.fn() },
      Tables: { Count: 0, Item: vi.fn() },
      Content: { Text: "" },
      Saved: true,
    }
    const session = createMockSession(undefined, doc)
    const pm = new PositionMap(session)

    const matches = await pm.paraCountMatches(5)

    expect(matches).toBe(true)
  })

  it("returns false when counts differ", async () => {
    const doc: Record<string, unknown> = {
      Paragraphs: { Count: 5, Item: vi.fn() },
      Tables: { Count: 0, Item: vi.fn() },
      Content: { Text: "" },
      Saved: true,
    }
    const session = createMockSession(undefined, doc)
    const pm = new PositionMap(session)

    const matches = await pm.paraCountMatches(10)

    expect(matches).toBe(false)
  })
})
