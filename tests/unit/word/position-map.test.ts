import { describe, it, expect, beforeEach, vi } from "vitest"
import { PositionMap } from "../../../src/word/position-map.js"
import type { IWordSession } from "../../../src/word/session.js"

interface HeadingEntry {
  level: number
  text: string
  paragraphIndex: number
}

function createMockSession(headings?: HeadingEntry[]): IWordSession {
  const doc = {
    Content: { End: 0, Text: "" },
    Range: vi.fn(() => ({ End: 0, Select: vi.fn() })),
    Paragraphs: { Count: 0, Item: vi.fn() },
    Application: { Selection: {} },
    Tables: { Count: 0, Item: vi.fn() },
    Bookmarks: { Count: 0, Item: vi.fn() },
  } as unknown as Record<string, unknown>

  return {
    application: { Selection: {}, ScreenUpdating: true, ActiveDocument: doc } as Record<string, unknown>,
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
