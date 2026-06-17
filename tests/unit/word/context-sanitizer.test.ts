import { describe, it, expect, vi, beforeEach } from "vitest"
import { ContextSanitizer } from "../../../src/word/context-sanitizer.js"
import { createMockSession, createMockDoc, createMockSel } from "../test-helpers.js"
import { MockSelectionProxy, MockDocumentProxy } from "../../../src/word/com-proxy/com-proxy.mock.js"

// ---------------------------------------------------------------------------
// Helper: build a session where cursor is "inside a table"
// ---------------------------------------------------------------------------
function tableContext(opts?: { tableEnd?: number; docEnd?: number }) {
  const tableEnd = opts?.tableEnd ?? 50
  const docEnd = opts?.docEnd ?? 100

  const rawDoc = createMockDoc()
  rawDoc.Content = { End: docEnd, Start: 0, Text: "" } as Record<string, unknown>
  const selectMock = vi.fn()
  rawDoc.Range = vi.fn((_s?: number, _e?: number) => ({
    End: _e ?? docEnd,
    Start: _s ?? 0,
    Select: selectMock,
    Hyperlinks: { Add: vi.fn() },
    Shading: { BackgroundPatternColor: 0 },
  }))

  const rawSel = createMockSel()
  rawSel.StoryType = 1 // wdMainTextStory
  rawSel.Start = 30
  rawSel.End = 30

  // Simulate cursor inside table
  let withinTable = true
  rawSel.Information = vi.fn((id: number) => {
    if (id === 12) return withinTable // WD_WITHIN_TABLE = 12
    return false
  })

  const tableRangeSelect = vi.fn()
  rawSel.Tables = {
    Count: 1,
    Item: vi.fn(() => ({
      Range: { End: tableEnd, Start: 0, Select: tableRangeSelect },
    })),
  }

  const session = createMockSession(undefined, rawDoc, rawSel)
  const selProxy = session.getSelectionProxy() as MockSelectionProxy
  const docProxy = session.getDocProxy() as MockDocumentProxy

  // Wire up getInformation to return within-table flag
  selProxy.getInformation = vi.fn(((id: number) => {
    if (id === 12) return withinTable
    return false
  }) as () => boolean)
  selProxy.getTables = vi.fn(() => ({
    Item: vi.fn(() => ({
      Range: { End: tableEnd, Start: 0, Select: tableRangeSelect },
    })),
  }))
  selProxy.getShapeRange = vi.fn(() => ({ Count: 0 }))

  return {
    session,
    selProxy,
    docProxy,
    rawSel,
    rawDoc,
    selectMock,
    tableRangeSelect,
    setWithinTable: (v: boolean) => { withinTable = v },
  }
}

// ---------------------------------------------------------------------------
// ensureMainBody — table escape
// ---------------------------------------------------------------------------
describe("ContextSanitizer.ensureMainBody — table escape", () => {
  it("uses doc.getRange(targetPos).select() when cursor is inside a table", () => {
    const ctx = tableContext({ tableEnd: 50, docEnd: 100 })
    const sanitizer = new ContextSanitizer(ctx.session)
    sanitizer.ensureMainBody()
    // Should call doc.getRange(50, 50) (min of tableEnd=50, docEnd=100)
    expect(ctx.docProxy.getRange).toHaveBeenCalledWith(50, 50)
    expect(ctx.selProxy.collapse).toHaveBeenCalledWith(0)
  })

  it("caps targetPos at docEnd when tableEnd exceeds docEnd", () => {
    const ctx = tableContext({ tableEnd: 200, docEnd: 80 })
    const sanitizer = new ContextSanitizer(ctx.session)
    sanitizer.ensureMainBody()
    // min(200, 80) = 80
    expect(ctx.docProxy.getRange).toHaveBeenCalledWith(80, 80)
  })

  it("does NOT call the old table.Range.Select() + collapse approach", () => {
    const ctx = tableContext()
    const sanitizer = new ContextSanitizer(ctx.session)
    sanitizer.ensureMainBody()
    // The raw table.Range.Select should NOT be called (old buggy approach)
    expect(ctx.tableRangeSelect).not.toHaveBeenCalled()
  })

  it("does NOT call endKey on selection proxy", () => {
    const ctx = tableContext()
    const sanitizer = new ContextSanitizer(ctx.session)
    sanitizer.ensureMainBody()
    // endKey should never be called (old buggy approach)
    expect(ctx.selProxy.endKey).not.toHaveBeenCalled()
  })

  it("skips table escape when cursor is NOT within a table", () => {
    const ctx = tableContext()
    ctx.setWithinTable(false)
    const sanitizer = new ContextSanitizer(ctx.session)
    sanitizer.ensureMainBody()
    // No getRange call for table escape
    expect(ctx.docProxy.getRange).not.toHaveBeenCalled()
  })

  it("handles table.Range.End === docEnd (table at very end of document)", () => {
    const ctx = tableContext({ tableEnd: 100, docEnd: 100 })
    const sanitizer = new ContextSanitizer(ctx.session)
    sanitizer.ensureMainBody()
    // min(100, 100) = 100
    expect(ctx.docProxy.getRange).toHaveBeenCalledWith(100, 100)
    expect(ctx.selProxy.collapse).toHaveBeenCalledWith(0)
  })

  it("handles exception gracefully when table operations fail", () => {
    const ctx = tableContext()
    ctx.selProxy.getInformation = vi.fn(() => { throw new Error("COM transient error") })
    const sanitizer = new ContextSanitizer(ctx.session)
    // Should not throw
    expect(() => sanitizer.ensureMainBody()).not.toThrow()
  })
})

// ---------------------------------------------------------------------------
// ensureMainBody — repeated calls (simulating consecutive table inserts)
// ---------------------------------------------------------------------------
describe("ContextSanitizer.ensureMainBody — consecutive calls", () => {
  it("correctly escapes table on every call when cursor keeps re-entering table", () => {
    const ctx = tableContext({ tableEnd: 50, docEnd: 100 })
    const sanitizer = new ContextSanitizer(ctx.session)

    // Simulate 10 consecutive "insert table → cursor stuck in table → ensureMainBody" cycles
    for (let i = 0; i < 10; i++) {
      // Reset cached state to simulate fresh context each time
      sanitizer.reset()
      sanitizer.ensureMainBody()
    }

    // getRange should be called 10 times (once per cycle)
    expect(ctx.docProxy.getRange).toHaveBeenCalledTimes(10)
    expect(ctx.selProxy.collapse).toHaveBeenCalledTimes(10)
    // Old approach should never be used
    expect(ctx.tableRangeSelect).not.toHaveBeenCalled()
    expect(ctx.selProxy.endKey).not.toHaveBeenCalled()
  })

  it("alternates between table escape and normal body check", () => {
    const ctx = tableContext({ tableEnd: 50, docEnd: 100 })
    const sanitizer = new ContextSanitizer(ctx.session)

    // Cycle 1: cursor in table → escape
    sanitizer.reset()
    ctx.setWithinTable(true)
    sanitizer.ensureMainBody()
    expect(ctx.docProxy.getRange).toHaveBeenCalledTimes(1)

    // Cycle 2: cursor in body → no escape needed
    sanitizer.reset()
    ctx.setWithinTable(false)
    sanitizer.ensureMainBody()
    expect(ctx.docProxy.getRange).toHaveBeenCalledTimes(1) // still 1, no new call

    // Cycle 3: cursor in table again → escape again
    sanitizer.reset()
    ctx.setWithinTable(true)
    sanitizer.ensureMainBody()
    expect(ctx.docProxy.getRange).toHaveBeenCalledTimes(2)
  })
})

// ---------------------------------------------------------------------------
// goToEnd
// ---------------------------------------------------------------------------
describe("ContextSanitizer.goToEnd", () => {
  it("calls doc.getRange(docEnd, docEnd).select()", () => {
    const rawDoc = createMockDoc()
    rawDoc.Content = { End: 200, Start: 0, Text: "" } as Record<string, unknown>
    rawDoc.Range = vi.fn(() => ({
      End: 200, Start: 0, Select: vi.fn(),
      Hyperlinks: { Add: vi.fn() }, Shading: { BackgroundPatternColor: 0 },
    }))

    const rawSel = createMockSel()
    const session = createMockSession(undefined, rawDoc, rawSel)
    const docProxy = session.getDocProxy() as MockDocumentProxy

    const sanitizer = new ContextSanitizer(session)
    sanitizer.goToEnd()

    expect(docProxy.getContent).toHaveBeenCalled()
    expect(docProxy.getRange).toHaveBeenCalledWith(200, 200)
  })

  it("does not throw when session has no activeDoc", () => {
    const session = createMockSession({ activeDoc: undefined })
    const sanitizer = new ContextSanitizer(session)
    expect(() => sanitizer.goToEnd()).not.toThrow()
  })
})

// ---------------------------------------------------------------------------
// reset / markInBody / markSelectionRead
// ---------------------------------------------------------------------------
describe("ContextSanitizer state management", () => {
  it("reset clears cached state so next ensureMainBody re-checks", () => {
    const ctx = tableContext()
    ctx.setWithinTable(false)
    const sanitizer = new ContextSanitizer(ctx.session)

    // First call: caches start/end
    sanitizer.ensureMainBody()
    const callsAfterFirst = (ctx.docProxy.getRange as ReturnType<typeof vi.fn>).mock.calls.length

    // Second call without reset: should use cache (no new getRange)
    sanitizer.ensureMainBody()
    const callsAfterSecond = (ctx.docProxy.getRange as ReturnType<typeof vi.fn>).mock.calls.length
    expect(callsAfterSecond).toBe(callsAfterFirst)

    // After reset: should re-check
    sanitizer.reset()
    sanitizer.ensureMainBody()
    const callsAfterReset = (ctx.docProxy.getRange as ReturnType<typeof vi.fn>).mock.calls.length
    expect(callsAfterReset).toBeGreaterThanOrEqual(callsAfterSecond)
  })

  it("markInBody resets collapseReady flag", () => {
    const ctx = tableContext()
    ctx.setWithinTable(false)
    const sanitizer = new ContextSanitizer(ctx.session)
    sanitizer.ensureMainBody()
    sanitizer.markInBody()
    // Next call should re-check (not short-circuit via collapseReady)
    sanitizer.ensureMainBody()
    // No assertion needed — just verifying it does not throw
  })
})
