import { describe, it, expect, vi } from "vitest"
import { WordFormatter } from "../../../src/word/word-formatter.js"
import { createMockSession, createMockDoc, createMockSel } from "../test-helpers.js"
import { createMockDocProxy, createMockRangeProxy } from "../../../src/word/com-proxy/com-proxy.mock.js"
import type { IHeaderFooterProxy, ISectionProxy, ISelectionProxy, IDocumentProxy } from "../../../src/word/com-proxy/types.js"

function setupHeaderFooterTest() {
  const rawDoc = createMockDoc()
  const rawSel = createMockSel()

  const hdrRangeMock = createMockRangeProxy()
  const ftrRangeMock = createMockRangeProxy()

  const hdrMock: IHeaderFooterProxy = {
    select: vi.fn(),
    clearContent: vi.fn(),
    typeText: vi.fn(),
    setAlignment: vi.fn(),
    setContent: vi.fn(),
    getEnd: vi.fn(() => 100),
    getText: vi.fn(() => ""),
    getRange: vi.fn(() => hdrRangeMock),
    getFields: vi.fn(() => ({ count: 0, item: vi.fn() })),
    getPageNumbersCount: vi.fn(() => 0),
    setPageNumbersAlignment: vi.fn(),
    raw: {},
  }

  const ftrMock: IHeaderFooterProxy = {
    select: vi.fn(),
    clearContent: vi.fn(),
    typeText: vi.fn(),
    setAlignment: vi.fn(),
    setContent: vi.fn(),
    getEnd: vi.fn(() => 150),
    getText: vi.fn(() => ""),
    getRange: vi.fn(() => ftrRangeMock),
    getFields: vi.fn(() => ({ count: 0, item: vi.fn() })),
    getPageNumbersCount: vi.fn(() => 0),
    setPageNumbersAlignment: vi.fn(),
    raw: {},
  }

  const sectionMock: ISectionProxy = {
    getHeader: vi.fn(() => hdrMock),
    getFooter: vi.fn(() => ftrMock),
    getPageSetup: vi.fn(() => ({})),
    raw: {},
  }

  const sectionsMock = { count: 1, item: vi.fn(() => sectionMock) }

  const docProxy = createMockDocProxy(rawDoc)
  vi.mocked(docProxy.getSections).mockReturnValue(sectionsMock)
  vi.mocked(docProxy.getRange).mockImplementation((_s, _e) => createMockRangeProxy())

  const session = createMockSession(
    {
      getDocProxy: () => docProxy,
      lockPrintView: vi.fn(),
    },
    rawDoc,
    rawSel,
  )
  const fmt = new WordFormatter(session)
  const selProxy = session.getSelectionProxy()

  return { fmt, rawDoc, rawSel, session, hdrMock, ftrMock, sectionMock, sectionsMock, docProxy, selProxy, hdrRangeMock, ftrRangeMock }
}

// ─── setHeader ───────────────────────────────────────────────────────────────

describe("WordFormatter.setHeader", () => {
  it("delegates to section.getHeader(1).setContent", async () => {
    const { fmt, hdrMock, sectionMock } = setupHeaderFooterTest()
    await fmt.setHeader("My Header")
    expect(sectionMock.getHeader).toHaveBeenCalledWith(1)
    expect(hdrMock.setContent).toHaveBeenCalledWith("My Header", undefined)
  })

  it("maps alignment strings correctly", async () => {
    const { fmt, hdrMock } = setupHeaderFooterTest()
    await fmt.setHeader("Left", "left")
    expect(hdrMock.setContent).toHaveBeenCalledWith("Left", 0)
    await fmt.setHeader("Center", "center")
    expect(hdrMock.setContent).toHaveBeenCalledWith("Center", 1)
    await fmt.setHeader("Right", "right")
    expect(hdrMock.setContent).toHaveBeenCalledWith("Right", 2)
  })

  it("saves cursor before and restores after", async () => {
    const { fmt, session, selProxy } = setupHeaderFooterTest()
    await fmt.setHeader("Test")
    expect(selProxy.getStart).toHaveBeenCalled()
    expect(session.lockPrintView).toHaveBeenCalled()
  })

  it("operates on last section when multiple sections exist", async () => {
    const { fmt, sectionMock, sectionsMock } = setupHeaderFooterTest()
    sectionsMock.count = 3
    await fmt.setHeader("Sec3")
    expect(sectionsMock.item).toHaveBeenCalledWith(3)
    expect(sectionMock.getHeader).toHaveBeenCalledWith(1)
  })

  it("handles empty text gracefully", async () => {
    const { fmt, hdrMock } = setupHeaderFooterTest()
    await expect(fmt.setHeader("")).resolves.toBeUndefined()
    expect(hdrMock.setContent).toHaveBeenCalledWith("", undefined)
  })
})

// ─── setFooter ───────────────────────────────────────────────────────────────

describe("WordFormatter.setFooter", () => {
  it("delegates to section.getFooter(1).setContent", async () => {
    const { fmt, ftrMock, sectionMock } = setupHeaderFooterTest()
    await fmt.setFooter("My Footer")
    expect(sectionMock.getFooter).toHaveBeenCalledWith(1)
    expect(ftrMock.setContent).toHaveBeenCalledWith("My Footer", undefined)
  })

  it("maps alignment strings", async () => {
    const { fmt, ftrMock } = setupHeaderFooterTest()
    await fmt.setFooter("Left", "left")
    expect(ftrMock.setContent).toHaveBeenCalledWith("Left", 0)
    await fmt.setFooter("Center", "center")
    expect(ftrMock.setContent).toHaveBeenCalledWith("Center", 1)
    await fmt.setFooter("Right", "right")
    expect(ftrMock.setContent).toHaveBeenCalledWith("Right", 2)
  })

  it("saves cursor before and restores after", async () => {
    const { fmt, session, selProxy } = setupHeaderFooterTest()
    await fmt.setFooter("Test")
    expect(selProxy.getStart).toHaveBeenCalled()
    expect(session.lockPrintView).toHaveBeenCalled()
  })

  it("operates on last section", async () => {
    const { fmt, sectionsMock, sectionMock } = setupHeaderFooterTest()
    sectionsMock.count = 2
    await fmt.setFooter("Footer")
    expect(sectionsMock.item).toHaveBeenCalledWith(2)
    expect(sectionMock.getFooter).toHaveBeenCalledWith(1)
  })
})

// ─── setPageNumbers ──────────────────────────────────────────────────────────

describe("WordFormatter.setPageNumbers", () => {
  it("adds PAGE field in header with default center alignment when empty", async () => {
    const { fmt, hdrMock, selProxy, hdrRangeMock } = setupHeaderFooterTest()
    vi.mocked(hdrMock.getText).mockReturnValue("")
    const selRangeMock = createMockRangeProxy()
    vi.mocked(selProxy.getRange).mockReturnValue(selRangeMock)

    await fmt.setPageNumbers("header")

    expect(hdrMock.getFields).toHaveBeenCalled()
    expect(hdrMock.getEnd).toHaveBeenCalled()
    expect(hdrMock.getText).toHaveBeenCalled()
    expect(hdrMock.select).toHaveBeenCalled()
    expect(hdrRangeMock.setRange).toHaveBeenCalledWith(100, 100)
    expect(hdrRangeMock.select).toHaveBeenCalled()
    expect(selProxy.typeText).not.toHaveBeenCalled()
    expect(selRangeMock.addField).toHaveBeenCalledWith(33)
    expect(hdrMock.setAlignment).toHaveBeenCalledWith(1)
    expect(hdrMock.setPageNumbersAlignment).toHaveBeenCalledWith(1)
    expect(selProxy.collapse).toHaveBeenCalledWith(0)
  })

  it("adds PAGE field in footer", async () => {
    const { fmt, ftrMock, selProxy, ftrRangeMock } = setupHeaderFooterTest()
    vi.mocked(ftrMock.getText).mockReturnValue("")
    const selRangeMock = createMockRangeProxy()
    vi.mocked(selProxy.getRange).mockReturnValue(selRangeMock)

    await fmt.setPageNumbers("footer")

    expect(ftrMock.getFields).toHaveBeenCalled()
    expect(ftrRangeMock.setRange).toHaveBeenCalledWith(150, 150)
    expect(ftrRangeMock.select).toHaveBeenCalled()
    expect(selRangeMock.addField).toHaveBeenCalledWith(33)
    expect(ftrMock.setPageNumbersAlignment).toHaveBeenCalledWith(1)
    expect(ftrMock.setAlignment).toHaveBeenCalledWith(1)
  })

  it("maps alignment values", async () => {
    const { fmt, hdrMock } = setupHeaderFooterTest()
    vi.mocked(hdrMock.getText).mockReturnValue("")

    await fmt.setPageNumbers("header", "left")
    expect(hdrMock.setPageNumbersAlignment).toHaveBeenCalledWith(0)
    expect(hdrMock.setAlignment).toHaveBeenCalledWith(0)

    await fmt.setPageNumbers("header", "right")
    expect(hdrMock.setPageNumbersAlignment).toHaveBeenCalledWith(2)
    expect(hdrMock.setAlignment).toHaveBeenCalledWith(2)
  })

  it("deletes existing PAGE fields and preserves other fields", async () => {
    const { fmt, hdrMock } = setupHeaderFooterTest()
    vi.mocked(hdrMock.getText).mockReturnValue("")

    const pageField1 = { type: 33, delete: vi.fn() }
    const tocField = { type: 16, delete: vi.fn() }
    const pageField2 = { type: 33, delete: vi.fn() }

    vi.mocked(hdrMock.getFields).mockReturnValue({
      count: 3,
      item: vi.fn((i: number) => [pageField1, tocField, pageField2][i - 1]),
    })

    await fmt.setPageNumbers("header")

    expect(pageField1.delete).toHaveBeenCalled()
    expect(pageField2.delete).toHaveBeenCalled()
    expect(tocField.delete).not.toHaveBeenCalled()
  })

  it("deletes fields in reverse order to preserve indices", async () => {
    const { fmt, hdrMock } = setupHeaderFooterTest()
    vi.mocked(hdrMock.getText).mockReturnValue("")

    const callOrder: number[] = []
    const pageField1 = { type: 33, delete: vi.fn(() => callOrder.push(1)) }
    const pageField2 = { type: 33, delete: vi.fn(() => callOrder.push(2)) }
    const pageField3 = { type: 33, delete: vi.fn(() => callOrder.push(3)) }

    vi.mocked(hdrMock.getFields).mockReturnValue({
      count: 3,
      item: vi.fn((i: number) => [pageField1, pageField2, pageField3][i - 1]),
    })

    await fmt.setPageNumbers("header")

    expect(callOrder).toEqual([3, 2, 1])
  })

  it("adds space before PAGE field in non-empty header", async () => {
    const { fmt, hdrMock, selProxy, hdrRangeMock } = setupHeaderFooterTest()
    vi.mocked(hdrMock.getText).mockReturnValue("Page ")
    const selRangeMock = createMockRangeProxy()
    vi.mocked(selProxy.getRange).mockReturnValue(selRangeMock)

    await fmt.setPageNumbers("header")

    expect(selProxy.typeText).toHaveBeenCalledWith(" ")
    expect(hdrRangeMock.setRange).toHaveBeenCalledWith(99, 99)
    expect(selRangeMock.addField).toHaveBeenCalledWith(33)
  })

  it("does not add space in empty header", async () => {
    const { fmt, hdrMock, selProxy, hdrRangeMock } = setupHeaderFooterTest()
    vi.mocked(hdrMock.getText).mockReturnValue("")
    const selRangeMock = createMockRangeProxy()
    vi.mocked(selProxy.getRange).mockReturnValue(selRangeMock)

    await fmt.setPageNumbers("header")

    expect(selProxy.typeText).not.toHaveBeenCalled()
    expect(hdrRangeMock.setRange).toHaveBeenCalledWith(100, 100)
    expect(selRangeMock.addField).toHaveBeenCalledWith(33)
  })

  it("handles whitespace-only header as empty", async () => {
    const { fmt, hdrMock, selProxy, hdrRangeMock } = setupHeaderFooterTest()
    vi.mocked(hdrMock.getText).mockReturnValue("  \r\n  ")
    const selRangeMock = createMockRangeProxy()
    vi.mocked(selProxy.getRange).mockReturnValue(selRangeMock)

    await fmt.setPageNumbers("header")

    expect(selProxy.typeText).not.toHaveBeenCalled()
    expect(hdrRangeMock.setRange).toHaveBeenCalledWith(100, 100)
  })

  it("restores cursor and locks print view after operation", async () => {
    const { fmt, session, selProxy, hdrMock } = setupHeaderFooterTest()
    vi.mocked(hdrMock.getText).mockReturnValue("")
    const selRangeMock = createMockRangeProxy()
    vi.mocked(selProxy.getRange).mockReturnValue(selRangeMock)

    await fmt.setPageNumbers("header")

    expect(session.lockPrintView).toHaveBeenCalled()
    expect(selProxy.collapse).toHaveBeenCalledWith(0)
  })

  it("operates on last section", async () => {
    const { fmt, sectionsMock, sectionMock } = setupHeaderFooterTest()
    sectionsMock.count = 4
    await fmt.setPageNumbers("header")
    expect(sectionsMock.item).toHaveBeenCalledWith(4)
    expect(sectionMock.getHeader).toHaveBeenCalledWith(1)
  })

  it("accepts explicit sectionIndex", async () => {
    const { fmt, sectionsMock, sectionMock } = setupHeaderFooterTest()
    sectionsMock.count = 9
    await fmt.setPageNumbers("header", undefined, 3)
    expect(sectionsMock.item).toHaveBeenCalledWith(3)
    expect(sectionMock.getHeader).toHaveBeenCalledWith(1)
  })

  it("throws on out-of-range sectionIndex", async () => {
    const { fmt, sectionsMock } = setupHeaderFooterTest()
    sectionsMock.count = 3
    await expect(fmt.setPageNumbers("header", undefined, 0)).rejects.toThrow()
    await expect(fmt.setPageNumbers("header", undefined, 4)).rejects.toThrow()
  })

  it("allows sectionIndex=1 for first section", async () => {
    const { fmt, sectionsMock, sectionMock } = setupHeaderFooterTest()
    sectionsMock.count = 5
    await fmt.setPageNumbers("header", undefined, 1)
    expect(sectionsMock.item).toHaveBeenCalledWith(1)
    expect(sectionMock.getHeader).toHaveBeenCalledWith(1)
  })
})

// ─── section targeting (shared by all three methods) ─────────────────────

describe("WordFormatter section targeting", () => {
  it("setHeader accepts explicit sectionIndex", async () => {
    const { fmt, sectionsMock, sectionMock } = setupHeaderFooterTest()
    sectionsMock.count = 5
    await fmt.setHeader("H", undefined, 2)
    expect(sectionsMock.item).toHaveBeenCalledWith(2)
    expect(sectionMock.getHeader).toHaveBeenCalledWith(1)
  })

  it("setHeader throws on out-of-range sectionIndex", async () => {
    const { fmt, sectionsMock } = setupHeaderFooterTest()
    sectionsMock.count = 2
    await expect(fmt.setHeader("H", undefined, 3)).rejects.toThrow()
    await expect(fmt.setHeader("H", undefined, 0)).rejects.toThrow()
  })

  it("setFooter accepts explicit sectionIndex", async () => {
    const { fmt, sectionsMock, sectionMock } = setupHeaderFooterTest()
    sectionsMock.count = 4
    await fmt.setFooter("F", undefined, 3)
    expect(sectionsMock.item).toHaveBeenCalledWith(3)
    expect(sectionMock.getFooter).toHaveBeenCalledWith(1)
  })

  it("setFooter throws on out-of-range sectionIndex", async () => {
    const { fmt, sectionsMock } = setupHeaderFooterTest()
    sectionsMock.count = 1
    await expect(fmt.setFooter("F", undefined, 2)).rejects.toThrow()
  })
})
