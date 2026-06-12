import { describe, it, expect, vi } from "vitest"
import { WordSession } from "../../../src/word/session.js"

function createMockWinax() {
  const mockApp: Record<string, unknown> = {
    Version: "16.0",
    Visible: false,
    DisplayAlerts: 999,
    AutomationSecurity: 999,
    Documents: { Count: 0, Item: vi.fn(), Add: vi.fn() },
    Quit: vi.fn(),
    Release: vi.fn(),
  }
  const MockObject = vi.fn(function (this: unknown, progid: string) {
    return mockApp
  })
  const mockWinax = { Object: MockObject as unknown as new (progid: string) => Record<string, unknown> }
  return { mockWinax, mockApp, MockObject }
}

describe("WordSession", () => {
  it("initial state is not running, no active doc", () => {
    const { mockWinax } = createMockWinax()
    const session = new WordSession(() => mockWinax)
    expect(session.isAlive()).toBe(false)
    expect(session.activeDoc).toBeNull()
  })

  it("start creates Word.Application via winax", () => {
    const { mockWinax, mockApp } = createMockWinax()
    const session = new WordSession(() => mockWinax)
    session.start()
    expect(mockWinax.Object).toHaveBeenCalledWith("Word.Application")
    expect(mockApp.Visible).toBe(true)
    expect(mockApp.DisplayAlerts).toBe(0)
    expect(mockApp.AutomationSecurity).toBe(3)
    expect(session.isAlive()).toBe(true)
  })

  it("start is idempotent", () => {
    const { mockWinax } = createMockWinax()
    const session = new WordSession(() => mockWinax)
    session.start()
    session.start()
    expect(mockWinax.Object).toHaveBeenCalledTimes(1)
  })

  it("ensureAlive starts if not running", () => {
    const { mockWinax } = createMockWinax()
    const session = new WordSession(() => mockWinax)
    session.ensureAlive()
    expect(session.isAlive()).toBe(true)
  })

  it("application lazy-starts", () => {
    const { mockWinax } = createMockWinax()
    const session = new WordSession(() => mockWinax)
    const app = session.application
    expect(app).not.toBeNull()
    expect(session.isAlive()).toBe(true)
  })

  it("application getter throws when unhealthy", () => {
    const { mockWinax } = createMockWinax()
    const session = new WordSession(() => mockWinax)
    session.markUnhealthy()
    expect(() => session.application).toThrow("unhealthy")
  })

  it("setActiveDoc stores and activeDoc returns it", () => {
    const { mockWinax } = createMockWinax()
    const session = new WordSession(() => mockWinax)
    const fakeDoc = { Name: "test.docx" }
    session.setActiveDoc(fakeDoc as unknown as Record<string, unknown>)
    expect(session.activeDoc).toBe(fakeDoc)
  })

  it("quit clears state when no real Word process", () => {
    const { mockWinax } = createMockWinax()
    const session = new WordSession(() => mockWinax)
    session.start()
    session.quit()
    expect(session.isAlive()).toBe(false)
    expect(session.activeDoc).toBeNull()
  })

  it("quit is safe when not started", () => {
    const session = new WordSession(() => ({ Object: vi.fn() }))
    expect(() => session.quit()).not.toThrow()
  })

  it("comCall throws on COM failure", () => {
    const { mockWinax } = createMockWinax()
    const session = new WordSession(() => mockWinax)
    session.start()
    expect(() => session.comCall(() => { throw new Error("COM error") })).toThrow("COM error")
  })

  it("healthCheck returns false when no Word process", () => {
    const { mockWinax } = createMockWinax()
    const session = new WordSession(() => mockWinax)
    session.start()
    expect(session.healthCheck()).toBe(false)
  })

  it("setOnLog stores handler", () => {
    const { mockWinax } = createMockWinax()
    const session = new WordSession(() => mockWinax)
    const handler = vi.fn()
    session.setOnLog(handler)
    session.start()
    expect(handler).toHaveBeenCalledWith("info", expect.stringContaining("Creating Word.Application"))
  })

  it("recover() restarts session asynchronously", async () => {
    const { mockWinax } = createMockWinax()
    const session = new WordSession(() => mockWinax)
    session.start()
    await session.recover()
    expect(session.isAlive()).toBe(true)
  })
})
