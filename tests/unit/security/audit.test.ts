import { describe, it, expect, vi } from "vitest"
import { logAudit } from "../../../src/security/audit.js"

describe("audit", () => {
  it("should log OK entry", () => {
    const spy = vi.spyOn(console, "error").mockImplementation(() => {})
    logAudit({ tool: "word_stream_block", durationMs: 42 })

    expect(spy).toHaveBeenCalledOnce()
    const msg = spy.mock.calls[0][0] as string
    expect(msg).toContain("[audit]")
    expect(msg).toContain("OK")
    expect(msg).toContain("word_stream_block")
    expect(msg).toContain("42ms")

    spy.mockRestore()
  })

  it("should log REJECTED entry", () => {
    const spy = vi.spyOn(console, "error").mockImplementation(() => {})
    logAudit({ tool: "word_open", durationMs: 5, error: true })

    const msg = spy.mock.calls[0][0] as string
    expect(msg).toContain("REJECTED")
    expect(msg).toContain("word_open")
    expect(msg).toContain("5ms")

    spy.mockRestore()
  })

  it("should include sanitized args when provided", () => {
    const spy = vi.spyOn(console, "error").mockImplementation(() => {})
    logAudit({ tool: "word_stream_block", durationMs: 10, args: { text: "hello", size: 5 } })

    const msg = spy.mock.calls[0][0] as string
    expect(msg).toContain("hello")
    expect(msg).toContain("5")

    spy.mockRestore()
  })

  it("should truncate long string args", () => {
    const spy = vi.spyOn(console, "error").mockImplementation(() => {})
    const longText = "x".repeat(200)
    logAudit({ tool: "test", durationMs: 0, args: { text: longText } })

    const msg = spy.mock.calls[0][0] as string
    expect(msg).toContain("x".repeat(100) + "...")

    spy.mockRestore()
  })
})
