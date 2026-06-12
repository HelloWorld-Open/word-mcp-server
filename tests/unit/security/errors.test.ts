import { describe, it, expect } from "vitest"
import {
  WordMcpError,
  WordEngineTimeoutError,
  PathSecurityError,
  toMcpContent,
  sanitizeErrorMessage,
} from "../../../src/security/errors.js"

describe("Error types", () => {
  it("WordMcpError should have code and recoverable", () => {
    const err = new WordMcpError("test", "TEST_CODE", true)
    expect(err.code).toBe("TEST_CODE")
    expect(err.recoverable).toBe(true)
    expect(err.message).toBe("test")
  })

  it("WordEngineTimeoutError should mention the method", () => {
    const err = new WordEngineTimeoutError("typeText")
    expect(err.code).toBe("ENGINE_TIMEOUT")
    expect(err.message).toContain("typeText")
  })

  it("PathSecurityError should not be recoverable", () => {
    const err = new PathSecurityError("Blocked path")
    expect(err.code).toBe("PATH_SECURITY")
    expect(err.recoverable).toBe(false)
  })

  it("toMcpContent should format error as text content", () => {
    const err = new WordMcpError("Something broke", "ERR", true, "Try again.")
    const content = toMcpContent(err)
    expect(content).toHaveLength(1)
    expect(content[0].type).toBe("text")
    expect(content[0].text).toContain("[ERR]")
    expect(content[0].text).toContain(">> Recovery: Try again.")
  })

  it("sanitizeErrorMessage should replace paths with placeholder", () => {
    const result = sanitizeErrorMessage(new Error("Access denied: C:\\Users\\test\\file.docx"))
    expect(result).toContain("[path]")
    expect(result).not.toContain("file.docx")
  })

  it("sanitizeErrorMessage should handle non-Error input", () => {
    expect(sanitizeErrorMessage("simple string")).toBe("simple string")
  })
})
