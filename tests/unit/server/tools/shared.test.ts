import { describe, it, expect, vi } from "vitest"
import { isEngineError, isReadOnlyTool, createRegTool } from "../../../../src/server/tools/shared.js"
import { WordEngineTimeoutError, WordMcpError } from "../../../../src/security/errors.js"
import { ComError } from "../../../../src/word/com-errors.js"

describe("isEngineError", () => {
  it("returns true for WordEngineTimeoutError", () => {
    expect(isEngineError(new WordEngineTimeoutError("test"))).toBe(true)
  })
  it("returns true for ComError", () => {
    expect(isEngineError(new ComError("COM failed", true))).toBe(true)
    expect(isEngineError(new ComError("COM fatal", false))).toBe(true)
  })
  it("returns false for WordMcpError", () => {
    expect(isEngineError(new WordMcpError("msg", "CODE", false))).toBe(false)
    expect(isEngineError(new WordMcpError("msg", "CODE", true))).toBe(false)
  })
  it("returns false for plain Error", () => {
    expect(isEngineError(new Error("generic"))).toBe(false)
  })
  it("returns false for null and undefined", () => {
    expect(isEngineError(null)).toBe(false)
    expect(isEngineError(undefined)).toBe(false)
  })
  it("returns false for non-Error values", () => {
    expect(isEngineError("string")).toBe(false)
    expect(isEngineError(42)).toBe(false)
    expect(isEngineError({})).toBe(false)
  })
})

describe("isReadOnlyTool", () => {
  const readOnlyTools = [
    "word_get_text", "word_get_paragraph", "word_get_structure", "word_get_info",
    "word_get_status", "word_get_table_data", "word_get_comments", "word_get_bookmarks",
    "word_get_lists", "word_get_sections", "word_get_cursor_info", "word_locate",
    "word_list_styles",
  ]
  for (const name of readOnlyTools) {
    it(`returns true for ${name}`, () => {
      expect(isReadOnlyTool(name)).toBe(true)
    })
  }
  it("returns false for write tools", () => {
    expect(isReadOnlyTool("word_type_text")).toBe(false)
    expect(isReadOnlyTool("word_insert_table")).toBe(false)
    expect(isReadOnlyTool("word_set_page_region")).toBe(false)
  })
  it("returns false for unknown tool names", () => {
    expect(isReadOnlyTool("nonexistent_tool")).toBe(false)
  })
  it("returns false for empty string", () => {
    expect(isReadOnlyTool("")).toBe(false)
  })
})

describe("createRegTool", () => {
  it("calls server.registerTool with name, config, and a function", () => {
    const server = { registerTool: vi.fn() }
    const security = { checkRateLimit: vi.fn() } as any
    const context = { session: null, positionMap: null, director: null } as any
    const regTool = createRegTool(server as any, security, context)

    const handler = vi.fn()
    regTool("my_tool", { description: "My test tool" }, handler)

    expect(server.registerTool).toHaveBeenCalledTimes(1)
    expect(server.registerTool).toHaveBeenCalledWith(
      "my_tool",
      { description: "My test tool" },
      expect.any(Function),
    )
  })

  it("passes options through to the wrapper", () => {
    const server = { registerTool: vi.fn() }
    const security = { checkRateLimit: vi.fn() } as any
    const context = { session: null, positionMap: null, director: null } as any
    const regTool = createRegTool(server as any, security, context)

    const handler = vi.fn()
    regTool("my_tool", { description: "Test" }, handler, { timeoutMs: 10000, preconditions: ["DOC"] })

    expect(server.registerTool).toHaveBeenCalledTimes(1)
    // The registered function is from mcpCall, verify it accepts args
    const registeredFn = server.registerTool.mock.calls[0][2]
    expect(registeredFn).toBeInstanceOf(Function)
  })
})
