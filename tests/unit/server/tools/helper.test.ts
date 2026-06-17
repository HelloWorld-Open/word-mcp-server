import { describe, it, expect, vi, beforeEach, afterEach } from "vitest"
import { ensureReady, mcpCall } from "../../../../src/server/tools/helper.js"
import { WordEngineTimeoutError, WordMcpError, ServerNotReadyError } from "../../../../src/security/errors.js"
import { ComError } from "../../../../src/word/com-errors.js"

function setupFull() {
  const security = { checkRateLimit: vi.fn() } as any
  const director = {
    precheck: vi.fn().mockResolvedValue({ ok: true }),
    circuitBreaker: { onSuccess: vi.fn(), onFailure: vi.fn(), forceReset: vi.fn() },
    markHealthy: vi.fn(),
    markDirtyIfNeeded: vi.fn(),
    schedulePositionRefresh: vi.fn(),
    recoverSession: vi.fn().mockResolvedValue(undefined),
    captureStatusSuffix: vi.fn().mockReturnValue(""),
    captureContextSuffix: vi.fn().mockReturnValue(""),
  } as any
  const session = {} as any
  const positionMap = {} as any
  const context: ServerContext = { session, positionMap, director }
  return { security, director, session, positionMap, context }
}

describe("ensureReady", () => {
  it("throws ServerNotReadyError when session is null", () => {
    expect(() => ensureReady({ session: null, positionMap: {} as any, director: {} as any })).toThrow(ServerNotReadyError)
  })
  it("throws ServerNotReadyError when positionMap is null", () => {
    expect(() => ensureReady({ session: {} as any, positionMap: null, director: {} as any })).toThrow(ServerNotReadyError)
  })
  it("throws ServerNotReadyError when director is null", () => {
    expect(() => ensureReady({ session: {} as any, positionMap: {} as any, director: null })).toThrow(ServerNotReadyError)
  })
  it("does not throw when all fields are present", () => {
    expect(() => ensureReady({ session: {} as any, positionMap: {} as any, director: {} as any })).not.toThrow()
  })
})

describe("mcpCall", () => {
  describe("happy path", () => {
    it("calls handler and returns formatted content", async () => {
      const { security, context, director } = setupFull()
      const handler = vi.fn().mockResolvedValue("operation succeeded")

      const wrapped = mcpCall(security, context, "test_tool", handler)
      const result = await wrapped({ foo: "bar" })

      expect(security.checkRateLimit).toHaveBeenCalledWith("test_tool")
      expect(director.precheck).toHaveBeenCalledWith("test_tool", undefined)
      expect(handler).toHaveBeenCalledWith({ foo: "bar" })
      expect(director.circuitBreaker.onSuccess).toHaveBeenCalled()
      expect(director.markHealthy).toHaveBeenCalled()
      expect(director.markDirtyIfNeeded).toHaveBeenCalledWith("test_tool")
      expect(director.schedulePositionRefresh).toHaveBeenCalled()
      expect(director.captureStatusSuffix).toHaveBeenCalled()
      expect(result).toEqual({ content: [{ type: "text", text: "operation succeeded" }] })
    })

    it("appends status suffix to result", async () => {
      const { security, context, director } = setupFull()
      director.captureStatusSuffix = vi.fn().mockReturnValue("\n---\ndoc: \"test.docx\"")
      const handler = vi.fn().mockResolvedValue("ok")

      const wrapped = mcpCall(security, context, "test_tool", handler)
      const result = await wrapped({})

      expect(result.content[0].text).toBe("ok\n---\ndoc: \"test.docx\"")
    })
  })

  describe("precheck failure", () => {
    it("returns error content when precheck fails", async () => {
      const { security, context, director } = setupFull()
      director.precheck = vi.fn().mockResolvedValue({ ok: false, error: "[STREAMING] Cannot do that" })
      director.captureStatusSuffix = vi.fn().mockReturnValue(" [suffix]")

      const handler = vi.fn()
      const wrapped = mcpCall(security, context, "test_tool", handler)
      const result = await wrapped({})

      expect(handler).not.toHaveBeenCalled()
      expect(result.content[0].text).toBe("[STREAMING] Cannot do that [suffix]")
    })
  })

  describe("timeout", () => {
    beforeEach(() => { vi.useFakeTimers() })
    afterEach(() => { vi.useRealTimers() })

    it("returns timeout error when handler takes too long", async () => {
      const { security, context, director } = setupFull()
      context.session = { setBusy: vi.fn() } as any
      const neverEnding = vi.fn().mockReturnValue(new Promise(() => {}))

      // Suppress the unused timeoutPromise unhandled rejection
      const rejectionHandler = vi.fn()
      process.on("unhandledRejection", rejectionHandler)

      const wrapped = mcpCall(security, context, "slow_tool", neverEnding, { timeoutMs: 5000 })
      const promise = wrapped({})

      await vi.advanceTimersByTimeAsync(5000)

      const result = await promise
      expect(result.content[0].text).toContain("timed out")
      expect(neverEnding).toHaveBeenCalled()
      expect(director.circuitBreaker.onFailure).toHaveBeenCalled()
      expect(director.recoverSession).toHaveBeenCalled()

      process.off("unhandledRejection", rejectionHandler)
    })
  })

  describe("engine errors", () => {
    it("handles engine error for non-read-only tools", async () => {
      const { security, context, director } = setupFull()
      const handler = vi.fn().mockRejectedValue(new ComError("COM crashed", true))

      const wrapped = mcpCall(security, context, "word_insert_table", handler)
      const result = await wrapped({})

      expect(director.circuitBreaker.onFailure).toHaveBeenCalled()
      expect(director.recoverSession).toHaveBeenCalled()
      expect(result.content[0].text).toContain("[COM_TRANSIENT]")
    })

    it("auto-retries read-only tools on engine error", async () => {
      const { security, context, director } = setupFull()
      const handler = vi.fn()
        .mockRejectedValueOnce(new ComError("transient", true))
        .mockResolvedValueOnce("retry success")

      const wrapped = mcpCall(security, context, "word_get_text", handler)
      const result = await wrapped({})

      expect(handler).toHaveBeenCalledTimes(2)
      expect(director.circuitBreaker.onSuccess).toHaveBeenCalled()
      expect(result.content[0].text).toContain("retry success")
    })

    it("surfaces error if retry also fails", async () => {
      const { security, context, director } = setupFull()
      const handler = vi.fn().mockRejectedValue(new ComError("persistent", true))

      const wrapped = mcpCall(security, context, "word_get_text", handler)
      const result = await wrapped({})

      expect(handler).toHaveBeenCalledTimes(2)
      expect(result.content[0].text).toContain("[COM_TRANSIENT]")
    })

    it("sets session busy on timeout errors", async () => {
      const { security, context, director } = setupFull()
      const session = { setBusy: vi.fn() } as any
      context.session = session
      const handler = vi.fn().mockRejectedValue(new WordEngineTimeoutError("test"))

      const wrapped = mcpCall(security, context, "test_tool", handler)
      const result = await wrapped({})

      expect(session.setBusy).toHaveBeenCalledWith(true)
      expect(result.content[0].text).toContain("timed out")
    })
  })

  describe("WordMcpError handling", () => {
    it("formats WordMcpError with code and hint", async () => {
      const { security, context } = setupFull()
      const handler = vi.fn().mockRejectedValue(
        new WordMcpError("Access denied", "ACCESS_DENIED", false, "Check permissions"),
      )

      const wrapped = mcpCall(security, context, "test_tool", handler)
      const result = await wrapped({})

      const text = result.content[0].text
      expect(text).toContain("[ACCESS_DENIED]")
      expect(text).toContain("Access denied")
      expect(text).toContain("Check permissions")
    })

    it("marks healthy for non-engine errors", async () => {
      const { security, context, director } = setupFull()
      const handler = vi.fn().mockRejectedValue(
        new WordMcpError("User error", "USER_ERR", false),
      )

      const wrapped = mcpCall(security, context, "test_tool", handler)
      const result = await wrapped({})

      expect(director.markHealthy).toHaveBeenCalled()
      expect(result.content[0].text).toContain("[USER_ERR]")
    })
  })

  describe("ComError handling", () => {
    it("formats recoverable ComError with COM_TRANSIENT tag", async () => {
      const { security, context } = setupFull()
      const handler = vi.fn().mockRejectedValue(new ComError("Network hiccup", true))

      const wrapped = mcpCall(security, context, "test_tool", handler)
      const result = await wrapped({})

      expect(result.content[0].text).toContain("[COM_TRANSIENT]")
      expect(result.content[0].text).toContain("Network hiccup")
      expect(result.content[0].text).toContain("retried automatically")
    })

    it("formats fatal ComError with COM_FATAL tag", async () => {
      const { security, context } = setupFull()
      const handler = vi.fn().mockRejectedValue(new ComError("Connection lost", false))

      const wrapped = mcpCall(security, context, "test_tool", handler)
      const result = await wrapped({})

      expect(result.content[0].text).toContain("[COM_FATAL]")
      expect(result.content[0].text).toContain("restart the document session")
    })
  })

  describe("unknown error handling", () => {
    it("sanitizes generic errors", async () => {
      const { security, context } = setupFull()
      const handler = vi.fn().mockRejectedValue(new Error("something broke"))

      const wrapped = mcpCall(security, context, "test_tool", handler)
      const result = await wrapped({})

      const text = result.content[0].text
      expect(text).not.toContain("Error:")
      expect(text.length).toBeGreaterThan(0)
      expect(text.length).toBeLessThanOrEqual(200 + 20)
    })

    it("handles non-Error thrown values", async () => {
      const { security, context } = setupFull()
      const handler = vi.fn().mockRejectedValue("just a string")

      const wrapped = mcpCall(security, context, "test_tool", handler)
      const result = await wrapped({})

      expect(result.content[0].text).toBe("just a string")
    })
  })

  describe("preconditions", () => {
    it("passes preconditions to director.precheck", async () => {
      const { security, context, director } = setupFull()
      const handler = vi.fn().mockResolvedValue("ok")

      const wrapped = mcpCall(security, context, "test_tool", handler, {
        preconditions: ["DOC"],
      })
      await wrapped({})

      expect(director.precheck).toHaveBeenCalledWith("test_tool", ["DOC"])
    })
  })

  describe("timeoutMs = 0", () => {
    it("disables timeout when timeoutMs is 0", async () => {
      const { security, context } = setupFull()
      const handler = vi.fn().mockResolvedValue("no timeout")

      const wrapped = mcpCall(security, context, "test_tool", handler, { timeoutMs: 0 })
      const result = await wrapped({})

      expect(handler).toHaveBeenCalled()
      expect(result.content[0].text).toBe("no timeout")
    })
  })
})
