import { describe, it, expect, beforeEach } from "vitest"
import { RateLimiter } from "../../../src/security/rate-limiter.js"
import { WordMcpError } from "../../../src/security/errors.js"

describe("RateLimiter", () => {
  let limiter: RateLimiter

  beforeEach(() => {
    limiter = new RateLimiter({ windowMs: 100, maxCalls: 3 })
  })

  it("should allow calls within limit", () => {
    expect(() => limiter.check("test")).not.toThrow()
    expect(() => limiter.check("test")).not.toThrow()
    expect(() => limiter.check("test")).not.toThrow()
  })

  it("should reject calls over limit", () => {
    limiter.check("test")
    limiter.check("test")
    limiter.check("test")
    expect(() => limiter.check("test")).toThrow(WordMcpError)
  })

  it("should throw with RATE_LIMIT code", () => {
    limiter.check("test")
    limiter.check("test")
    limiter.check("test")
    try {
      limiter.check("test")
      expect.fail("should have thrown")
    } catch (err) {
      expect(err).toBeInstanceOf(WordMcpError)
      expect((err as WordMcpError).code).toBe("RATE_LIMIT")
      expect((err as WordMcpError).recoverable).toBe(true)
    }
  })

  it("should track keys independently", () => {
    limiter.check("a")
    limiter.check("a")
    limiter.check("a")
    expect(() => limiter.check("a")).toThrow(WordMcpError)
    expect(() => limiter.check("b")).not.toThrow()
  })
})
