import { describe, it, expect } from "vitest"
import { SecurityManager } from "../../../src/security/policy.js"
import { WordMcpError } from "../../../src/security/errors.js"

describe("SecurityManager", () => {
  it("should use default policy values", () => {
    const sm = new SecurityManager()
    expect(sm.policy.maxFileSize).toBe(50 * 1024 * 1024)
    expect(sm.policy.maxTextLength).toBe(1000000)
    expect(sm.policy.allowMacros).toBe(false)
    expect(sm.policy.allowNetworkPaths).toBe(false)
    expect(sm.policy.operationTimeoutMs).toBe(30000)
  })

  it("should merge custom policy", () => {
    const sm = new SecurityManager({ maxFileSize: 1024, allowNetworkPaths: true })
    expect(sm.policy.maxFileSize).toBe(1024)
    expect(sm.policy.allowNetworkPaths).toBe(true)
    expect(sm.policy.maxTextLength).toBe(1000000) // unchanged default
  })

  it("should validate text length", () => {
    const sm = new SecurityManager({ maxTextLength: 10 })
    expect(() => sm.validateTextLength("short")).not.toThrow()
    expect(() => sm.validateTextLength("this is too long")).toThrow(WordMcpError)
  })

  it("should have a path sanitizer instance", () => {
    const sm = new SecurityManager()
    expect(sm.pathSanitizer).toBeDefined()
    expect(typeof sm.pathSanitizer.resolveAndValidate).toBe("function")
  })

  it("should have a rate limiter instance", () => {
    const sm = new SecurityManager()
    expect(sm.rateLimiter).toBeDefined()
    expect(typeof sm.checkRateLimit).toBe("function")
  })

  it("should check rate limit through security manager", () => {
    const sm = new SecurityManager(undefined, { windowMs: 100, maxCalls: 2 })
    expect(() => sm.checkRateLimit("test")).not.toThrow()
    expect(() => sm.checkRateLimit("test")).not.toThrow()
    expect(() => sm.checkRateLimit("test")).toThrow(WordMcpError)
  })

  it("should accept rate limit config", () => {
    const sm = new SecurityManager({}, { windowMs: 5000, maxCalls: 10 })
    expect(sm.rateLimiter).toBeDefined()
  })
})
