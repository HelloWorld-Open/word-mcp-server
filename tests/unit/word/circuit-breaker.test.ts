import { describe, it, expect, vi, beforeEach, afterEach } from "vitest"
import { CircuitBreaker, BreakerOpenError } from "../../../src/word/circuit-breaker.js"

describe("CircuitBreaker", () => {
  beforeEach(() => {
    vi.useFakeTimers()
  })

  afterEach(() => {
    vi.useRealTimers()
  })

  it("starts CLOSED with zero failures", () => {
    const cb = new CircuitBreaker()
    expect(cb.getState()).toBe("CLOSED")
    expect(cb.getFailureCount()).toBe(0)
    expect(cb.isOpen()).toBe(false)
  })

  it("tracks failure count without opening below threshold", () => {
    const cb = new CircuitBreaker({ threshold: 3 })
    cb.onFailure()
    expect(cb.getState()).toBe("CLOSED")
    expect(cb.getFailureCount()).toBe(1)
    cb.onFailure()
    expect(cb.getState()).toBe("CLOSED")
    expect(cb.getFailureCount()).toBe(2)
  })

  it("opens after threshold failures", () => {
    const cb = new CircuitBreaker({ threshold: 3 })
    cb.onFailure()
    cb.onFailure()
    cb.onFailure()
    expect(cb.getState()).toBe("OPEN")
    expect(cb.isOpen()).toBe(true)
  })

  it("transitions to HALF_OPEN after cooldown", () => {
    const cb = new CircuitBreaker({ threshold: 2, cooldownMs: 10000 })
    cb.onFailure()
    cb.onFailure()
    expect(cb.getState()).toBe("OPEN")
    expect(cb.isOpen()).toBe(true)
    vi.advanceTimersByTime(10000)
    expect(cb.isOpen()).toBe(false)
    expect(cb.getState()).toBe("HALF_OPEN")
  })

  it("resets to CLOSED on success during HALF_OPEN", () => {
    const cb = new CircuitBreaker({ threshold: 1, cooldownMs: 5000 })
    cb.onFailure()
    expect(cb.getState()).toBe("OPEN")
    vi.advanceTimersByTime(5000)
    expect(cb.isOpen()).toBe(false)
    expect(cb.getState()).toBe("HALF_OPEN")
    cb.onSuccess()
    expect(cb.getState()).toBe("CLOSED")
    expect(cb.getFailureCount()).toBe(0)
  })

  it("re-opens on failure during HALF_OPEN", () => {
    const cb = new CircuitBreaker({ threshold: 2, cooldownMs: 5000 })
    cb.onFailure()
    cb.onFailure()
    expect(cb.getState()).toBe("OPEN")
    vi.advanceTimersByTime(5000)
    cb.isOpen()
    expect(cb.getState()).toBe("HALF_OPEN")
    cb.onFailure()
    expect(cb.getState()).toBe("OPEN")
  })

  it("check() throws BreakerOpenError when OPEN", () => {
    const cb = new CircuitBreaker({ threshold: 1, cooldownMs: 10000 })
    cb.onFailure()
    expect(() => cb.check()).toThrow(BreakerOpenError)
  })

  it("check() does not throw when CLOSED", () => {
    const cb = new CircuitBreaker()
    expect(() => cb.check()).not.toThrow()
  })

  it("check() does not throw during HALF_OPEN", () => {
    const cb = new CircuitBreaker({ threshold: 1, cooldownMs: 5000 })
    cb.onFailure()
    vi.advanceTimersByTime(5000)
    expect(() => cb.check()).not.toThrow()
  })

  it("forceReset() resets to CLOSED from any state", () => {
    const cb = new CircuitBreaker({ threshold: 1 })
    cb.onFailure()
    expect(cb.getState()).toBe("OPEN")
    cb.forceReset()
    expect(cb.getState()).toBe("CLOSED")
    expect(cb.getFailureCount()).toBe(0)
  })

  it("onSuccess resets failure count in CLOSED state", () => {
    const cb = new CircuitBreaker({ threshold: 5 })
    cb.onFailure()
    cb.onFailure()
    cb.onSuccess()
    expect(cb.getFailureCount()).toBe(0)
    expect(cb.getState()).toBe("CLOSED")
  })

  it("handles default constructor values", () => {
    const cb = new CircuitBreaker()
    expect(cb.threshold).toBe(3)
    expect(cb.cooldownMs).toBe(15000)
  })
})
