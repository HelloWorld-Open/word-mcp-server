import { WordMcpError } from "../security/errors.js"

export class BreakerOpenError extends WordMcpError {
  constructor(retryAfterMs: number) {
    const seconds = Math.ceil(retryAfterMs / 1000)
    super(
      `Circuit breaker is OPEN — Word COM temporarily blocked after repeated failures`,
      "CIRCUIT_OPEN",
      true,
      `Automatic retry in ~${seconds}s. Use word_get_status to check current state, or wait for cooldown.`,
    )
    this.name = "BreakerOpenError"
  }
}

export type BreakerState = "CLOSED" | "OPEN" | "HALF_OPEN"

export class CircuitBreaker {
  private state: BreakerState = "CLOSED"
  private failureCount = 0
  private openedAt = 0
  readonly threshold: number
  readonly cooldownMs: number

  constructor(opts?: { threshold?: number; cooldownMs?: number }) {
    this.threshold = opts?.threshold ?? 3
    this.cooldownMs = opts?.cooldownMs ?? 15000
  }

  getState(): BreakerState {
    return this.state
  }

  getFailureCount(): number {
    return this.failureCount
  }

  isOpen(): boolean {
    if (this.state === "OPEN" && Date.now() - this.openedAt >= this.cooldownMs) {
      this.state = "HALF_OPEN"
    }
    return this.state === "OPEN"
  }

  check(): void {
    if (this.isOpen()) {
      const remaining = this.openedAt + this.cooldownMs - Date.now()
      throw new BreakerOpenError(Math.max(remaining, 0))
    }
  }

  onSuccess(): void {
    if (this.state === "HALF_OPEN") {
      this.state = "CLOSED"
    }
    this.failureCount = 0
  }

  onFailure(): void {
    this.failureCount++
    if (this.failureCount >= this.threshold && this.state !== "OPEN") {
      this.state = "OPEN"
      this.openedAt = Date.now()
    }
  }

  forceReset(): void {
    this.state = "CLOSED"
    this.failureCount = 0
    this.openedAt = 0
  }
}
