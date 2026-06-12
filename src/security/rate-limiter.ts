import { WordMcpError } from "./errors.js"

function envNum(key: string, fallback: number): number {
  const val = process.env[key]
  return val ? parseInt(val, 10) : fallback
}

export interface RateLimitConfig {
  windowMs: number
  maxCalls: number
}

const DEFAULT_CONFIG: RateLimitConfig = {
  windowMs: envNum("RATE_LIMIT_WINDOW_MS", 5000),
  maxCalls: envNum("RATE_LIMIT_MAX_CALLS", 30),
}

const CLEANUP_INTERVAL_MS = 60_000

export class RateLimiter {
  private windowMs: number
  private maxCalls: number
  private calls: Map<string, number[]> = new Map()
  private _cleanupTimer: ReturnType<typeof setInterval> | null = null

  constructor(config?: Partial<RateLimitConfig>) {
    this.windowMs = config?.windowMs ?? DEFAULT_CONFIG.windowMs
    this.maxCalls = config?.maxCalls ?? DEFAULT_CONFIG.maxCalls
    this._cleanupTimer = setInterval(() => this._cleanup(), CLEANUP_INTERVAL_MS)
    this._cleanupTimer.unref()
  }

  private _cleanup(): void {
    const now = Date.now()
    for (const [key, timestamps] of this.calls) {
      if (timestamps.every(t => t + this.windowMs <= now)) {
        this.calls.delete(key)
      }
    }
  }

  check(key: string): void {
    const now = Date.now()
    const timestamps = this.calls.get(key) || []
    const windowStart = now - this.windowMs
    const recent = timestamps.filter((t) => t > windowStart)
    if (recent.length >= this.maxCalls) {
      throw new WordMcpError(
        `Rate limit exceeded for "${key}": ${this.maxCalls} calls per ${this.windowMs / 1000}s`,
        "RATE_LIMIT",
        true,
        "Wait a moment before making more requests. Consider batching operations into fewer calls."
      )
    }
    recent.push(now)
    this.calls.set(key, recent)
  }
}
