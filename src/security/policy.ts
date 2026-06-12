import { PathSanitizer } from "./path-sanitizer.js"
import { RateLimiter, type RateLimitConfig } from "./rate-limiter.js"
import { WordMcpError } from "./errors.js"

function envNum(key: string, fallback: number): number {
  const val = process.env[key]
  return val ? parseInt(val, 10) : fallback
}

interface SecurityPolicy {
  allowedDirectories: string[]
  allowNetworkPaths: boolean
  maxFileSize: number
  maxTextLength: number
  allowMacros: boolean
  operationTimeoutMs: number
}

const DEFAULT_SECURITY_POLICY: SecurityPolicy = {
  allowedDirectories: process.env.ALLOWED_DIRECTORIES
    ? process.env.ALLOWED_DIRECTORIES.split(";").filter(Boolean)
    : [],
  allowNetworkPaths: process.env.ALLOW_NETWORK_PATHS === "true",
  maxFileSize: envNum("MAX_FILE_SIZE", 50 * 1024 * 1024),
  maxTextLength: envNum("MAX_TEXT_LENGTH", 1000000),
  allowMacros: process.env.ALLOW_MACROS === "true",
  operationTimeoutMs: envNum("OPERATION_TIMEOUT_MS", 30000),
}

export class SecurityManager {
  public readonly policy: SecurityPolicy
  public readonly pathSanitizer: PathSanitizer
  public readonly rateLimiter: RateLimiter

  constructor(policy?: Partial<SecurityPolicy>, rateLimitConfig?: Partial<RateLimitConfig>) {
    this.policy = { ...DEFAULT_SECURITY_POLICY, ...policy }
    this.pathSanitizer = new PathSanitizer({
      allowedDirectories: this.policy.allowedDirectories,
      allowNetworkPaths: this.policy.allowNetworkPaths,
      maxFileSize: this.policy.maxFileSize,
    })
    this.rateLimiter = new RateLimiter(rateLimitConfig)
  }

  checkRateLimit(toolName: string): void {
    this.rateLimiter.check(toolName)
  }

  validateTextLength(text: string): void {
    if (text.length > this.policy.maxTextLength) {
      throw new WordMcpError(
        `Text exceeds maximum length of ${this.policy.maxTextLength} characters`,
        "TEXT_TOO_LONG",
        false,
        "Reduce text length and split into shorter segments or multiple calls."
      )
    }
  }

}
