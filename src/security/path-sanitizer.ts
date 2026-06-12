import { resolve, normalize, relative, parse } from "node:path"
import { existsSync, statSync } from "node:fs"
import { PathSecurityError } from "./errors.js"

export interface PathPolicy {
  allowedDirectories: string[]
  allowNetworkPaths: boolean
  maxFileSize: number
}

export class PathSanitizer {
  private policy: PathPolicy

  constructor(policy: PathPolicy) {
    this.policy = policy
  }

  resolveAndValidate(rawPath: string): string {
    const resolved = resolve(normalize(rawPath.trim()))

    this.rejectIfTraversal(rawPath, resolved)
    this.rejectIfSpecialDevicePath(rawPath)
    this.rejectIfNetworkPath(resolved)
    this.rejectIfBlocklisted(resolved)

    if (this.policy.allowedDirectories.length > 0) {
      this.rejectIfNotAllowed(resolved)
    }

    return resolved
  }

  validateForRead(rawPath: string): string {
    const resolved = this.resolveAndValidate(rawPath)

    if (!existsSync(resolved)) {
      throw new PathSecurityError(`File does not exist: ${resolved}`)
    }

    const stats = statSync(resolved)
    if (!stats.isFile()) {
      throw new PathSecurityError(`Not a file: ${resolved}`)
    }

    if (stats.size > this.policy.maxFileSize) {
      throw new PathSecurityError(
        `File too large: ${stats.size} bytes (max ${this.policy.maxFileSize} bytes)`
      )
    }

    return resolved
  }

  validateForWrite(rawPath: string): string {
    const resolved = this.resolveAndValidate(rawPath)
    return resolved
  }

  private rejectIfTraversal(rawPath: string, resolved: string): void {
    const cwd = process.cwd()
    if (rawPath.includes("..") || rawPath.startsWith("~") || rawPath.includes("\\~") || rawPath.includes("/~")) {
      const resRoot = parse(resolved).root
      const cwdRoot = parse(cwd).root
      if (resRoot !== cwdRoot) {
        throw new PathSecurityError(`Cross-drive path not allowed: ${resolved}`)
      }
      const rel = relative(cwd, resolved)
      if (rel.startsWith("..")) {
        throw new PathSecurityError(`Path traversal detected: ${resolved}`)
      }
    }
  }

  private rejectIfNetworkPath(resolved: string): void {
    if (!this.policy.allowNetworkPaths) {
      if (resolved.startsWith("\\\\") || resolved.startsWith("//")) {
        throw new PathSecurityError(`Network paths not allowed: ${resolved}`)
      }
    }
  }

  private rejectIfBlocklisted(resolved: string): void {
    const blocklist = [
      /^[a-zA-Z]:\\windows(\\|$)/i,
      /^[a-zA-Z]:\\program files(\\|$)/i,
      /^[a-zA-Z]:\\program files \(x86\)(\\|$)/i,
      /^[a-zA-Z]:\\programdata(\\|$)/i,
    ]
    for (const pattern of blocklist) {
      if (pattern.test(resolved)) {
        throw new PathSecurityError(`Access to system directory blocked: ${resolved}`)
      }
    }
  }

  private rejectIfSpecialDevicePath(rawPath: string): void {
    if (rawPath.startsWith("\\\\.") || rawPath.startsWith("\\\\.\\") ||
        rawPath.startsWith("\\\\?\\") || rawPath.startsWith("\\??\\")) {
      throw new PathSecurityError(`Device path not allowed: ${rawPath}`)
    }
    const normal = rawPath.replace(/\//g, "\\")
    const fileName = normal.split("\\").pop() ?? normal
    if (fileName.includes(":") && normal !== fileName) {
      throw new PathSecurityError(`Alternate data stream not allowed: ${rawPath}`)
    }
  }

  private rejectIfNotAllowed(resolved: string): void {
    const allowed = this.policy.allowedDirectories.map((d) => resolve(normalize(d)))
    const isAllowed = allowed.some((dir) => resolved.startsWith(dir))
    if (!isAllowed) {
      throw new PathSecurityError(`Path not in allowed directories: ${resolved}`)
    }
  }
}
