import { normalizePath } from "./types.js"

export interface RegisteredDocument {
  doc: Record<string, unknown>
  path: string
  displayName: string
  openedAt: number
}

export class DocumentRegistry {
  private entries = new Map<string, RegisteredDocument>()
  private activePath: string | null = null

  register(path: string, doc: Record<string, unknown>): void {
    const key = this.keyFor(path)
    this.entries.set(key, { doc, path, displayName: String(doc.Name ?? "unknown"), openedAt: Date.now() })
    this.activePath = key
  }

  unregister(path: string): void {
    const key = this.keyFor(path)
    this.entries.delete(key)
    if (this.activePath === key) {
      this.activePath = null
    }
  }

  getByPath(path: string): RegisteredDocument | null {
    return this.entries.get(this.keyFor(path)) ?? null
  }

  isOpen(path: string): boolean {
    return this.entries.has(this.keyFor(path))
  }

  setActive(path: string): void {
    const key = this.keyFor(path)
    if (this.entries.has(key)) {
      this.activePath = key
    }
  }

  listAll(): RegisteredDocument[] {
    return Array.from(this.entries.values())
  }

  clear(): void {
    this.entries.clear()
    this.activePath = null
  }

  /**
   * Remove entries whose COM doc object is stale (e.g. closed via Word UI).
   */
  pruneStale(): void {
    for (const [key, entry] of this.entries) {
      try {
        const _ = entry.doc.Name
      } catch {
        this.entries.delete(key)
        if (this.activePath === key) this.activePath = null
      }
    }
  }

  private keyFor(p: string): string {
    return normalizePath(p)
  }
}
