import { tmpdir } from "node:os"
import { join, resolve, normalize } from "node:path"
import { existsSync, copyFileSync } from "node:fs"
import type { IWordSession } from "./session.js"
import { DocumentRegistry } from "./document-registry.js"
import { WordMcpError } from "../security/errors.js"

const FORMAT_EXT_MAP: Record<string, number> = {
  ".docx": 16, ".doc": 0, ".pdf": 17, ".rtf": 6, ".txt": 2,
  ".htm": 8, ".html": 8, ".mht": 9, ".mhtml": 9, ".xml": 11,
  ".odt": 23, ".dotx": 15, ".dotm": 13, ".docm": 12,
}

const FORMAT_NAMES: Record<string, number> = {
  docx: 16, doc: 0, pdf: 17, rtf: 6, txt: 2,
  htm: 8, html: 8, mht: 9, mhtml: 9, xml: 11,
  odt: 23, dotx: 15, dotm: 13, docm: 12,
}

function formatForPath(p: string): number {
  const ext = p.match(/\.(\w+)$/)?.[0]?.toLowerCase()
  return FORMAT_EXT_MAP[ext ?? ""] ?? 16
}

interface ActiveDocumentInfo {
  name: string; fullName: string; path: string; saved: boolean
}

interface OpenDocumentInfo {
  name: string; fullName: string; path: string
}

interface PathStatus {
  path: string
  existsOnDisk: boolean
  isOpenInWord: boolean
  isActive: boolean
  isTrackedByRegistry: boolean
}

interface WordStatusResult {
  wordRunning: boolean
  activeDocument: ActiveDocumentInfo | null
  openDocuments: OpenDocumentInfo[]
  trackedByRegistry: { name: string; path: string }[]
  pathStatus?: PathStatus
}

export class WordApplicationManager {
  public readonly registry: DocumentRegistry

  constructor(private session: IWordSession) {
    this.registry = new DocumentRegistry()
  }

  private getWord(): Record<string, unknown> {
    return this.session.application as Record<string, unknown>
  }

  private tryGetWordSafe(): Record<string, unknown> | null {
    try { return this.session.application } catch { return null }
  }

  private getDoc(): Record<string, unknown> {
    return (this.session.activeDoc ?? (this.getWord().ActiveDocument as Record<string, unknown>)) as Record<string, unknown>
  }

  isDocumentActive(): boolean {
    return this.session.activeDoc !== null
  }

  getActiveDocumentName(): string | null {
    const doc = this.session.activeDoc
    return doc ? (doc.Name as string) : null
  }

  private isPathOpenInWord(path: string): Record<string, unknown> | null {
    const word = this.tryGetWordSafe()
    if (!word) return null
    const keyPath = resolve(normalize(path)).toLowerCase()
    const docs = word.Documents as { Count: number; Item: (i: number) => Record<string, unknown> }
    for (let i = 1; i <= docs.Count; i++) {
      const doc = docs.Item(i) as Record<string, unknown>
      try {
        const raw = (doc.FullName as string) ?? ""
        if (!raw) continue
        const fp = resolve(normalize(raw)).toLowerCase()
        if (fp === keyPath) return doc
      } catch { /* skip stale */ }
    }
    return null
  }

  async getStatus(resolvedPath?: string): Promise<WordStatusResult> {
    const word = this.tryGetWordSafe()
    if (!word) {
      return { wordRunning: false, activeDocument: null, openDocuments: [], trackedByRegistry: [] }
    }

    const openDocuments: OpenDocumentInfo[] = []
    let activeDocument: ActiveDocumentInfo | null = null

    const docs = word.Documents as { Count: number; Item: (i: number) => Record<string, unknown> }
    const comActiveDoc = word.ActiveDocument as Record<string, unknown> | undefined
    const cachedActiveDoc = this.session.activeDoc

    for (let i = 1; i <= docs.Count; i++) {
      const doc = docs.Item(i) as Record<string, unknown>
      try {
        const name = doc.Name as string
        const fullName = doc.FullName as string
        const docPath = doc.Path as string
        openDocuments.push({ name, fullName, path: docPath })
        const docFullName = (doc.FullName as string) ?? ""
        const comActiveFullName = (comActiveDoc?.FullName as string) ?? ""
        if (comActiveDoc && docFullName && docFullName === comActiveFullName) {
          const saved = (doc.Saved as boolean) ?? false
          activeDocument = { name, fullName, path: docPath, saved }
        }
      } catch { /* skip stale */ }
    }

    if (!activeDocument && cachedActiveDoc && openDocuments.length > 0) {
      try {
        const name = cachedActiveDoc.Name as string
        const fullName = cachedActiveDoc.FullName as string
        activeDocument = { name, fullName, path: cachedActiveDoc.Path as string, saved: (cachedActiveDoc.Saved as boolean) ?? false }
      } catch { /* cached doc is stale */ }
    }

    this.registry.pruneStale()
    const trackedByRegistry = this.registry.listAll().map(e => ({ name: e.displayName, path: e.path }))

    let pathStatus: PathStatus | undefined
    if (resolvedPath) {
      pathStatus = {
        path: resolvedPath,
        existsOnDisk: existsSync(resolvedPath),
        isOpenInWord: this.isPathOpenInWord(resolvedPath) !== null,
        isActive: activeDocument ? resolve(normalize(activeDocument.fullName)).toLowerCase() === resolve(normalize(resolvedPath)).toLowerCase() : false,
        isTrackedByRegistry: this.registry.isOpen(resolvedPath),
      }
    }

    return { wordRunning: true, activeDocument, openDocuments, trackedByRegistry, pathStatus }
  }

  async ensureDocument(path?: string, title?: string): Promise<"reused" | "opened" | "created"> {
    if (path) {
      const alreadyOpen = this.isPathOpenInWord(path)
      if (alreadyOpen) {
        this.session.setActiveDoc(alreadyOpen)
        this.session.setActiveDocPath(path)
        if (this.registry.isOpen(path)) {
          this.registry.setActive(path)
        } else {
          this.registry.register(path, alreadyOpen)
        }
        return "reused"
      }

      if (existsSync(path)) {
        await this.openDocument(path)
        return "opened"
      }

      throw new WordMcpError(
        `File does not exist: "${path}". Use word_create to create a new document, or provide a path to an existing file.`,
        "FILE_NOT_FOUND", false,
        "Use word_get_status(path) to check available files, or word_create to create a new document."
      )
    }

    if (this.session.activeDoc) {
      return "reused"
    }

    await this.createDocument()
    return "created"
  }

  async createDocument(params?: { title?: string; author?: string }): Promise<{ name: string; fullName: string }> {
    const w = this.getWord()
    const docs = w.Documents as { Add: () => Record<string, unknown> }
    const doc = docs.Add() as Record<string, unknown>
    try {
      if (params?.title) {
        const props = doc.BuiltInDocumentProperties as { Item: (n: string) => { Value: string } }
        props.Item("Title").Value = params.title
      }
      if (params?.author) {
        const props = doc.BuiltInDocumentProperties as { Item: (n: string) => { Value: string } }
        props.Item("Author").Value = params.author
      }
    } catch {
      // Suppress property errors on some Word versions
    }
    this.session.setActiveDoc(doc)
    this.session.setActiveDocPath(null)
    const fullName = doc.FullName as string
    if (fullName && fullName !== doc.Name as string) {
      this.registry.register(fullName, doc)
    }
    return { name: doc.Name as string, fullName }
  }

  async createDocumentFromTemplate(
    templatePath: string,
    params?: { title?: string; author?: string },
  ): Promise<{ name: string; fullName: string }> {
    const w = this.getWord()
    const docs = w.Documents as { Add: (t: string) => Record<string, unknown> }
    const doc = docs.Add(templatePath) as Record<string, unknown>
    try {
      if (params?.title) {
        const props = doc.BuiltInDocumentProperties as { Item: (n: string) => { Value: string } }
        props.Item("Title").Value = params.title
      }
      if (params?.author) {
        const props = doc.BuiltInDocumentProperties as { Item: (n: string) => { Value: string } }
        props.Item("Author").Value = params.author
      }
    } catch {
      // Suppress property errors on some Word versions
    }
    this.session.setActiveDoc(doc)
    this.session.setActiveDocPath(null)
    return { name: doc.Name as string, fullName: doc.FullName as string }
  }

  async openDocument(path: string): Promise<{ name: string; fullName: string }> {
    const existing = this.registry.getByPath(path)
    if (existing) {
      try {
        const _ = (existing.doc as Record<string, unknown>).Name
        this.registry.setActive(path)
        this.session.setActiveDoc(existing.doc)
        this.session.setActiveDocPath(path)
        return { name: existing.displayName, fullName: existing.path }
      } catch {
        this.registry.unregister(path)
      }
    }
    const w = this.getWord()
    const docs = w.Documents as { Open: (p: string) => Record<string, unknown> }
    const doc = docs.Open(path) as Record<string, unknown>
    this.session.setActiveDoc(doc)
    this.session.setActiveDocPath(path)
    this.registry.register(path, doc)
    return { name: doc.Name as string, fullName: doc.FullName as string }
  }

  async saveDocument(): Promise<void> {
    const doc = this.getDoc()
    const docPath = doc.Path as string
    if (!docPath || docPath === "") {
      const tempPath = join(tmpdir(), "word-mcp-doc.docx")
      try {
        const bakPath = tempPath + ".bak"
        if (existsSync(tempPath)) {
          copyFileSync(tempPath, bakPath)
        }
      } catch { /* skip backup */ }
      ;(doc.SaveAs as (p: string, f: number) => void)(tempPath, 16)
      this.session.setActiveDoc(doc)
    } else {
      const origPath = doc.FullName as string
      try {
        const bakPath = origPath + ".bak"
        if (existsSync(origPath)) {
          copyFileSync(origPath, bakPath)
        }
      } catch { /* skip backup */ }
      ;(doc.Save as () => void)()
    }
  }

  async saveDocumentAs(path: string, format?: unknown): Promise<{ path: string }> {
    const doc = this.getDoc()
    const f = typeof format === "string" ? (FORMAT_NAMES[format] ?? formatForPath(path)) : (format as number | undefined) ?? formatForPath(path)
    try {
      const bakPath = path + ".bak"
      if (existsSync(path)) {
        copyFileSync(path, bakPath)
      }
    } catch { /* skip backup */ }
    ;(doc.SaveAs as (p: string, f: number) => void)(path, f)
    this.session.setActiveDoc(doc)
    this.session.setActiveDocPath(path)
    this.registry.register(path, doc)
    return { path }
  }

  async closeDocument(saveChanges?: boolean): Promise<void> {
    const doc = this.getDoc()
    const path = this.session.activeDocPath
    ;(doc.Close as (s: boolean) => void)(saveChanges ?? false)
    if (path) this.registry.unregister(path)
    this.session.setActiveDoc(null)
    this.session.setActiveDocPath(null)
  }

  async quit(): Promise<void> {
    this.registry.clear()
    this.session.quit()
  }
}
