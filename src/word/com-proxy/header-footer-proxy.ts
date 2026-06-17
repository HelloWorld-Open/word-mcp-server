import { CollectionProxy } from "./collection-proxy.js"
import type { ICollectionProxy, IFieldProxy, IHeaderFooterProxy, IRangeProxy } from "./types.js"

interface IHFProxyDeps {
  getSelectionProxy: () => {
    typeText: (t: string) => void
    delete: () => void
    getParagraphFormat: () => Record<string, unknown>
  }
  wrapRange: (raw: Record<string, unknown>) => IRangeProxy
  log?: (level: string, msg: string) => void
}

class FieldProxy implements IFieldProxy {
  constructor(private rawField: Record<string, unknown>) {}

  get type(): number {
    try { return this.rawField.Type as number } catch { return -1 }
  }

  delete(): void {
    try { ;(this.rawField.Delete as () => void)() } catch { /* ignore */ }
  }
}

export class HeaderFooterProxy implements IHeaderFooterProxy {
  constructor(
    private rawHF: Record<string, unknown>,
    private deps: IHFProxyDeps,
  ) {}

  get raw(): Record<string, unknown> { return this.rawHF }

  select(): void {
    this.getRange().select()
  }

  clearContent(): void {
    this.select()
    try {
      this.deps.getSelectionProxy().delete()
    } catch { this.deps.log?.("warn", "HeaderFooterProxy.clearContent failed") }
  }

  typeText(text: string): void {
    try {
      this.deps.getSelectionProxy().typeText(text)
    } catch { this.deps.log?.("error", "HeaderFooterProxy.typeText failed") }
  }

  setAlignment(align: number): void {
    try {
      const pf = this.deps.getSelectionProxy().getParagraphFormat()
      pf.Alignment = align
    } catch { this.deps.log?.("warn", "HeaderFooterProxy.setAlignment failed") }
  }

  setContent(text: string, alignment?: number): void {
    this.clearContent()
    if (alignment != null) this.setAlignment(alignment)
    if (text) this.typeText(text)
  }

  getEnd(): number {
    return this.getRange().getEnd()
  }

  getText(): string {
    return this.getRange().getText()
  }

  getRange(): IRangeProxy {
    try {
      const rng = this.rawHF.Range as Record<string, unknown>
      return this.deps.wrapRange(rng)
    } catch {
      this.deps.log?.("debug", "HeaderFooterProxy.getRange failed")
      return this.deps.wrapRange({} as Record<string, unknown>)
    }
  }

  getFields(): ICollectionProxy<IFieldProxy> {
    const rawFields = this.rawHF.Fields as Record<string, unknown>
    if (!rawFields) {
      return new CollectionProxy<IFieldProxy>({} as Record<string, unknown>, () => ({ type: -1, delete: () => {} }))
    }
    return new CollectionProxy(rawFields, (raw) => new FieldProxy(raw))
  }

  getPageNumbersCount(): number {
    try {
      const pn = this.rawHF.PageNumbers as { Count: number }
      return pn.Count as number
    } catch { return 0 }
  }

  setPageNumbersAlignment(val: number): void {
    try {
      const pn = this.rawHF.PageNumbers as { Count: number; Item: (i: number) => Record<string, unknown> }
      if ((pn.Count as number) > 0) {
        pn.Item(1).Alignment = val
      }
    } catch { this.deps.log?.("warn", "HeaderFooterProxy.setPageNumbersAlignment failed") }
  }
}
