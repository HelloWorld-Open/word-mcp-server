import { HeaderFooterProxy } from "./header-footer-proxy.js"
import type { IHeaderFooterProxy, IRangeProxy, ISectionProxy } from "./types.js"

interface ISectionDeps {
  getSelectionProxy: () => {
    typeText: (t: string) => void
    delete: () => void
    getParagraphFormat: () => Record<string, unknown>
  }
  wrapRange: (raw: Record<string, unknown>) => IRangeProxy
  log?: (level: string, msg: string) => void
}

export class SectionProxy implements ISectionProxy {
  private deps: ISectionDeps

  constructor(
    private rawSection: Record<string, unknown>,
    deps: ISectionDeps,
  ) {
    this.deps = deps
  }

  get raw(): Record<string, unknown> { return this.rawSection }

  getHeader(index: number = 1): IHeaderFooterProxy {
    try {
      const hdr = (this.rawSection.Headers as { Item: (i: number) => Record<string, unknown> }).Item(index)
      return new HeaderFooterProxy(hdr, this.deps)
    } catch {
      this.deps.log?.("debug", "SectionProxy.getHeader failed")
      return new HeaderFooterProxy({} as Record<string, unknown>, this.deps)
    }
  }

  getFooter(index: number = 1): IHeaderFooterProxy {
    try {
      const ftr = (this.rawSection.Footers as { Item: (i: number) => Record<string, unknown> }).Item(index)
      return new HeaderFooterProxy(ftr, this.deps)
    } catch {
      this.deps.log?.("debug", "SectionProxy.getFooter failed")
      return new HeaderFooterProxy({} as Record<string, unknown>, this.deps)
    }
  }

  getPageSetup(): Record<string, unknown> {
    try { return this.rawSection.PageSetup as Record<string, unknown> } catch { return {} }
  }
}
