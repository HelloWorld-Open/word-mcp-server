import type { IWordSession } from "./session.js"
import { WordBase } from "./word-base.js"
import { parseBlocks } from "./markdown-parser.js"
import { MarkdownRenderer } from "./markdown-renderer.js"

export class WordMarkdown extends WordBase {
  private renderer: MarkdownRenderer

  constructor(session: IWordSession) {
    super(session)
    this.renderer = new MarkdownRenderer({
      getSelection: () => this.getSelection(),
      requireDoc: () => this.requireDoc(),
      goToEnd: () => this.goToEnd(),
    })
  }

  async writeBlocks(markdown: string): Promise<{ blocks: number; chars: number }> {
    const blocks = parseBlocks(markdown)
    if (blocks.length === 0) return { blocks: 0, chars: 0 }
    let totalChars = 0
    this.collapseSelection()

    const word = this.getWord()
    let needsEnd = true

    try {
      for (let bi = 0; bi < blocks.length; bi++) {
        totalChars += await this.renderer.renderBlock(blocks[bi], bi, blocks.length, "end", !needsEnd)
        needsEnd = blocks[bi].type === "table"
        try { ;(word.ScreenRefresh as () => void)() } catch { /* ignore */ }
      }
    } finally {
      try { word.ScreenUpdating = true } catch { /* ignore */ }
    }

    try {
      ;(this.getSelection().TypeParagraph as () => void)()
    } catch { /* ignore */ }
    return { blocks: blocks.length, chars: totalChars }
  }

  async insertAtCursor(markdown: string): Promise<{ blocks: number; chars: number }> {
    const blocks = parseBlocks(markdown)
    if (blocks.length === 0) return { blocks: 0, chars: 0 }
    let totalChars = 0
    this.collapseSelection()

    const TIME_BUDGET = 50
    const word = this.getWord()
    let bi = 0

    try {
      while (bi < blocks.length) {
        try { word.ScreenUpdating = false } catch { /* ignore */ }

        const batchStart = Date.now()
        do {
          totalChars += await this.renderer.renderBlock(blocks[bi], bi, blocks.length, "cursor")
          bi++
        } while (bi < blocks.length && Date.now() - batchStart < TIME_BUDGET)

        try { word.ScreenUpdating = true } catch { /* ignore */ }
        try { ;(word.ScreenRefresh as () => void)() } catch { /* ignore */ }
      }
    } finally {
      try { word.ScreenUpdating = true } catch { /* ignore */ }
    }

    return { blocks: blocks.length, chars: totalChars }
  }
}
