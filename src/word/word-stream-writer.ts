import { WordBase } from "./word-base.js"
import type { IWordSession } from "./session.js"
import type { WordMarkdown } from "./word-markdown.js"
import type { WordApplicationManager } from "./application.js"
import type { WordFormatting } from "./formatting.js"
import type { IStreamLock } from "./types.js"

import type { StyleProfile } from "./formatting.js"

export interface StreamStartParams {
  title?: string
  author?: string
  templatePath?: string
  orientation?: "portrait" | "landscape"
  topMargin?: number
  bottomMargin?: number
  leftMargin?: number
  rightMargin?: number
  baseStyleProfile?: Record<string, StyleProfile>
}

export interface StreamBlockResult {
  blockType: string
  chars: number
  blockIndex: number
}

export interface StreamEndResult {
  blockCount: number
  charCount: number
  elapsed: number
  savedPath?: string
  pdfPath?: string
}

interface StreamSession {
  blockCount: number
  charCount: number
  startTime: number
  isActive: boolean
}

export class StreamingMarkdownWriter extends WordBase {
  private streamSession: StreamSession | null = null

  constructor(
    session: IWordSession,
    private markdown: WordMarkdown,
    private appManager: WordApplicationManager,
    private formatting: WordFormatting,
    private director: IStreamLock,
  ) {
    super(session)
  }

  get isActive(): boolean {
    return this.streamSession?.isActive === true
  }

  async start(params: StreamStartParams): Promise<string> {
    // 清理任何可能残留的旧流式会话（文档被 word_close / word_quit 外部关闭后）
    if (this.streamSession?.isActive) {
      try {
        // 试探获取活跃文档，若 COM 已断开或文档已关闭则静默清理
        const doc = this.session.activeDoc
        if (doc) {
          const _ = (doc.Name as string) // 探活
          // 文档还在，不能覆盖——通知用户先正常结束
          throw new Error("已有活跃的流式会话，请先调用 word_stream_end 结束")
        }
      } catch (e) {
        if ((e as Error).message?.includes("已有活跃的流式会话")) throw e
        // COM 或文档已不可用，自动清理旧会话
      }
      this.streamSession = null
    }

    if (this.session.activeDoc) {
      try {
        const doc = this.session.activeDoc
        const saved = (doc.Saved as boolean) ?? true
        if (!saved) {
          const name = (doc.Name as string) ?? "未命名文档"
          throw new Error(`文档"${name}"有未保存的更改。请先用 word_save 保存，或明确放弃更改后再启动流式会话。`)
        }
      } catch (e) {
        if ((e as Error).message?.includes("未保存的更改")) throw e
      }
      await this.appManager.closeDocument(false)
    }
    await this.session.start()

    const lockErr = this.director.acquireStreamLock("word_stream_start")
    if (lockErr) throw new Error(lockErr)

    try {
      if (params.templatePath) {
        await this.appManager.createDocumentFromTemplate(params.templatePath, {
          title: params.title,
          author: params.author,
        })
      } else {
        await this.appManager.createDocument({
          title: params.title,
          author: params.author,
        })
      }

      if (params.topMargin != null || params.bottomMargin != null ||
          params.leftMargin != null || params.rightMargin != null ||
          params.orientation != null) {
        await this.formatting.setPageSetup({
          topMargin: params.topMargin, bottomMargin: params.bottomMargin,
          leftMargin: params.leftMargin, rightMargin: params.rightMargin,
          orientation: params.orientation,
        })
      }

      // 应用默认正文样式 + 用户自定义样式配置
      try {
        await this.formatting.modifyStyle("Normal", {
          paragraph: { spaceAfter: 6, firstLineIndent: 0.74 },
        })
      } catch { /* ignore */ }

      if (params.baseStyleProfile) {
        for (const [styleName, profile] of Object.entries(params.baseStyleProfile)) {
          try {
            await this.formatting.modifyStyle(styleName, profile)
          } catch { /* skip invalid styles */ }
        }
      }
    } catch (e) {
      this.director.releaseStreamLock()
      throw e
    }

    this.goToEnd()
    this.streamSession = {
      blockCount: 0,
      charCount: 0,
      startTime: Date.now(),
      isActive: true,
    }
    return "流式会话已启动"
  }

  async writeBlock(text: string): Promise<StreamBlockResult> {
    if (!this.streamSession?.isActive) {
      throw new Error("没有活跃的流式会话。请先调用 word_stream_start 创建文档。")
    }

    this.goToEnd()
    try {
      const result = await this.markdown.writeBlocks(text)
      this.director.refreshWatchdog()
      this.streamSession.blockCount += result.blocks
      this.streamSession.charCount += result.chars
      const blockType = this.detectBlockType(text)
      return { blockType, chars: result.chars, blockIndex: this.streamSession.blockCount }
    } catch (e) {
      this.streamSession = null
      throw e
    }
  }

  async end(params: { save?: boolean; exportPath?: string }): Promise<StreamEndResult> {
    if (!this.streamSession) {
      throw new Error("没有活跃的流式会话。")
    }

    const session = this.streamSession
    this.director.releaseStreamLock()
    this.streamSession = null

    let savedPath: string | undefined
    let pdfPath: string | undefined

    // 若文档已不存在（被外部关闭），跳过所有文档操作，直接返回统计
    let docAlive = false
    try {
      const doc = this.session.activeDoc
      if (doc) {
        const _ = (doc.Name as string)
        docAlive = true
      }
    } catch { /* doc is gone */ }

    if (docAlive) {
      try {
        try {
          ;(this.getSelection().TypeParagraph as () => void)()
        } catch { }

        if (params.save !== false) {
          await this.appManager.saveDocument()
          try {
            const doc = this.requireDoc()
            savedPath = (doc.FullName as string) || undefined
          } catch { }
        }

        if (params.exportPath) {
          try {
            const doc = this.requireDoc()
            ;(doc.ExportAsFixedFormat as (path: string, format: number) => void)(params.exportPath, 17)
            pdfPath = params.exportPath
          } catch { }
        }
      } catch { /* ignore document-operation errors during end */ }
    }

    return {
      blockCount: session.blockCount,
      charCount: session.charCount,
      elapsed: Date.now() - session.startTime,
      savedPath,
      pdfPath,
    }
  }

  private detectBlockType(text: string): string {
    const trimmed = text.trim()
    if (/^#{1,6}\s+/.test(trimmed)) return "heading"
    if (trimmed.startsWith("```")) return "codeblock"
    if (trimmed.startsWith("![")) return "image"
    if (trimmed === "[[pagebreak]]") return "pagebreak"
    if (/^[-*+]\s/.test(trimmed)) return "list"
    if (/^\d+[.)]\s/.test(trimmed)) return "list"
    if (trimmed.startsWith("> ")) return "blockquote"
    if (trimmed.startsWith("|")) return "table"
    if (/^(-{3,}|\*{3,}|_{3,})\s*$/.test(trimmed)) return "hr"
    return "paragraph"
  }
}
