import { WordBase } from "./word-base.js"
import type { IDocumentProxy, ISelectionProxy } from "./com-proxy/types.js"
import type { IWordSession } from "./session.js"
import type { WordApplicationManager } from "./application.js"
import type { WordContentWriter } from "./word-content-writer.js"
import type { WordFormatter } from "./word-formatter.js"
import type { IStreamLock } from "./types.js"
import { WordMcpError } from "../security/errors.js"
import { EXPORT_FORMAT_PDF } from "./types.js"

import type { StyleProfile } from "./word-formatter.js"

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
    private contentWriter: WordContentWriter,
    private appManager: WordApplicationManager,
    private formatting: WordFormatter,
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
        if (this.session.activeDoc) {
          const _ = this.getDocProxy().getName()
          // 文档还在，不能覆盖——通知用户先正常结束
          throw new WordMcpError("Active stream session exists. Call word_stream_end first.", "STREAM_SESSION_ACTIVE", true, "End the current session with word_stream_end then retry.")
        }
      } catch (e) {
        if ((e as Error).message?.includes("Active stream session exists")) throw e
        // COM 或文档已不可用，自动清理旧会话
      }
      this.streamSession = null
    }

    if (this.session.activeDoc) {
      try {
        const dp = this.getDocProxy()
        const saved = dp.getSaved()
        if (!saved) {
          const name = dp.getName() ?? "未命名文档"
          throw new WordMcpError(`Document "${name}" has unsaved changes. Use word_save to save, or close without saving then retry.`, "UNSAVED_CHANGES", true, "Use word_save to save the current document before starting a stream session.")
        }
      } catch (e) {
        if ((e as Error).message?.includes("unsaved changes")) throw e
      }
      await this.appManager.closeDocument(false)
    }
    await this.session.start()

    const lockErr = this.director.acquireStreamLock("word_stream_start")
    if (lockErr) throw new WordMcpError(lockErr, "STREAM_LOCK_DENIED", true, lockErr)

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
    return "Stream session started"
  }

  async writeBlock(text: string): Promise<StreamBlockResult> {
    if (!this.streamSession?.isActive) {
      throw new WordMcpError("No active stream session. Use word_stream_start to create a document.", "NO_STREAM_SESSION", true, "Create a new document with word_stream_start.")
    }

    this.goToEnd()
    try {
      const result = await this.contentWriter.writeBlocks(text)
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
      throw new WordMcpError("No active stream session.", "NO_STREAM_SESSION", true, "Create a new document with word_stream_start.")
    }

    const session = this.streamSession
    this.director.releaseStreamLock()
    this.streamSession = null

    let savedPath: string | undefined
    let pdfPath: string | undefined

    // 若文档已不存在（被外部关闭），跳过所有文档操作，直接返回统计
    let docAlive = false
    try {
      if (this.session.activeDoc) {
        const _ = this.getDocProxy().getName()
        docAlive = true
      }
    } catch { /* doc is gone */ }

    if (docAlive) {
      try {
        try {
          this.getSelProxy().typeParagraph()
        } catch { }

        if (params.save !== false) {
          await this.appManager.saveDocument()
          try {
            savedPath = this.getDocProxy().getFullName() || undefined
          } catch { }
        }

        if (params.exportPath) {
          try {
            this.getDocProxy().exportAsFixedFormat(params.exportPath, EXPORT_FORMAT_PDF)
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
