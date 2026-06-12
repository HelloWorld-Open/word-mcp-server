import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js"
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js"
import { WordSession } from "./word/session.js"
import { WordApplicationManager } from "./word/application.js"
import { WordDocument } from "./word/document.js"
import { WordTextEditor } from "./word/word-text-editor.js"
import { WordTableEditor } from "./word/word-table-editor.js"
import { WordMediaEditor } from "./word/word-media-editor.js"
import { WordDocumentStructure } from "./word/word-document-structure.js"
import { WordFormatting } from "./word/formatting.js"
import { VariableReplacer } from "./word/variable-replacer.js"
import { WordMarkdown } from "./word/word-markdown.js"
import { PositionMap } from "./word/position-map.js"
import { WordDocumentManager } from "./manager/word-manager.js"
import { SecurityManager } from "./security/policy.js"

import { registerDocumentTools } from "./server/tools/document.js"
import { registerContentTools } from "./server/tools/content.js"
import { registerFormattingTools } from "./server/tools/formatting.js"
import { registerTableTools } from "./server/tools/tables.js"
import { registerMediaTools } from "./server/tools/media.js"
import { registerStructureTools } from "./server/tools/structure.js"
import { registerMarkdownTool } from "./server/tools/markdown.js"
import { registerReaderTools } from "./server/tools/reader.js"
import { registerBatchTools } from "./server/tools/batch.js"
import { registerVariableTool } from "./server/tools/variable.js"
import { registerSemanticTools } from "./server/tools/semantic.js"
import { registerManagerTools } from "./server/tools/manager.js"
import { registerReportPrompts } from "./server/prompts/report-prompts.js"
import { registerStateMachinePrompt } from "./server/prompts/state-machine.js"
import { setStatusSession, setPositionMap } from "./server/tools/helper.js"

function createLogger(prefix: string) {
  return (...args: unknown[]) => {
    console.error(`[${prefix}]`, ...args)
  }
}

async function main() {
  const log = createLogger("word-mcp")

  // 定时 stderr 心跳，让看门狗检测进程是否存活
  const heartbeat = setInterval(() => {
    console.error("[heartbeat] alive")
  }, 10_000)

  log("Starting Word MCP Server (child process)...")

  const session = new WordSession()
  session.setOnLog((level, message) => {
    log(`[session/${level}] ${message}`)
  })
  setStatusSession(session)

  const appManager = new WordApplicationManager(session)
  const positionMap = new PositionMap(session)
  const docOps = new WordDocument(session, positionMap)
  const textEditor = new WordTextEditor(session)
  const tableEditor = new WordTableEditor(session)
  const mediaEditor = new WordMediaEditor(session)
  const documentStructure = new WordDocumentStructure(session)
  const formatting = new WordFormatting(session)
  const markdown = new WordMarkdown(session)
  const variableReplacer = new VariableReplacer(session)
  const security = new SecurityManager()
  const documentManager = new WordDocumentManager(
    session, appManager, docOps, textEditor, tableEditor,
    mediaEditor, documentStructure, formatting, markdown,
  )

  setPositionMap(positionMap)

  const cleanup = async () => {
    log("Shutting down...")
    clearInterval(heartbeat)
    try { await appManager.quit() } catch { /* ignore */ }
    process.exit(0)
  }
  process.on("SIGINT", cleanup)
  process.on("SIGTERM", cleanup)
  process.on("unhandledRejection", (reason) => {
    log("Unhandled rejection:", reason)
  })
  process.on("uncaughtException", (err) => {
    log("Uncaught exception:", err)
    process.exit(1)
  })

  const server = new McpServer(
    {
      name: "word-mcp-server",
      version: "1.0.0",
    },
    {
      instructions: [
        "Word MCP Server — Microsoft Word COM automation.",
        "",
        "States (call word_get_status first):",
        "- NO_WORD: restart server.",
        "- NO_DOC: word_create (new) or word_document(path).",
        "- DOC_ACTIVE: all tools available.",
        "- DIALOG: dismiss dialog manually.",
        "",
        "New documents: word_mgr_* tools (auto cursor).",
        "Editing: raw word_* tools.",
        "Bulk content: word_write_markdown ONCE (not split).",
        "",
        "ANTI-PATTERNS:",
        "1. word_create while doc active → close first.",
        "2. word_save then word_undo → undo cleared.",
        "3. Format before typing → type first, select, format.",
        "4. word_open → use word_document (smart switch).",
        "5. word_find_text + word_go_to_paragraph → find already moves cursor.",
        "6. Multiple word_write_markdown calls → one call covers all.",
      ].join("\n"),
    }
  )

  registerDocumentTools(server, appManager, docOps, security)
  registerContentTools(server, textEditor, security)
  registerFormattingTools(server, formatting, security)
  registerTableTools(server, tableEditor, security)
  registerMediaTools(server, mediaEditor, security)
  registerStructureTools(server, documentStructure, security)
  registerReportPrompts(server)
  registerStateMachinePrompt(server)
  registerMarkdownTool(server, markdown, textEditor, positionMap, security)
  registerVariableTool(server, variableReplacer, security)
  registerReaderTools(server, docOps, security)
  registerBatchTools(server, session, security)
  registerSemanticTools(server, textEditor, tableEditor, positionMap, security)
  registerManagerTools(server, documentManager, security)

  const transport = new StdioServerTransport()
  await server.connect(transport)

  log("Word MCP Server (child) connected and ready")
}

main().catch((error) => {
  console.error("[word-mcp/child] Fatal error:", error)
  process.exit(1)
})
