import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js"
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js"
import type { ServerContext } from "./server-context.js"
import { WordSession } from "../word/session.js"
import { WordApplicationManager } from "../word/application.js"
import { WordDocument } from "../word/document.js"
import { WordTextEditor } from "../word/word-text-editor.js"
import { WordTableEditor } from "../word/word-table-editor.js"
import { WordMediaEditor } from "../word/word-media-editor.js"
import { ChartDataBridge } from "../word/chart-data-bridge.js"
import { WordDocumentStructure } from "../word/word-document-structure.js"
import { WordFormatting } from "../word/formatting.js"
import { VariableReplacer } from "../word/variable-replacer.js"
import { WordMarkdown } from "../word/word-markdown.js"
import { PositionMap } from "../word/position-map.js"
import { SecurityManager } from "../security/policy.js"
import { SessionDirector } from "./session-director.js"
import { registerDocumentTools } from "./tools/document.js"
import { registerContentTools } from "./tools/content.js"
import { registerFormattingTools } from "./tools/formatting.js"
import { registerTableTools } from "./tools/tables.js"
import { registerMediaTools } from "./tools/media.js"
import { registerStructureTools } from "./tools/structure.js"
/* markdown tools removed — streaming is the sole content path */
import { registerReaderTools } from "./tools/reader.js"
import { registerVariableTool } from "./tools/variable.js"
import { registerSemanticTools } from "./tools/semantic.js"
import { registerManagerTools } from "./tools/manager.js"
import { StreamingMarkdownWriter } from "../word/word-stream-writer.js"
import { registerStreamTools } from "./tools/stream.js"
import { registerReportPrompts } from "./prompts/report-prompts.js"
import { registerStateMachinePrompt } from "./prompts/state-machine.js"


export interface CreateServerOptions {
  instructions: string
  enableHeartbeat?: boolean
  logPrefix?: string
}

function createLogger(prefix: string) {
  return (...args: unknown[]) => {
    console.error(`[${prefix}]`, ...args)
  }
}

export async function createServer(options: CreateServerOptions): Promise<void> {
  const logPrefix = options.logPrefix ?? "word-mcp"
  const log = createLogger(logPrefix)
  const errorTag = logPrefix === "word-mcp" ? "word-mcp" : `word-mcp/${logPrefix}`

  log("Starting Word MCP Server...")

  const heartbeat = options.enableHeartbeat
    ? setInterval(() => { console.error("[heartbeat] alive") }, 10_000)
    : null

  const session = new WordSession()
  session.setOnLog((level, message) => {
    log(`[session/${level}] ${message}`)
  })
  const appManager = new WordApplicationManager(session)
  const positionMap = new PositionMap(session)
  const docOps = new WordDocument(session, positionMap)
  const textEditor = new WordTextEditor(session)
  const tableEditor = new WordTableEditor(session)
  const chartBridge = new ChartDataBridge()
  const mediaEditor = new WordMediaEditor(session, chartBridge)
  const documentStructure = new WordDocumentStructure(session)
  const formatting = new WordFormatting(session)
  const markdown = new WordMarkdown(session)
  const variableReplacer = new VariableReplacer(session)
  const security = new SecurityManager()
  const director = new SessionDirector(session, positionMap)
  const context: ServerContext = { session, positionMap, director }
  const streamWriter = new StreamingMarkdownWriter(
    session, markdown, appManager, formatting, director,
  )

  const cleanup = async () => {
    log("Shutting down...")
    if (heartbeat) clearInterval(heartbeat)
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
    { name: "word-mcp-server", version: "1.0.0" },
    { instructions: options.instructions },
  )

  registerDocumentTools(server, context, appManager, docOps, security)
  registerContentTools(server, context, textEditor, security)
  registerFormattingTools(server, context, formatting, security)
  registerTableTools(server, context, tableEditor, security)
  registerMediaTools(server, context, mediaEditor, security)
  registerStructureTools(server, context, documentStructure, security)
  registerReportPrompts(server)
  registerStateMachinePrompt(server)
  registerVariableTool(server, context, variableReplacer, security)
  registerReaderTools(server, context, docOps, security)
  registerSemanticTools(server, context, textEditor, tableEditor, markdown, positionMap, security)
  registerManagerTools(server, context, tableEditor, mediaEditor, documentStructure, textEditor, security)
  registerStreamTools(server, context, streamWriter, security)

  const transport = new StdioServerTransport()
  await server.connect(transport)

  log("Word MCP Server connected and ready")
}
