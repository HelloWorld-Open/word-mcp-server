import { z } from "zod"
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js"
import { WordApplicationManager } from "../../word/application.js"
import { WordDocument } from "../../word/document.js"
import { SecurityManager } from "../../security/policy.js"
import type { ServerContext } from "../server-context.js"
import { createRegTool } from "./shared.js"

export function registerDocumentTools(
  server: McpServer,
  context: ServerContext,
  appManager: WordApplicationManager,
  docOps: WordDocument,
  security: SecurityManager,
): void {
  const regTool = createRegTool(server, security, context)
  regTool("word_get_status",
    {
      description: "Query the current state of Word and document(s). WHEN: always call first to understand what state Word is in before deciding next action. NOT: want document content? use word_get_text.",
      inputSchema: {
        path: z.string().min(1).max(4096).optional().describe("Optional path to get file-specific status (exists? open? active? tracked?)"),
      },
    },
    async ({ path }) => {
      const safePath = path ? security.pathSanitizer.resolveAndValidate(path) : undefined
      const status = await appManager.getStatus(safePath)

      if (!status.wordRunning) {
        return "State: NO_WORD — Word is not running."
      }

      if (status.dialogBlocked) {
        return "State: DIALOG — Word is blocked by a dialog (macro security, file-in-use, etc.). Close the dialog in the Word window to continue, or wait for auto-recovery."
      }

      if (safePath) {
        const active = status.activeDocument
        const p = status.pathStatus!
        const activeDocLine = active
          ? `Active: "${active.name}"`
          : "Active: NONE"
        const targetLine = `Target: ${p.path} — Exists: ${p.existsOnDisk} — Open: ${p.isOpenInWord} — Active: ${p.isActive}`
        let actionLine = ""
        if (p.existsOnDisk) {
          actionLine = p.isOpenInWord
            ? `→ Already open. Use word_document("${p.path}") to switch.`
            : `→ On disk. Use word_document("${p.path}") to open.`
        } else {
          actionLine = "→ Not on disk. Use word_stream_start to create."
        }
        return `State: DOC_ACTIVE — ${activeDocLine}\n${targetLine}\n${actionLine}`
      }

      if (status.activeDocument) {
        const d = status.activeDocument
        return `State: DOC_ACTIVE — Active: "${d.name}" — ${status.openDocuments.length} document(s) open — Saved: ${d.saved}`
      }
      return "State: NO_DOC — Word is running, no active document."
    },
    { preconditions: [] },
  )

  regTool("word_document",
    {
      description: "Switch to or open an existing document by path. WHEN: need to change active document. NOT: want to create new blank doc? use word_stream_start.",
      inputSchema: {
        path: z.string().min(1).max(4096).optional().describe("Path to an existing file. If the file exists, it will be opened; if not, an error is returned. Omit to reuse active document or create an untitled one."),
        title: z.string().max(255).optional().describe("Document title (used only when creating an untitled document with no path)"),
      },
    },
    async ({ path, title }) => {
      const safePath = path ? security.pathSanitizer.resolveAndValidate(path) : undefined
      const action = await appManager.ensureDocument(safePath, title)
      context.director?.enterEditMode()
      const doc = await appManager.getStatus()
      const name = doc.activeDocument?.name ?? "untitled"
      const labels: Record<string, string> = { created: "Created", opened: "Opened", reused: "Reused" }
      return `Action: ${labels[action]} "${name}"\nNext: word_get_info() or word_get_structure() or word_insert_at(...)`
    },
    { preconditions: [] },
  )

  regTool("word_open",
    {
      description: "Open an existing Word document from disk. WHEN: need to open a file that is not already tracked by the session. NOT: want to switch to an already-open document? use word_document (preferred in most cases).",
      inputSchema: {
        path: z.string().min(1).max(4096).describe("Full path to the .docx file"),
      },
    },
    async ({ path }) => {
      const safePath = security.pathSanitizer.validateForRead(path)
      const result = await appManager.openDocument(safePath)
      context.director?.enterEditMode()
      return `Action: Document opened "${result.name}"\nDetail: ${result.fullName}\nNext: word_get_info() or word_get_structure() or word_insert_at(...)`
    },
  )

  regTool("word_save",
    {
      description: "Save the current document (creates .bak backup before overwriting). WHEN: want to persist changes after editing. NOT: want to save to a different file or format? use word_save_as.",
    },
    async () => {
      await appManager.saveDocument()
      return "Action: Document saved\nNext: word_close({saveChanges: true}) or word_save_as({path:\"C:\\output.docx\"})"
    },
  )

  regTool("word_save_as",
    {
      description: "Save the current document to a new file or format. WHEN: need to create a copy or export to a different format (PDF, RTF, TXT, etc.). NOT: just want to save changes to the current file? use word_save.",
      inputSchema: {
        path: z.string().min(1).max(4096).describe("Full save path for the document"),
        format: z.enum(["docx", "doc", "pdf", "rtf", "txt", "html", "mht", "xml", "odt", "dotx", "dotm", "docm"]).optional().describe("File format (default: determined from file extension)"),
      },
    },
    async ({ path, format }) => {
      const safePath = security.pathSanitizer.validateForWrite(path)
      const result = await appManager.saveDocumentAs(safePath, format)
      return `Action: Document saved to ${result.path}\nNext: word_close({saveChanges: false})`
    },
  )

  regTool("word_close",
    {
      description: "Close the current document (does not close Word). WHEN: done editing and want to return to NO_DOC state or switch to a different workflow. NOT: want to quit Word entirely? use word_quit.",
      inputSchema: {
        saveChanges: z.boolean().optional().describe("Whether to save changes before closing (default: false)"),
      },
    },
    async ({ saveChanges }) => {
      await appManager.closeDocument(saveChanges)
      context.director?.exitEditMode()
      context.director?.releaseStreamLock()
      const saved = saveChanges ?? false
      return `Action: Document closed\nDetail: Changes saved: ${saved}\nNext: word_document({path:"C:\\file.docx"}) or word_stream_start({title:"New Doc"})`
    },
  )

  regTool("word_get_info",
    {
      description: "Get document statistics (word count, paragraph count, page count, etc.). WHEN: need to check document size, table count, or save status. NOT: want heading outline? use word_get_structure.",
    },
    async () => {
      const status = await appManager.getStatus()
      if (!status.activeDocument) {
        return "No active document. Use word_document(path) to open a file, or word_stream_start to create a new one."
      }
      const info = await docOps.getInfo()
      return `Action: Document info\nDetail: "${info.name}" — ${info.wordCount} words, ${info.paragraphCount} paras, ${info.pageCount} pages, ${info.characterCount} chars, ${info.sectionCount} sections, saved: ${info.saved}\nNext: word_get_structure() or word_insert_at(...)`
    },
  )

  regTool("word_get_structure",
    {
      description: "Get the heading structure of the document (hierarchical outline with paragraph indices). WHEN: need to navigate or understand the document's organization. NOT: want raw text content? use word_get_text.",
    },
    async () => {
      const structure = await docOps.getStructure()
      if (structure.headings.length === 0) {
        return `Action: No headings found\nDetail: ${structure.totalParagraphs} total paragraphs\nNext: word_go_to_paragraph({index:2}) to navigate or word_apply_style({styleName:"Heading 1"}) to add`
      }
      const lines = [
        `Action: Document structure — ${structure.totalParagraphs} paragraphs, ${structure.headings.length} headings`,
        `Next: word_go_to_paragraph({index:H.paragraphIndex}) to navigate`,
        "",
      ]
      for (const h of structure.headings) {
        const indent = "  ".repeat(h.level - 1)
        lines.push(`${indent}H${h.level} ¶${h.paragraphIndex} — ${h.text}`)
      }
      return lines.join("\n")
    },
  )

  regTool("word_quit",
    {
      description: "Quit Microsoft Word entirely (does not lose already-saved documents). WHEN: need to completely shut down Word. NOT: just want to close the current document? use word_close.",
    },
    async () => {
      await appManager.quit()
      return "Action: Word closed\nNext: Restart MCP server to reconnect to Word"
    },
    { preconditions: [] },
  )
}
