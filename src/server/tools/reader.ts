import { z } from "zod"
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js"
import { WordDocument } from "../../word/document.js"
import { SecurityManager } from "../../security/policy.js"
import type { ServerContext } from "../server-context.js"
import { createRegTool } from "./shared.js"

export function registerReaderTools(
  server: McpServer,
  context: ServerContext,
  docOps: WordDocument,
  security: SecurityManager,
): void {
  const regTool = createRegTool(server, security, context)
  regTool("word_get_text",
    {
      description: "Get the full text content of the current document. WHEN: need to read what's written, verify content, or analyze text. NOT: need structure/headings? use word_get_structure.",
    },
    async () => {
      const text = await docOps.getFullText()
      return text
    },
  )

  regTool("word_get_paragraph",
    {
      description: "Get text from a specific paragraph by index. WHEN: need to read a specific section. NOT: want all text? use word_get_text.",
      inputSchema: {
        index: z.number().int().positive().describe("Paragraph index (1-based). Use word_get_structure() to find heading paragraph indices."),
      },
    },
    async ({ index }) => {
      const text = await docOps.getParagraphText(index)
      return `Action: Paragraph ${index}\nDetail: "${text.slice(0, 200)}${text.length > 200 ? "..." : ""}"\nNext: word_get_paragraph({index:${index + 1}}) or word_get_text()`
    },
  )

  regTool("word_export_to_pdf",
    {
      description: "Export the current document to PDF without changing the document. WHEN: need a PDF copy for sharing or preview. NOT: want to save document in another format? use word_save_as.",
      inputSchema: {
        path: z.string().min(1).max(4096).optional().describe("Output PDF path (default: same name as source document with .pdf extension)"),
      },
    },
    async ({ path }) => {
      let outputPath = path
      if (!outputPath) {
        const info = await docOps.getInfo()
        if (info.fullName && info.fullName.includes("\\")) {
          outputPath = info.fullName.replace(/\.docx?$/i, "") + ".pdf"
        } else {
          return "Action: Export requires a path\nDetail: Document is untitled; specify output path\nNext: word_export_to_pdf({path:'C:\\path\\to\\output.pdf'})"
        }
      }
      const safePath = security.pathSanitizer.validateForWrite(outputPath)
      await docOps.exportToPdf(safePath)
      return `Action: PDF exported\nDetail: ${safePath}\nNext: word_get_status() or word_document({path:"${safePath.replace(/\\/g, "\\\\")}"})`
    },
  )

  regTool("word_get_table_data",
    {
      description: "Extract table content as structured data. WHEN: need to verify table content, read table data, or check table structure. NOT: want to edit a table? use word_edit_cells.",
      inputSchema: {
        index: z.number().int().positive().default(1).describe("Table index (1-based). Use word_get_info to check table count."),
      },
    },
    async ({ index }) => {
      const result = await docOps.getTableData(index)
      const header = `Table ${index} of ${result.tableCount} (${result.rows} rows × ${result.columns} columns)`
      const lines = result.data.map((row, ri) =>
        `  Row ${ri + 1}: ${row.map((c, ci) => `[${ci + 1}] ${c}`).join(" | ")}`
      )
      const suffix = result.tableCount > 1
        ? `\nNext: word_get_table_data({index:${index + 1 > result.tableCount ? 1 : index + 1}}) to check other tables`
        : ""
      return `Action: ${header}\nDetail:\n${lines.join("\n")}${suffix}`
    },
  )

  regTool("word_get_comments",
    {
      description: "List all comments in the document. WHEN: need to review existing comments. NOT: want to add a comment? use word_add_comment.",
    },
    async () => {
      const comments = await docOps.getComments()
      if (comments.length === 0) {
        return `Action: No comments found\nDetail: The document has no comments`
      }
      return `Action: ${comments.length} comment(s) found\nDetail:\n${comments.map(c =>
        `  #${c.index} by "${c.author}": "${c.text.slice(0, 200)}${c.text.length > 200 ? "..." : ""}"`
      ).join("\n")}`
    },
  )

  regTool("word_get_bookmarks",
    {
      description: "List all bookmarks in the document. WHEN: need to see available bookmarks for navigation. NOT: want to add a bookmark? use word_add_bookmark.",
    },
    async () => {
      const bookmarks = await docOps.getBookmarks()
      if (bookmarks.length === 0) {
        return `Action: No bookmarks found\nDetail: The document has no bookmarks`
      }
      return `Action: ${bookmarks.length} bookmark(s) found\nDetail:\n${bookmarks.map(b =>
        `  #${b.index}: "${b.name}"`
      ).join("\n")}`
    },
  )

  regTool("word_get_lists",
    {
      description: "List all bullet and numbered lists in the document with hierarchy. WHEN: need to review list structure or verify list content. NOT: want raw text? use word_get_text.",
    },
    async () => {
      const result = await docOps.getLists()
      if (result.listCount === 0) {
        return `Action: No lists found\nDetail: The document has no lists`
      }
      const lines: string[] = []
      for (let i = 0; i < result.lists.length; i++) {
        const list = result.lists[i]
        lines.push(`  List ${i + 1} (${list.type}, ${list.items.length} items):`)
        for (const item of list.items) {
          const indent = "  ".repeat(item.level - 1)
          const trimmed = item.text.slice(0, 150)
          lines.push(`    ${indent}${item.prefix} ${trimmed}${item.text.length > 150 ? "..." : ""}`)
        }
      }
      return `Action: ${result.listCount} list(s) found\nDetail:\n${lines.join("\n")}`
    },
  )

  regTool("word_get_sections",
    {
      description: "List all sections with page setup info (orientation, columns, page size). WHEN: need to understand document layout or section boundaries. NOT: need heading structure? use word_get_structure.",
    },
    async () => {
      const result = await docOps.getSections()
      const lines = result.sections.map(s =>
        `  Section ${s.index}: ${s.orientation}, ${s.columnCount} column(s), ${s.pageWidth.toFixed(0)}×${s.pageHeight.toFixed(0)}pt`
      )
      return `Action: ${result.count} section(s) found\nDetail:\n${lines.join("\n")}`
    },
  )
}
