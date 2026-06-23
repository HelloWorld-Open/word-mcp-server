import { z } from "zod"
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js"
import { WordDocument, type TextSummary } from "../../word/document.js"
import { SecurityManager } from "../../security/policy.js"
import type { ServerContext } from "../server-context.js"
import { createRegTool } from "./shared.js"
import type { ToolResponse } from "../tool-result.js"

function isTextSummary(v: string | TextSummary): v is TextSummary {
  return typeof v === "object" && v !== null && "hasMore" in v
}

export function registerReaderTools(
  server: McpServer,
  context: ServerContext,
  docOps: WordDocument,
  security: SecurityManager,
): void {
  const regTool = createRegTool(server, security, context)
  regTool("word_get_text",
    {
      description: "WHEN: need to read the entire document content to understand what's written. WHAT: returns the full plain text of the current document. CONSTRAINT: read-only. For documents over 10,000 characters, a summary with pagination info is returned instead. For heading structure, use word_get_structure.",
    },
    async (): Promise<string | ToolResponse> => {
      const result = await docOps.getFullText()
      if (isTextSummary(result)) {
        return {
          text: `Document text (${result.totalChars} total chars — showing first ${result.returnedChars}):\n\n${result.text}`,
          data: {
            textSummary: {
              totalChars: result.totalChars,
              returnedChars: result.returnedChars,
              paraCount: result.paraCount,
              headingCount: result.headingCount,
              hasMore: result.hasMore,
              nextAction: result.nextAction,
            },
          },
        }
      }
      return result
    },
  )

  regTool("word_get_paragraph",
    {
      description: "WHEN: need to read a specific paragraph's text when you know its index (e.g., from word_get_structure). WHAT: returns the text content of a single paragraph by 1-based index with preview truncation at 200 chars. CONSTRAINT: read-only. For full document content, use word_get_text.",
      inputSchema: {
        index: z.number().int().positive().describe("Paragraph index (1-based). Use word_get_structure() output like 'H1 ¶3 — Introduction' to find paragraph indices."),
      },
    },
    async ({ index }) => {
      const text = await docOps.getParagraphText(index)
      return `Action: Paragraph ${index}\nDetail: "${text.slice(0, 200)}${text.length > 200 ? "..." : ""}"\nNext: word_get_paragraph({index:${index + 1}}) or word_get_text()`
    },
  )

  regTool("word_export_to_pdf",
    {
      description: "WHEN: need a PDF copy of the current document for sharing or preview without changing the original docx. WHAT: exports the current document to PDF at the specified path. CONSTRAINT: does NOT change the active document or its save state. For saving in other formats (RTF, TXT, HTML), use word_save_as.",
      inputSchema: {
        path: z.string().min(1).max(4096).optional().describe("Output PDF path (default: same name as source document with .pdf extension). Required if document is untitled."),
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
      description: "WHEN: need to read table content as structured data to verify or analyze it. WHAT: extracts all rows and columns of a table by index (1-based) with table count info. CONSTRAINT: read-only. For editing table content, use word_edit_cell or word_edit_cells.",
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
      description: "WHEN: need to review existing comments/feedback in the document. WHAT: lists all comments with index, author, and text preview. CONSTRAINT: read-only. For adding new comments, use word_add_comment.",
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
      description: "WHEN: need to see available bookmarks for navigation or hyperlink targets. WHAT: lists all bookmarks with index and name. CONSTRAINT: read-only. For adding new bookmarks, use word_add_bookmark.",
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
      description: "WHEN: need to review bullet/numbered list structure or verify list content. WHAT: returns all lists with hierarchy (indentation level), item text, and item prefix (bullet/number). CONSTRAINT: read-only. For creating new lists, use word_insert_list.",
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
      description: "WHEN: need to understand document layout boundaries or section formatting. WHAT: returns all sections with orientation, column count, and page dimensions. CONSTRAINT: read-only. For changing section layout, use word_set_page_setup or word_insert_section_break.",
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
