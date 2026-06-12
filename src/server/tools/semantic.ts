import { z } from "zod"
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js"
import { WordTextEditor } from "../../word/word-text-editor.js"
import { WordTableEditor } from "../../word/word-table-editor.js"
import { PositionMap, type Locator, type ResolvedPosition } from "../../word/position-map.js"
import { SecurityManager } from "../../security/policy.js"
import { mcpCall } from "./helper.js"

function locatorFromArgs(args: Record<string, unknown>): Locator {
  const by = (args.by as string) ?? "heading"
  const base: Record<string, unknown> = { by }

  if (by === "heading" || by === "paragraph") {
    if (args.match != null) base.match = args.match
    if (args.matchMode != null) base.matchMode = args.matchMode
    if (args.occurrence != null) base.occurrence = args.occurrence
  }

  if (by === "table") {
    if (args.occurrence != null) base.occurrence = args.occurrence
  }

  if (by === "bookmark") {
    base.name = args.name
  }

  if (args.offsetDirection != null || args.offsetCount != null) {
    base.offset = {
      direction: (args.offsetDirection as string) ?? "after",
      count: (args.offsetCount as number) ?? 1,
    }
  }

  return base as unknown as Locator
}

function formatPosition(pos: ResolvedPosition): string {
  if (!pos.found) return `Location not found: ${pos.error ?? "unknown reason"}`
  const parts = [`📍 Paragraph ${pos.paragraphIndex}`]
  if (pos.headingContext) parts.push(`Under: ${pos.headingContext}`)
  if (pos.tableIndex != null) parts.push(`Table #${pos.tableIndex}`)
  return parts.join(" | ")
}

const locatorFields = {
  by: z.enum(["heading", "paragraph", "table", "bookmark"]).optional().describe("Target type (default: heading)"),
  match: z.string().max(5000).optional().describe("Text to match (for heading/paragraph)"),
  matchMode: z.enum(["exact", "contains", "startsWith", "regex"]).optional().describe("Matching mode (default: exact)"),
  occurrence: z.number().int().min(1).max(1000).optional().describe("Which occurrence to target (1-based, default: 1)"),
  offsetDirection: z.enum(["before", "after"]).optional().describe("Offset direction from the matched element"),
  offsetCount: z.number().int().min(1).max(1000).optional().describe("Number of paragraphs to offset (default: 1)"),
  name: z.string().min(1).max(255).optional().describe("Bookmark name (required when by='bookmark')"),
} as const

export function registerSemanticTools(
  server: McpServer,
  textEditor: WordTextEditor,
  tableEditor: WordTableEditor,
  positionMap: PositionMap,
  security: SecurityManager,
): void {
  server.registerTool(
    "word_locate",
    {
      description: "Resolve a semantic location in the document and return position info. WHEN: need to find where a heading/paragraph/table/bookmark is before editing. NOT: want to move cursor? use word_select_at instead.",
      inputSchema: locatorFields,
    },
    mcpCall(security, "word_locate", async (args) => {
      const locator = locatorFromArgs(args)
      const pos = await positionMap.resolve(locator)
      if (!pos.found) {
        return `Action: Locate failed\nDetail: ${pos.error}`
      }
      const lines = [
        `Action: Location resolved`,
        `Detail: ${formatPosition(pos)}`,
      ]
      if (pos.tableIndex != null) {
        lines.push(`Next: word_select_at({by:'table', occurrence:${pos.tableIndex}}) to move cursor, or word_edit_cell_at({table:{by:'table', occurrence:${pos.tableIndex}}, row:1, column:1, text:'...'})`)
      } else {
        lines.push(`Next: word_select_at(${JSON.stringify(args)}) to move cursor`)
      }
      return lines.join("\n")
    }),
  )

  server.registerTool(
    "word_select_at",
    {
      description: "Move cursor to a semantic location (heading/paragraph/table/bookmark). WHEN: need to navigate precisely without calculating paragraph numbers. NOT: just want to check location? use word_locate.",
      inputSchema: locatorFields,
    },
    mcpCall(security, "word_select_at", async (args) => {
      const locator = locatorFromArgs(args)
      const pos = await positionMap.resolve(locator)
      if (!pos.found) {
        return `Action: Select failed\nDetail: ${pos.error}`
      }
      await textEditor.goToParagraph(pos.paragraphIndex)
      return `Action: Cursor moved\nDetail: ${formatPosition(pos)}`
    }),
  )

  server.registerTool(
    "word_insert_at",
    {
      description: "Insert text at a semantic location. WHEN: need to write content at a specific heading/paragraph without calculating indices. NOT: just want to move cursor? use word_select_at.",
      inputSchema: {
        ...locatorFields,
        text: z.string().max(1000000).describe("Text to insert"),
        mode: z.enum(["smooth", "instant"]).optional().describe("'smooth' (default) splits into sentence chunks; 'instant' writes all at once"),
      },
    },
    mcpCall(security, "word_insert_at", async (args) => {
      const locator = locatorFromArgs(args)
      const pos = await positionMap.resolve(locator)
      if (!pos.found) {
        return `Action: Insert failed\nDetail: ${pos.error}`
      }
      await textEditor.goToParagraph(pos.paragraphIndex)
      await textEditor.typeText(args.text as string, args.mode as "smooth" | "instant" | undefined)
      const preview = (args.text as string).slice(0, 80) + ((args.text as string).length > 80 ? "..." : "")
      return `Action: Text inserted\nDetail: "${preview}" (${(args.text as string).length} chars) at ${formatPosition(pos)}`
    }),
  )

  server.registerTool(
    "word_edit_cell_at",
    {
      description: "Edit a table cell by first locating the table semantically. WHEN: need to edit a specific cell but don't know the table index. NOT: already know the table index? use word_edit_cell.",
      inputSchema: {
        ...locatorFields,
        row: z.number().int().min(1).max(1000).describe("Row number (1-based)"),
        column: z.number().int().min(1).max(1000).describe("Column number (1-based)"),
        text: z.string().max(100000).describe("New cell text"),
      },
    },
    mcpCall(security, "word_edit_cell_at", async (args) => {
      const locator = locatorFromArgs(args)
      const pos = await positionMap.resolve(locator)
      if (!pos.found) {
        return `Action: Edit cell failed\nDetail: ${pos.error}`
      }
      const tableIdx = pos.tableIndex ?? 1
      await tableEditor.editTableCell({
        tableIndex: tableIdx,
        row: args.row as number,
        column: args.column as number,
        text: args.text as string,
      })
      const preview = (args.text as string).slice(0, 50) + ((args.text as string).length > 50 ? "..." : "")
      return `Action: Cell (${args.row},${args.column}) in table at ${formatPosition(pos)} updated\nDetail: "${preview}"`
    }),
  )
}
