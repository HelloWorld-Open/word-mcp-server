import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js"
import { SecurityManager } from "../../security/policy.js"
import type { ServerContext } from "../server-context.js"
import type { IWordSession } from "../../word/session.js"
import type { PositionMap } from "../../word/position-map.js"
import { createRegTool } from "./shared.js"

const wdWithinTable = 26
const wdStartOfRangeRowNumber = 16
const wdStartOfRangeColumnNumber = 17

export function registerWhereAmITool(
  server: McpServer,
  context: ServerContext,
  positionMap: PositionMap,
  security: SecurityManager,
): void {
  const regTool = createRegTool(server, security, context)

  regTool("word_where_am_i",
    {
      description: "WHEN: need to verify current cursor location after navigating with word_select_at, word_find_text, or word_go_to_paragraph. WHAT: returns the semantic heading path, position relative to nearest heading, table context (if inside a table), and document stats. CONSTRAINT: read-only; does NOT modify cursor or document. For document-level heading overview, use word_get_structure.",
      inputSchema: {},
    },
    async () => {
      const session = context.session as IWordSession
      await positionMap.ensureFresh()

      const sel = session.getSelectionProxy()
      const range = sel.getRange()
      const cursorStart = range.getStart()
      const paraIndex = positionMap.getParagraphIndex(cursorStart)
      const paraCount = positionMap.cachedParaCount

      const path = positionMap.getHeadingPath(paraIndex)
      const lastHeading = path.length > 0 ? path[path.length - 1] : null

      const inTable = sel.getInformation(wdWithinTable)
      let tableRow = 0, tableCol = 0
      if (inTable) {
        tableRow = sel.getInformation(wdStartOfRangeRowNumber) as number
        tableCol = sel.getInformation(wdStartOfRangeColumnNumber) as number
      }

      const lines: string[] = []
      lines.push(`Action: Cursor position resolved`)
      if (path.length > 0) {
        const pathStr = path.map(h => h.text).join(" > ")
        lines.push(`    Path: ${pathStr}`)
        const offset = paraIndex - lastHeading!.paragraphIndex
        if (offset === 0) {
          lines.push(`    Position: at heading "${lastHeading!.text}" (H${lastHeading!.level})`)
        } else {
          lines.push(`    Position: paragraph ${offset} under "${lastHeading!.text}" (H${lastHeading!.level})`)
        }
      }
      lines.push(`    Absolute: para ${paraIndex} / ${paraCount}`)
      if (inTable) lines.push(`    In table: yes (row ${tableRow}, column ${tableCol})`)
      else lines.push(`    In table: no`)
      lines.push(`    Headings: ${positionMap.cachedHeadingCount} | Tables: ${positionMap.cachedTableCount} | Doc v: ${positionMap.docVersion}`)
      lines.push(`    Next: word_select_at()/word_locate()/word_insert_at()`)

      return lines.join("\n")
    },
  )
}
