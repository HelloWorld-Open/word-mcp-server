import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js"
import type { ServerContext } from "../server-context.js"
import type { WordDocument } from "../../word/document.js"
import type { PositionMap } from "../../word/position-map.js"
import { WordMcpError } from "../../security/errors.js"

function activeSession(context: ServerContext) {
  const { session } = context
  if (!session || !session.activeDoc) return null
  return session
}

function errorResponse(uri: string, msg: string) {
  return {
    contents: [{ uri, mimeType: "application/json" as const, text: JSON.stringify({ error: msg }) }],
  }
}

export function registerDocumentStructureResource(
  server: McpServer,
  context: ServerContext,
  docOps: WordDocument,
  positionMap: PositionMap,
): void {
  server.registerResource(
    "document-structure",
    "doc://structure",
    {
      description: "Document structure summary. Use sub-resources (doc://structure/headings) for detailed heading tree.",
      mimeType: "application/json",
    },
    async () => {
      const session = activeSession(context)
      if (!session) return errorResponse("doc://structure", "No document open")

      try {
        await positionMap.ensureFresh()
        const info = docOps.getInfo()
        const sections = docOps.getSections()
        const docVersion = positionMap.docVersion

        return {
          contents: [{
            uri: "doc://structure",
            mimeType: "application/json",
            text: JSON.stringify({
              version: docVersion,
              document: {
                name: info.name,
                wordCount: info.wordCount,
                paragraphCount: info.paragraphCount,
                pageCount: info.pageCount,
                sectionCount: info.sectionCount,
                characterCount: info.characterCount,
                saved: info.saved,
              },
              counts: {
                headings: positionMap.cachedHeadingCount,
                tables: positionMap.cachedTableCount,
                sections: info.sectionCount,
              },
              _links: {
                self: "doc://structure",
                headings: "doc://structure/headings",
              },
            }),
          }],
        }
      } catch (err) {
        const msg = err instanceof WordMcpError ? err.message : String(err)
        return errorResponse("doc://structure", msg)
      }
    },
  )

  server.registerResource(
    "document-headings",
    "doc://structure/headings",
    {
      description: "Full document heading hierarchy tree",
      mimeType: "application/json",
    },
    async () => {
      const session = activeSession(context)
      if (!session) return errorResponse("doc://structure/headings", "No document open")

      try {
        await positionMap.ensureFresh()
        const headings = positionMap.getHeadings()

        return {
          contents: [{
            uri: "doc://structure/headings",
            mimeType: "application/json",
            text: JSON.stringify({
              version: positionMap.docVersion,
              total: headings.length,
              headings: headings.map(h => ({
                level: h.level,
                text: h.text,
                paragraphIndex: h.paragraphIndex,
              })),
            }),
          }],
        }
      } catch (err) {
        const msg = err instanceof WordMcpError ? err.message : String(err)
        return errorResponse("doc://structure/headings", msg)
      }
    },
  )
}
