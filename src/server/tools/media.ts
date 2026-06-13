import { z } from "zod"
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js"
import { WordMediaEditor } from "../../word/word-media-editor.js"
import { SecurityManager } from "../../security/policy.js"
import type { ServerContext } from "../server-context.js"
import { mcpCall } from "./helper.js"

export function registerMediaTools(
  server: McpServer,
  context: ServerContext,
  content: WordMediaEditor,
  security: SecurityManager,
): void {
  server.registerTool(
    "word_insert_image",
    {
      description: "Insert an image at the cursor position. WHEN: need to embed a picture from disk into the document. NOT: want to create a chart? use word_insert_chart.",
      inputSchema: {
        imagePath: z.string().min(1).max(4096).describe("Full path to the image file (jpg, png, gif, bmp)"),
        width: z.number().positive().max(1000).optional().describe("Width in points"),
        height: z.number().positive().max(1000).optional().describe("Height in points"),
      },
    },
    mcpCall(security, context, "word_insert_image", async ({ imagePath, width, height }) => {
      const safePath = security.pathSanitizer.validateForRead(imagePath)
      await content.insertImage({ imagePath: safePath, width, height })
      const dims = width && height ? `${width}x${height}pt` : "original size"
      return `Action: Image inserted (${dims})\nNext: word_type_text({text:"Figure caption..."})`
    }),
  )

  server.registerTool(
    "word_insert_chart",
    {
      description: "Insert a chart with data (column/bar/line/pie/area).",
      inputSchema: {
        type: z.enum(["column", "bar", "line", "pie", "area"]).describe("Chart type"),
        data: z.array(z.array(z.union([z.string(), z.number()])).min(1)).min(1).max(100).describe("2D data: first row = headers, first column = categories"),
        title: z.string().max(255).optional().describe("Chart title"),
        width: z.number().positive().max(1000).optional().describe("Width in points (default: 400)"),
        height: z.number().positive().max(1000).optional().describe("Height in points (default: 250)"),
      },
    },
    mcpCall(security, context, "word_insert_chart", async ({ type, data, title, width, height }) => {
      const result = await content.insertChart({ type, data, title, width, height })
      return `Action: Chart inserted (${result.type})\nDetail: ${result.series} series, title: ${title ?? "none"}\nNext: word_type_text({text:"Chart shows..."})`
    }),
  )

  server.registerTool(
    "word_insert_textbox",
    {
      description: "Insert a floating text box. WHEN: need positioned overlay text outside normal flow. NOT: want inline text? use word_type_text.",
      inputSchema: {
        text: z.string().min(1).max(10000).describe("Text content"),
        width: z.number().positive().max(1000).optional().describe("Width in points (default: 200)"),
        height: z.number().positive().max(1000).optional().describe("Height in points (default: 100)"),
        orientation: z.enum(["horizontal", "vertical"]).optional().describe("Text orientation"),
        positionLeft: z.number().min(0).max(2000).optional().describe("Left position in points (default: 50)"),
        positionTop: z.number().min(0).max(2000).optional().describe("Top position in points (default: 50)"),
      },
    },
    mcpCall(security, context, "word_insert_textbox", async ({ text, width, height, orientation, positionLeft, positionTop }) => {
      const result = await content.insertTextbox({ text, width, height, orientation, positionLeft, positionTop })
      return `Action: Text box inserted (${result.width}x${result.height}pt)\nOrientation: ${orientation ?? "horizontal"}\nNext: word_type_text({text:"..."}) or word_insert_textbox({text:"Another box", positionLeft:200})`
    }),
  )
}
