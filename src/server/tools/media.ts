import { z } from "zod"
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js"
import { WordContentWriter } from "../../word/word-content-writer.js"
import { SecurityManager } from "../../security/policy.js"
import type { ServerContext } from "../server-context.js"
import { createRegTool } from "./shared.js"

export function registerMediaTools(
  server: McpServer,
  context: ServerContext,
  content: WordContentWriter,
  security: SecurityManager,
): void {
  const regTool = createRegTool(server, security, context)
  regTool("word_insert_image",
    {
      description: "Insert an image at the cursor position. WHEN: need to embed a picture from disk into the document. NOT: want to create a chart? use word_insert_chart.",
      inputSchema: {
        imagePath: z.string().min(1).max(4096).describe("Full path to the image file (jpg, png, gif, bmp)"),
        width: z.number().positive().max(1000).optional().describe("Width in points"),
        height: z.number().positive().max(1000).optional().describe("Height in points"),
        quiet: z.boolean().optional().describe("简洁输出模式"),
      },
    },
    async ({ imagePath, width, height, quiet }) => {
      const safePath = security.pathSanitizer.resolveAndValidate(imagePath)
      await content.insertImage({ imagePath: safePath, width, height })
      if (quiet) return "Image inserted"
      const dims = width && height ? `${width}x${height}pt` : "original size"
      return `Action: Image inserted (${dims})\nNext: word_type_text({text:"Figure caption..."})`
    },
  )

  regTool("word_insert_chart",
    {
      description: "Insert a chart with data (column/bar/line/pie/area). WHEN: need to visualize data as a chart. NOT: want to insert an image file from disk? use word_insert_image.",
      inputSchema: {
        type: z.enum(["column", "bar", "line", "pie", "area"]).describe("Chart type"),
        data: z.array(z.array(z.union([z.string(), z.number()])).min(1)).min(1).max(100).describe("2D data: first row = headers, first column = categories"),
        title: z.string().max(255).optional().describe("Chart title"),
        width: z.number().positive().max(1000).optional().describe("Width in points (default: 400)"),
        height: z.number().positive().max(1000).optional().describe("Height in points (default: 250)"),
        quiet: z.boolean().optional().describe("简洁输出模式"),
      },
    },
    async ({ type, data, title, width, height, quiet }) => {
      const result = await content.insertChart({ type, data, title, width, height })
      if (quiet) return `Chart inserted (${result.type}, ${result.series} series)`
      return `Action: Chart inserted (${result.type})\nDetail: ${result.series} series, title: ${title ?? "none"}\nNext: word_type_text({text:"Chart shows..."})`
    },
  )

  regTool("word_insert_textbox",
    {
      description: "Insert a floating text box. WHEN: need positioned overlay text outside normal flow. NOT: want inline text? use word_type_text.",
      inputSchema: {
        text: z.string().min(1).max(10000).describe("Text content"),
        width: z.number().positive().max(1000).optional().describe("Width in points (default: 200)"),
        height: z.number().positive().max(1000).optional().describe("Height in points (default: 100)"),
        orientation: z.enum(["horizontal", "vertical"]).optional().describe("Text orientation"),
        positionLeft: z.number().min(0).max(2000).optional().describe("Left position in points (default: 50)"),
        positionTop: z.number().min(0).max(2000).optional().describe("Top position in points (default: 50)"),
        quiet: z.boolean().optional().describe("简洁输出模式"),
      },
    },
    async ({ text, width, height, orientation, positionLeft, positionTop, quiet }) => {
      const result = await content.insertTextbox({ text, width, height, orientation, positionLeft, positionTop })
      if (quiet) return `Textbox inserted (${result.width}x${result.height})`
      return `Action: Text box inserted (${result.width}x${result.height}pt)\nOrientation: ${orientation ?? "horizontal"}\nNext: word_type_text({text:"..."}) or word_insert_textbox({text:"Another box", positionLeft:200})`
    },
  )
}
