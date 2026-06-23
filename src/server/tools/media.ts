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
      description: "WHEN: need to embed a picture from disk into the document. WHAT: inserts an image (jpg, png, gif, bmp) at the cursor position at optional dimensions. CONSTRAINT: file must exist on disk; only raster formats supported (no SVG). The image is embedded (copied into .docx), not linked.",
      inputSchema: {
        imagePath: z.string().min(1).max(4096).describe("Full path to the image file (supported: .jpg, .png, .gif, .bmp)"),
        width: z.number().positive().max(1000).optional().describe("Width in points (default: original image width). 72pt = 1 inch."),
        height: z.number().positive().max(1000).optional().describe("Height in points (default: original image height). Omit to preserve aspect ratio."),
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
      description: "WHEN: need to visualize data as a chart (column, bar, line, pie, area). WHAT: inserts a chart with 2D data array at cursor position. CONSTRAINT: first row = column headers, first column = row labels. Chart data is embedded and editable in Word. For inserting image files from disk, use word_insert_image.",
      inputSchema: {
        type: z.enum(["column", "bar", "line", "pie", "area"]).describe("Chart type: 'column'=vertical bars, 'bar'=horizontal bars, 'line'=line graph, 'pie'=circular, 'area'=filled line"),
        data: z.array(z.array(z.union([z.string(), z.number()])).min(1)).min(1).max(100).describe("2D data table. First row = headers (strings). First column = categories/labels. Example for sales by quarter: [['Q1','Q2','Q3'], [100, 200, 150]]"),
        title: z.string().max(255).optional().describe("Chart title shown above the chart"),
        width: z.number().positive().max(1000).optional().describe("Width in points (default: 400). About 14cm at 72dpi."),
        height: z.number().positive().max(1000).optional().describe("Height in points (default: 250). About 9cm at 72dpi."),
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
      description: "WHEN: need positioned overlay text outside the normal document flow (e.g., callout, sidebar, annotation). WHAT: inserts a floating text box at specified position and dimensions. CONSTRAINT: text box is positioned absolutely (not inline with text). For normal inline content, use word_stream_block or word_type_text.",
      inputSchema: {
        text: z.string().min(1).max(10000).describe("Text content inside the text box (plain text only)"),
        width: z.number().positive().max(1000).optional().describe("Width in points (default: 200). About 7cm at 72dpi."),
        height: z.number().positive().max(1000).optional().describe("Height in points (default: 100). About 3.5cm at 72dpi."),
        orientation: z.enum(["horizontal", "vertical"]).optional().describe("Text orientation inside the box (default: horizontal)"),
        positionLeft: z.number().min(0).max(2000).optional().describe("Left position from page edge in points (default: 50)"),
        positionTop: z.number().min(0).max(2000).optional().describe("Top position from page edge in points (default: 50)"),
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
