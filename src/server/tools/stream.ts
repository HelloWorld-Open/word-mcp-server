import { z } from "zod"
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js"
import { SecurityManager } from "../../security/policy.js"
import type { ServerContext } from "../server-context.js"
import type { StreamingMarkdownWriter } from "../../word/word-stream-writer.js"
import { mcpCall } from "./helper.js"

export function registerStreamTools(
  server: McpServer,
  context: ServerContext,
  streamWriter: StreamingMarkdownWriter,
  security: SecurityManager,
): void {
  server.registerTool(
    "word_stream_start",
    {
      description: "启动一个流式文档会话，创建新文档并准备逐块写入。自动关闭当前活动文档。支持空白文档和模板(.dotx)创建。支持通过 baseStyleProfile 预定义内置样式（Normal、Heading 1 等），markdown 写入时自动继承样式格式，零额外 COM 开销。",
      inputSchema: {
        title: z.string().max(255).optional().describe("文档标题"),
        author: z.string().max(255).optional().describe("文档作者"),
        templatePath: z.string().max(4096).optional().describe(".dotx 或 .dotm 模板文件完整路径"),
        orientation: z.enum(["portrait", "landscape"]).optional().describe("页面方向"),
        topMargin: z.number().min(0).max(100).optional().describe("上边距 (cm)"),
        bottomMargin: z.number().min(0).max(100).optional().describe("下边距 (cm)"),
        leftMargin: z.number().min(0).max(100).optional().describe("左边距 (cm)"),
        rightMargin: z.number().min(0).max(100).optional().describe("右边距 (cm)"),
        baseStyleProfile: z.record(
          z.string().max(100),
          z.object({
            font: z.object({
              name: z.string().max(100).optional().describe("字体名称"),
              size: z.number().min(1).max(1638).optional().describe("字号 (pt)"),
              bold: z.boolean().optional().describe("加粗"),
              italic: z.boolean().optional().describe("斜体"),
              color: z.enum(["auto", "black", "blue", "turquoise", "bright_green", "pink", "red", "yellow", "white", "dark_blue", "teal", "green", "violet", "dark_red", "dark_yellow", "gray_50", "gray_25"]).optional().describe("字体颜色"),
              underline: z.enum(["none", "single", "double", "wavy"]).optional().describe("下划线"),
              strikethrough: z.boolean().optional().describe("删除线"),
              highlight: z.string().max(20).optional().describe("高亮色 (17 色枚举名 或 #RRGGBB hex)"),
            }).optional(),
            paragraph: z.object({
              alignment: z.enum(["left", "center", "right", "justify"]).optional().describe("对齐方式"),
              firstLineIndent: z.number().min(-100).max(100).optional().describe("首行缩进 (cm)"),
              spaceBefore: z.number().min(0).max(1584).optional().describe("段前间距 (pt)"),
              spaceAfter: z.number().min(0).max(1584).optional().describe("段后间距 (pt)"),
              lineSpacing: z.number().min(0).max(1584).optional().describe("行距值 (pt/倍数)"),
              lineSpacingRule: z.enum(["single", "one_point_five", "double", "at_least", "exactly", "multiple"]).optional().describe("行距规则"),
              keepWithNext: z.boolean().optional().describe("与下段同页"),
              pageBreakBefore: z.boolean().optional().describe("段前分页"),
              borders: z.object({
                style: z.enum(["none", "single", "dot", "dash", "double"]).describe("边框线型"),
                color: z.enum(["auto", "black", "blue", "turquoise", "bright_green", "pink", "red", "yellow", "white", "dark_blue", "teal", "green", "violet", "dark_red", "dark_yellow", "gray_50", "gray_25"]).optional().describe("边框颜色"),
                size: z.number().min(1).max(48).optional().describe("线宽 (1/4pt)"),
                sides: z.array(z.enum(["top", "bottom", "left", "right"])).optional().describe("应用到的边，默认四边"),
              }).optional(),
            }).optional(),
          }),
        ).optional().describe("内置样式配置，如 Normal、Heading 1 等。在文档创建时修改样式定义，所有应用该样式的内容自动继承格式"),
      },
    },
    mcpCall(security, context, "word_stream_start", async (args) => {
      const safeArgs = args.templatePath
        ? { ...args, templatePath: security.pathSanitizer.resolveAndValidate(args.templatePath) }
        : args
      return await streamWriter.start(safeArgs)
    }, { preconditions: [] }),
  )

  server.registerTool(
    "word_stream_block",
    {
      description: "写入一个或多个 markdown 内容块到当前流式文档中，内容即时在 Word 窗口中呈现。必须在 word_stream_start 之后调用。支持标题、段落、列表、表格、代码块、引用等所有 markdown 语法。",
      inputSchema: {
        text: z.string().min(1).max(100000).describe("Markdown 内容（单个或多个块，建议按自然章节分批发送）"),
      },
    },
    mcpCall(security, context, "word_stream_block", async (args) => {
      const result = await streamWriter.writeBlock(args.text)
      return `已写入 ${result.chars} 字符（${result.blockType}），累计 ${result.blockIndex} 块`
    }),
  )

  server.registerTool(
    "word_stream_end",
    {
      description: "结束流式文档会话，保存文档并可选导出为 PDF。保存后流式会话终止。",
      inputSchema: {
        save: z.boolean().optional().describe("是否保存文档（默认 true）"),
        exportPath: z.string().max(4096).optional().describe("可选 PDF 导出路径"),
      },
    },
    mcpCall(security, context, "word_stream_end", async (args) => {
      const safePath = args.exportPath
        ? security.pathSanitizer.validateForWrite(args.exportPath)
        : undefined
      const result = await streamWriter.end({ save: args.save, exportPath: safePath })
      const lines: string[] = [`流式会话已结束`]
      lines.push(`累计写入 ${result.blockCount} 块，${result.charCount} 字符，耗时 ${result.elapsed}ms`)
      if (result.savedPath) lines.push(`保存至: ${result.savedPath}`)
      if (result.pdfPath) lines.push(`PDF 导出至: ${result.pdfPath}`)
      return lines.join("\n")
    }),
  )
}
