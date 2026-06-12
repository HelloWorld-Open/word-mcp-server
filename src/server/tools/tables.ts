import { z } from "zod"
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js"
import { WordTableEditor } from "../../word/word-table-editor.js"
import { SecurityManager } from "../../security/policy.js"
import { mcpCall } from "./helper.js"

export function registerTableTools(
  server: McpServer,
  content: WordTableEditor,
  security: SecurityManager,
): void {
  server.registerTool(
    "word_insert_table",
    {
      description: "Insert a table. WHEN: cursor is where the table should appear. NOT: want to fill cells immediately? pass data param.",
      inputSchema: {
        rows: z.number().int().min(1).max(500).describe("Number of rows"),
        columns: z.number().int().min(1).max(100).describe("Number of columns"),
        data: z.array(z.array(z.string().max(100000)).max(100)).max(1000).optional().describe("Optional 2D cell data"),
        autoFitBehavior: z.enum(["fixed", "contents", "window"]).optional().describe("Auto-fit behavior"),
      },
    },
    mcpCall(security, "word_insert_table", async ({ rows, columns, data, autoFitBehavior }) => {
      const result = await content.insertTable({ rows, columns, data, autoFitBehavior })
      const failInfo = result.failCount > 0 ? `\nWarning: ${result.failCount} cell(s) failed to write` : ""
      return `Action: Table inserted (${result.rows}x${result.columns})${failInfo}\nData rows: ${data?.length ?? 0}\nNext: word_edit_cell({row:1, column:1, text:"..."}) or word_edit_cells({data:[[\"a\",\"b\"]]}) or word_apply_table_style({styleName:"Light List Accent 1"})`
    }),
  )

  server.registerTool(
    "word_edit_cell",
    {
      description: "Edit a single table cell's text.",
      inputSchema: {
        tableIndex: z.number().int().min(1).optional().describe("Table index (1-based, default: 1)"),
        row: z.number().int().min(1).describe("Row number (1-based)"),
        column: z.number().int().min(1).describe("Column number (1-based)"),
        text: z.string().max(100000).describe("New cell text"),
      },
    },
    mcpCall(security, "word_edit_cell", async ({ tableIndex, row, column, text }) => {
      await content.editTableCell({ tableIndex, row, column, text })
      const preview = text.slice(0, 50) + (text.length > 50 ? "..." : "")
      return `Action: Cell (${row},${column}) updated\nDetail: "${preview}"\nNext: word_edit_cell({row:${row+1}, column:${column}, text:"..."}) or word_set_cell_font({row:${row}, column:${column}, bold:true})`
    }),
  )

  server.registerTool(
    "word_edit_cells",
    {
      description: "Batch-fill multiple table cells with a 2D data array.",
      inputSchema: {
        tableIndex: z.number().int().min(1).optional().describe("Table index (1-based, default: 1)"),
        data: z.array(z.array(z.string().max(100000)).max(100)).min(1).max(1000).describe("2D array of cell data (row-major)"),
      },
    },
    mcpCall(security, "word_edit_cells", async ({ tableIndex, data }) => {
      const result = await content.editCells({ tableIndex, data })
      const failInfo = result.failCount > 0 ? `, ${result.failCount} cell(s) failed` : ""
      return `Action: Batch-filled ${result.rows}x${result.columns} table${failInfo}\nDetail: ${data.length} row(s) written\nNext: word_set_cell_font({row:1, column:1, bold:true}) or word_edit_cell({row:1, column:1, text:"..."})`
    }),
  )

  server.registerTool(
    "word_add_table_row",
    {
      description: "Add a new row to a table.",
      inputSchema: {
        tableIndex: z.number().int().min(1).optional().describe("Table index (1-based, default: 1)"),
        data: z.array(z.string().max(100000)).max(100).optional().describe("Cell data for the new row"),
      },
    },
    mcpCall(security, "word_add_table_row", async ({ tableIndex, data }) => {
      await content.addTableRow({ tableIndex, data })
      return `Action: Row added\nDetail: ${data?.length ?? 0} cells\nNext: word_edit_cell({row:1, column:1, text:"..."}) or word_set_row_height({row:1, height:30})`
    }),
  )

  server.registerTool(
    "word_delete_table_row",
    {
      description: "Delete a row from a table.",
      inputSchema: {
        tableIndex: z.number().int().min(1).optional().describe("Table index (1-based, default: 1)"),
        rowIndex: z.number().int().min(1).describe("Row number to delete (1-based)"),
      },
    },
    mcpCall(security, "word_delete_table_row", async ({ tableIndex, rowIndex }) => {
      await content.deleteTableRow({ tableIndex, rowIndex })
      return `Action: Row ${rowIndex} deleted\nNext: word_undo({count:1}) or word_add_table_row({data:["a","b"]})`
    }),
  )

  server.registerTool(
    "word_set_table_borders",
    {
      description: "Set table border style, color, and width. WHEN: table exists and needs visual styling. NOT: want background color? use word_set_table_shading.",
      inputSchema: {
        tableIndex: z.number().int().min(1).optional().describe("Table index (1-based, default: 1)"),
        inside: z.object({
          style: z.enum(["none", "single", "dot", "dash_small", "dash_large", "dash", "dash_dot", "double"]).optional().describe("Border line style"),
          color: z.enum(["auto", "black", "blue", "turquoise", "bright_green", "pink", "red", "yellow", "white", "dark_blue", "teal", "green", "violet", "dark_red", "dark_yellow", "gray_50", "gray_25"]).optional().describe("Border color"),
          size: z.number().int().min(2).max(48).optional().describe("Line width in 1/4pt (8=1pt)"),
        }).optional(),
        outside: z.object({
          style: z.enum(["none", "single", "dot", "dash_small", "dash_large", "dash", "dash_dot", "double"]).optional().describe("Border line style"),
          color: z.enum(["auto", "black", "blue", "turquoise", "bright_green", "pink", "red", "yellow", "white", "dark_blue", "teal", "green", "violet", "dark_red", "dark_yellow", "gray_50", "gray_25"]).optional().describe("Border color"),
          size: z.number().int().min(2).max(48).optional().describe("Line width in 1/4pt (8=1pt)"),
        }).optional(),
      },
    },
    mcpCall(security, "word_set_table_borders", async ({ tableIndex, inside, outside }) => {
      await content.setTableBorders({ tableIndex, inside, outside })
      const changes: string[] = []
      if (inside) changes.push(`inside: style=${inside.style ?? "def"}, color=${inside.color ?? "def"}, size=${inside.size ?? "def"}`)
      if (outside) changes.push(`outside: style=${outside.style ?? "def"}, color=${outside.color ?? "def"}, size=${outside.size ?? "def"}`)
      return `Action: Table borders updated\nDetail: ${changes.join("; ")}\nNext: word_set_table_shading({color:"#E8F0FE", target:"row"})`
    }),
  )

  server.registerTool(
    "word_set_table_shading",
    {
      description: "Set background shading color for a table or first row.",
      inputSchema: {
        tableIndex: z.number().int().min(1).optional().describe("Table index (1-based, default: 1)"),
        color: z.string().regex(/^#[0-9a-fA-F]{6}$/, "Must be hex color like #FF0000").describe("Hex color like #E8F0FE"),
        target: z.enum(["table", "row"]).optional().describe("Apply to entire table or first row only"),
      },
    },
    mcpCall(security, "word_set_table_shading", async ({ tableIndex, color, target }) => {
      await content.setTableShading({ tableIndex, color, target })
      const t = target === "row" ? "First row" : "Entire table"
      return `Action: Table shading set (${color})\nDetail: Target: ${t}\nNext: word_set_table_borders({outside:{style:"single", color:"black", size:8}})`
    }),
  )

  server.registerTool(
    "word_merge_table_cells",
    {
      description: "Merge a range of table cells.",
      inputSchema: {
        tableIndex: z.number().int().min(1).optional().describe("Table index (1-based, default: 1)"),
        rowStart: z.number().int().min(1).describe("Start row (1-based)"),
        colStart: z.number().int().min(1).describe("Start column (1-based)"),
        rowEnd: z.number().int().min(1).describe("End row (1-based)"),
        colEnd: z.number().int().min(1).describe("End column (1-based)"),
      },
    },
    mcpCall(security, "word_merge_table_cells", async ({ tableIndex, rowStart, colStart, rowEnd, colEnd }) => {
      await content.mergeTableCells({ tableIndex, rowStart, colStart, rowEnd, colEnd })
      return `Action: Cells merged (${rowStart},${colStart})→(${rowEnd},${colEnd})\nNext: word_edit_cell({row:${rowStart}, column:${colStart}, text:"..."})`
    }),
  )

  server.registerTool(
    "word_set_column_width",
    {
      description: "Set the width of a specific column in points.",
      inputSchema: {
        tableIndex: z.number().int().min(1).optional().describe("Table index (1-based, default: 1)"),
        column: z.number().int().min(1).describe("Column number (1-based)"),
        width: z.number().min(1).max(5000).describe("Width in points (e.g. 100 ≈ 3.5cm)"),
      },
    },
    mcpCall(security, "word_set_column_width", async ({ tableIndex, column, width }) => {
      await content.setColumnWidth({ tableIndex, column, width })
      return `Action: Column ${column} width = ${width}pt\nNext: word_set_row_height({row:1, height:30}) or word_set_cell_font({row:1, column:${column}, bold:true})`
    }),
  )

  server.registerTool(
    "word_set_row_height",
    {
      description: "Set the height of a specific row in points.",
      inputSchema: {
        tableIndex: z.number().int().min(1).optional().describe("Table index (1-based, default: 1)"),
        row: z.number().int().min(1).describe("Row number (1-based)"),
        height: z.number().min(1).max(5000).describe("Height in points (e.g. 30 ≈ 1cm)"),
      },
    },
    mcpCall(security, "word_set_row_height", async ({ tableIndex, row, height }) => {
      await content.setRowHeight({ tableIndex, row, height })
      return `Action: Row ${row} height = ${height}pt\nNext: word_set_cell_font({row:${row}, column:1, bold:true, size:12})`
    }),
  )

  server.registerTool(
    "word_set_cell_font",
    {
      description: "Set font formatting for a specific table cell. WHEN: after filling a cell, to style its text. NOT: want to format all cells at once? consider word_apply_table_style.",
      inputSchema: {
        tableIndex: z.number().int().min(1).optional().describe("Table index (1-based, default: 1)"),
        row: z.number().int().min(1).describe("Row number (1-based)"),
        column: z.number().int().min(1).describe("Column number (1-based)"),
        name: z.string().max(100).optional().describe("Font name (e.g. 'Arial')"),
        size: z.number().min(1).max(1638).optional().describe("Font size in points"),
        bold: z.boolean().optional().describe("Bold"),
        italic: z.boolean().optional().describe("Italic"),
        underline: z.enum(["none", "single", "double", "wavy"]).optional().describe("Underline style"),
        color: z.enum(["auto", "black", "blue", "turquoise", "bright_green", "pink", "red", "yellow", "white", "dark_blue", "teal", "green", "violet", "dark_red", "dark_yellow", "gray_50", "gray_25"]).optional().describe("Font color"),
        strikethrough: z.boolean().optional().describe("Strikethrough"),
      },
    },
    mcpCall(security, "word_set_cell_font", async ({ tableIndex, row, column, name, size, bold, italic, underline, color, strikethrough }) => {
      await content.setCellFont({ tableIndex, row, column, name, size, bold, italic, underline, color, strikethrough })
      const props: string[] = []
      if (name) props.push(`font: ${name}`)
      if (size) props.push(`${size}pt`)
      if (bold !== undefined) props.push(bold ? "bold" : "no bold")
      if (italic !== undefined) props.push(italic ? "italic" : "no italic")
      if (underline) props.push(`underline: ${underline}`)
      if (color) props.push(`color: ${color}`)
      return `Action: Cell (${row},${column}) font updated\nDetail: ${props.join(", ")}\nNext: word_set_cell_vertical_alignment({row:${row}, column:${column}, alignment:"center"})`
    }),
  )

  server.registerTool(
    "word_apply_table_style",
    {
      description: "Apply a built-in Word table style by name.",
      inputSchema: {
        tableIndex: z.number().int().min(1).optional().describe("Table index (1-based, default: 1)"),
        styleName: z.string().min(1).max(255).describe("Style name (e.g. 'Table Grid', 'Light List Accent 1')"),
      },
    },
    mcpCall(security, "word_apply_table_style", async ({ tableIndex, styleName }) => {
      await content.applyTableStyle({ tableIndex, styleName })
      return `Action: Table style "${styleName}" applied\nNext: word_set_table_shading({color:"#E8F0FE"}) or word_set_table_borders({outside:{style:"single"}})`
    }),
  )

  server.registerTool(
    "word_add_table_column",
    {
      description: "Add a new column to a table.",
      inputSchema: {
        tableIndex: z.number().int().min(1).optional().describe("Table index (1-based, default: 1)"),
        column: z.number().int().min(1).optional().describe("Add to left of this column (omit to append at end)"),
      },
    },
    mcpCall(security, "word_add_table_column", async ({ tableIndex, column }) => {
      await content.addTableColumn({ tableIndex, column })
      const pos = column ? `left of column ${column}` : "end"
      return `Action: Column added (${pos})\nNext: word_set_column_width({column:1, width:100}) or word_edit_cell({row:1, column:1, text:"..."})`
    }),
  )

  server.registerTool(
    "word_delete_table_column",
    {
      description: "Delete a column from a table.",
      inputSchema: {
        tableIndex: z.number().int().min(1).optional().describe("Table index (1-based, default: 1)"),
        column: z.number().int().min(1).describe("Column number to delete (1-based)"),
      },
    },
    mcpCall(security, "word_delete_table_column", async ({ tableIndex, column }) => {
      await content.deleteTableColumn({ tableIndex, column })
      return `Action: Column ${column} deleted\nNext: word_undo({count:1}) or word_add_table_column({column:${column}})`
    }),
  )

  server.registerTool(
    "word_set_cell_vertical_alignment",
    {
      description: "Set vertical alignment for a specific cell. WHEN: text inside a cell needs to be top/bottom aligned. NOT: want horizontal alignment? use word_set_paragraph.",
      inputSchema: {
        tableIndex: z.number().int().min(1).optional().describe("Table index (1-based, default: 1)"),
        row: z.number().int().min(1).describe("Row number (1-based)"),
        column: z.number().int().min(1).describe("Column number (1-based)"),
        alignment: z.enum(["top", "center", "bottom"]).describe("Vertical alignment"),
      },
    },
    mcpCall(security, "word_set_cell_vertical_alignment", async ({ tableIndex, row, column, alignment }) => {
      await content.setCellVerticalAlignment({ tableIndex, row, column, alignment })
      return `Action: Cell (${row},${column}) vertical alignment: ${alignment}\nNext: word_set_cell_font({row:${row}, column:${column}, bold:true})`
    }),
  )

  server.registerTool(
    "word_table_to_text",
    {
      description: "Convert a table to plain text (removes table structure).",
      inputSchema: {
        tableIndex: z.number().int().min(1).optional().describe("Table index (1-based, default: 1)"),
        separator: z.string().max(10).optional().describe("Separator character (default: tab)"),
      },
    },
    mcpCall(security, "word_table_to_text", async ({ tableIndex, separator }) => {
      const result = await content.tableToText({ tableIndex, separator })
      return `Action: Table converted to text\nDetail: ${result}\nNext: word_text_to_table() to convert back or word_undo()`
    }),
  )

  server.registerTool(
    "word_text_to_table",
    {
      description: "Convert selected text into a table. WHEN: text is structured with tab/separator delimiters. NOT: no selection? use word_select_text first.",
      inputSchema: {
        separator: z.string().max(10).optional().describe("Column separator in text (default: tab)"),
        autoFitBehavior: z.enum(["fixed", "contents", "window"]).optional().describe("Auto-fit behavior"),
      },
    },
    mcpCall(security, "word_text_to_table", async ({ separator, autoFitBehavior }) => {
      const result = await content.textToTable({ separator, autoFitBehavior })
      return `Action: Text → table (${result.rows}x${result.columns})\nNext: word_set_column_width({column:1, width:120}) or word_apply_table_style({styleName:"Table Grid"})`
    }),
  )
}
