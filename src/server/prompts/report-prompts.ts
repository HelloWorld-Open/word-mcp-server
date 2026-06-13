import { z } from "zod"
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js"

export function registerReportPrompts(server: McpServer): void {
  server.registerPrompt(
    "create_report",
    {
      description: "Generate a step-by-step plan for creating a structured Word report",
      argsSchema: {
        title: z.string().describe("Report title"),
        sections: z.string().describe("Comma-separated section headings"),
        style: z.string().describe("Report style (professional, academic, casual)").default("professional"),
      },
    },
    ({ title, sections, style }) => ({
      messages: [
        {
          role: "user",
          content: {
            type: "text",
            text: `Create a Word document report with the following specifications:

Title: ${title}
Sections: ${sections}
Style: ${style}

Use these tools in order:
1. word_stream_start — create the document with title "${title}"
2. word_set_properties — set author and metadata
3. word_stream_block — write all section content as markdown
4. word_mgr_set_header — add header with report title
5. word_mgr_set_page_numbers — add page numbers
6. word_stream_end — save the document
7. word_get_info — verify the document`,
          },
        },
      ],
    })
  )

  server.registerPrompt(
    "format_document",
    {
      description: "Get a guided workflow for formatting an existing document",
      argsSchema: {
        style: z.string().describe("Target formatting style").default("professional"),
      },
    },
    ({ style }) => ({
      messages: [
        {
          role: "user",
          content: {
            type: "text",
            text: `Reformat the current document with ${style} styling. Use these tools:

1. word_select_all — select entire document
2. word_set_font — set base font
3. word_set_paragraph — set paragraph spacing and alignment
4. word_set_page_setup — ensure consistent margins
5. word_list_styles — check available styles
6. word_set_properties — update metadata`,
          },
        },
      ],
    })
  )
}
