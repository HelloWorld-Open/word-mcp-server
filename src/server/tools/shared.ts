import { z } from "zod"
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js"
import { SecurityManager } from "../../security/policy.js"
import { ComError } from "../../word/com-errors.js"
import { WordMcpError, WordEngineTimeoutError } from "../../security/errors.js"
import type { ServerContext } from "../server-context.js"
import { mcpCall } from "./helper.js"
import type { Precondition } from "../session-director.js"
import type { ToolResponse } from "../tool-result.js"

export const ColorSchema = z.enum([
  "auto", "black", "blue", "turquoise", "bright_green", "pink", "red",
  "yellow", "white", "dark_blue", "teal", "green", "violet", "dark_red",
  "dark_yellow", "gray_50", "gray_25",
])

export const locatorFields = {
  by: z.enum(["heading", "paragraph", "table", "bookmark", "cursor"]).optional().describe("Target type (default: heading)"),
  match: z.string().max(5000).optional().describe("Text to match (for heading/paragraph)"),
  matchMode: z.enum(["exact", "contains", "startsWith", "regex"]).optional().describe("Matching mode (default: exact)"),
  occurrence: z.number().int().min(1).max(1000).optional().describe("Which occurrence to target (1-based, default: 1)"),
  offsetDirection: z.enum(["before", "after"]).optional().describe("Offset direction from the matched element"),
  offsetCount: z.number().int().min(1).max(1000).optional().describe("Number of paragraphs to offset (default: 1)"),
  name: z.string().min(1).max(255).optional().describe("Bookmark name (required when by='bookmark')"),
  level: z.number().int().min(1).max(9).optional().describe("Heading level filter (1-9). Applied BEFORE match and occurrence, narrowing to same-level headings only."),
} as const

interface RegToolOptions {
  preconditions?: Precondition[]
  timeoutMs?: number
}

export function createRegTool(server: McpServer, security: SecurityManager, context: ServerContext) {
  return function regTool(
    name: string,
    config: { description?: string; inputSchema?: any },
    handler: (args: any) => Promise<string | ToolResponse>,
    options?: RegToolOptions,
  ): void {
    server.registerTool(name, config, mcpCall(security, context, name, handler, options))
  }
}

export const READ_ONLY_TOOLS = new Set([
  "word_get_text", "word_get_paragraph", "word_get_structure", "word_get_info",
  "word_get_status", "word_get_table_data", "word_get_comments", "word_get_bookmarks",
  "word_get_lists", "word_get_sections", "word_get_cursor_info", "word_locate",
  "word_list_styles", "word_where_am_i",
])

export function isEngineError(err: unknown): boolean {
  if (err instanceof WordEngineTimeoutError) return true
  if (err instanceof ComError) return true
  if (err instanceof WordMcpError) return false
  return false
}

export function isReadOnlyTool(toolName: string): boolean {
  return READ_ONLY_TOOLS.has(toolName)
}
