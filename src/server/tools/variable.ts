import { z } from "zod"
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js"
import { WordContentWriter } from "../../word/word-content-writer.js"
import { SecurityManager } from "../../security/policy.js"
import type { ServerContext } from "../server-context.js"
import { createRegTool } from "./shared.js"

export function registerVariableTool(
  server: McpServer,
  context: ServerContext,
  contentWriter: WordContentWriter,
  security: SecurityManager,
): void {
  const regTool = createRegTool(server, security, context)
  regTool("word_replace_variables",
    {
      description: "WHEN: generating documents from templates with {{placeholder}} markers such as {{clientName}}, {{date}}, {{amount}}. WHAT: finds all {{key}} patterns in the document and replaces each with the corresponding value from the provided object. CONSTRAINT: keys must match exactly (case-sensitive). For plain text find-and-replace without templates, use word_find_replace.",
      inputSchema: {
        variables: z.record(z.string().min(1).max(100), z.string().max(100000)).describe("Key-value pairs: keys are variable names (without {{}}), values are replacement text"),
      },
    },
    async ({ variables }) => {
      const results = await contentWriter.replaceVariables(variables)
      const lines = results.map((r) => `- {{${r.key}}}: ${r.count} replacement(s)`)
      const total = results.reduce((s, r) => s + r.count, 0)
      return [
        `Action: ${results.length} variable(s) processed, ${total} total replacements`,
        ...lines,
        'Next: word_save() or word_type_text({text:"...", mode:"instant"})',
      ].join("\n")
    },
  )
}
