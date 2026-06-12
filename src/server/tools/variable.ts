import { z } from "zod"
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js"
import { VariableReplacer } from "../../word/variable-replacer.js"
import { SecurityManager } from "../../security/policy.js"
import { mcpCall } from "./helper.js"

export function registerVariableTool(
  server: McpServer,
  replacer: VariableReplacer,
  security: SecurityManager,
): void {
  server.registerTool(
    "word_replace_variables",
    {
      description: "Replace {{placeholder}} variables in the document with provided values. WHEN: generating documents from templates with {{变量名}} markers. NOT: for simple find-and-replace, use word_find_replace.",
      inputSchema: {
        variables: z.record(z.string().min(1).max(100), z.string().max(100000)).describe("Key-value pairs: keys are variable names (without {{}}), values are replacement text"),
      },
    },
    mcpCall(security, "word_replace_variables", async ({ variables }) => {
      const results = await replacer.replaceVariables(variables)
      const lines = results.map((r) => `- {{${r.key}}}: ${r.count} replacement(s)`)
      const total = results.reduce((s, r) => s + r.count, 0)
      return [
        `Action: ${results.length} variable(s) processed, ${total} total replacements`,
        ...lines,
        'Next: word_save() or word_type_text({text:"...", mode:"instant"})',
      ].join("\n")
    }),
  )
}
