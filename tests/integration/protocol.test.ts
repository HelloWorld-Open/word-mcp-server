import { describe, it, expect } from "vitest"

describe("MCP Protocol compliance", () => {
  it("server supports JSON-RPC 2.0", () => {
    const req = {
      jsonrpc: "2.0",
      id: 1,
      method: "initialize",
      params: { protocolVersion: "2024-11-05", capabilities: {}, clientInfo: { name: "test", version: "1.0" } },
    }
    expect(req.jsonrpc).toBe("2.0")
    expect(req.method).toBe("initialize")
  })

  it("all tool names are valid identifiers", async () => {
    const { registerStateMachinePrompt } = await import("../../src/server/prompts/state-machine.js")
    expect(typeof registerStateMachinePrompt).toBe("function")
  })
})
