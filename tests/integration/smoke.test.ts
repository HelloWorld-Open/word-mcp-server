import { describe, it, expect } from "vitest"
import { spawn } from "node:child_process"
import { resolve } from "node:path"

const buildIndex = resolve(import.meta.dirname, "..", "..", "build", "index.js")

function mcpRequest(request: string): Promise<Record<string, unknown>[]> {
  return new Promise((resolve, reject) => {
    const child = spawn("node", [buildIndex], {
      stdio: ["pipe", "pipe", "pipe"],
    })

    let stdout = ""
    let stderr = ""

    const timeout = setTimeout(() => {
      child.kill()
      reject(new Error("timeout"))
    }, 15000)

    child.stdout.on("data", (data: Buffer) => {
      stdout += data.toString()
    })

    child.stderr.on("data", (data: Buffer) => {
      stderr += data.toString()
    })

    child.on("error", (err: Error) => {
      clearTimeout(timeout)
      reject(err)
    })

    child.on("close", () => {
      clearTimeout(timeout)
      if (stderr.trim()) {
        console.error("[mcp-stderr]", stderr.trim())
      }
      const lines = stdout.trim().split("\n").filter(Boolean)
      resolve(lines.map((l) => JSON.parse(l)))
    })

    child.stdin.write(request + "\n")
    child.stdin.end()
  })
}

describe.skipIf(process.platform !== "win32")("MCP Server smoke tests", () => {
  it("should respond to initialize", async () => {
    const req = {
      jsonrpc: "2.0",
      id: 1,
      method: "initialize",
      params: {
        protocolVersion: "2024-11-05",
        capabilities: {},
        clientInfo: { name: "test", version: "1.0" },
      },
    }
    const responses = await mcpRequest(JSON.stringify(req))
    const init = responses.find((r: any) => r.id === 1)
    expect(init).toBeDefined()
    expect((init as any).result?.serverInfo?.name).toBe("word-mcp-server")
  })

  it("should list all tools", async () => {
    const req = { jsonrpc: "2.0", id: 2, method: "tools/list", params: {} }
    const responses = await mcpRequest(JSON.stringify(req))
    const list = responses.find((r: any) => r.id === 2)
    expect(list).toBeDefined()
    const tools = (list as any).result?.tools
    expect(tools).toBeDefined()
    expect(tools.length).toBe(113)
  })

  it("word tool names follow word_ prefix convention", async () => {
    const req = { jsonrpc: "2.0", id: 3, method: "tools/list", params: {} }
    const responses = await mcpRequest(JSON.stringify(req))
    const list = responses.find((r: any) => r.id === 3)
    const tools = (list as any).result?.tools as Array<{ name: string }>
    for (const t of tools) {
      expect(t.name).toMatch(/^word_/)
    }
  })
})
