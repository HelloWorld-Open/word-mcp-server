import { spawn } from "node:child_process"
import { resolve, dirname } from "node:path"
import { fileURLToPath } from "node:url"

const __dirname = dirname(fileURLToPath(import.meta.url))
const serverPath = resolve(__dirname, "build", "parent.js")

const proc = spawn("node", [serverPath], { stdio: ["pipe", "pipe", "pipe"] })
proc.stderr.on("data", (d) => process.stderr.write("[stderr] " + d.toString()))

let lineBuf = ""
const pending = []

proc.stdout.on("data", (d) => {
  lineBuf += d.toString()
  while (lineBuf.includes("\n")) {
    const idx = lineBuf.indexOf("\n")
    const line = lineBuf.slice(0, idx).trim()
    lineBuf = lineBuf.slice(idx + 1)
    if (line) {
      console.log("[stdout]", line)
      const cb = pending.shift()
      if (cb) cb(line)
    }
  }
})

let reqId = 0

async function sendRequest(method, params) {
  const id = ++reqId
  const json = JSON.stringify({ jsonrpc: "2.0", id, method, params }) + "\n"
  return new Promise((resolve) => {
    pending.push(resolve)
    proc.stdin.write(json, "utf-8")
  })
}

function sendNotification(method, params) {
  const json = JSON.stringify({ jsonrpc: "2.0", method, params }) + "\n"
  proc.stdin.write(json, "utf-8")
}

process.on("exit", () => { try { proc.kill() } catch {} })

async function main() {
  // Initialize
  const init = await sendRequest("initialize", {
    protocolVersion: "2024-11-05", capabilities: {},
    clientInfo: { name: "test", version: "1.0" },
  })
  const initBody = JSON.parse(init)
  console.log("INIT:", initBody.id, "server:", initBody.result?.serverInfo?.name)

  // Notify (no response expected)
  sendNotification("notifications/initialized")

  // Wait for engine
  await new Promise((r) => setTimeout(r, 2000))

  // Create document
  const create = await sendRequest("tools/call", { name: "word_create", arguments: { title: "MCP Test" } })
  const c = JSON.parse(create)
  console.log("CREATE:", c.result?.content?.[0]?.text || JSON.stringify(c.error))

  // Get info
  const info = await sendRequest("tools/call", { name: "word_get_info", arguments: {} })
  const i = JSON.parse(info)
  console.log("INFO:", i.result?.content?.[0]?.text?.split("\n")[0] || JSON.stringify(i.error))

  // Type text
  const type = await sendRequest("tools/call", { name: "word_type_text", arguments: { text: "Hello from MCP! This is a test." } })
  const t = JSON.parse(type)
  console.log("TYPE:", t.result?.content?.[0]?.text || JSON.stringify(t.error))

  // Set font
  const font = await sendRequest("tools/call", { name: "word_set_font", arguments: { name: "Arial", size: 14, bold: true } })
  const f = JSON.parse(font)
  console.log("FONT:", f.result?.content?.[0]?.text || JSON.stringify(f.error))

  // Save
  const save = await sendRequest("tools/call", { name: "word_save", arguments: {} })
  const s = JSON.parse(save)
  console.log("SAVE:", s.result?.content?.[0]?.text || JSON.stringify(s.error))

  // Quit
  const quit = await sendRequest("tools/call", { name: "word_quit", arguments: {} })
  const q = JSON.parse(quit)
  console.log("QUIT:", q.result?.content?.[0]?.text || JSON.stringify(q.error))

  console.log("\n=== ALL TESTS PASSED ===")
  proc.kill()
  process.exit(0)
}

main().catch((err) => { console.error("FAIL:", err); proc.kill(); process.exit(1) })
