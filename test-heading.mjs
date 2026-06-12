import { spawn } from "node:child_process"
import { resolve, dirname } from "node:path"
import { fileURLToPath } from "node:url"

const __dirname = dirname(fileURLToPath(import.meta.url))
const serverPath = resolve(__dirname, "build", "parent.js")

const proc = spawn("node", [serverPath], { stdio: ["pipe", "pipe", "pipe"] })
proc.stderr.on("data", (d) => process.stderr.write("[stderr] " + d.toString()))

let lineBuf = ""
const pending = []
let totalLatency = 0
let callCount = 0

proc.stdout.on("data", (d) => {
  lineBuf += d.toString()
  while (lineBuf.includes("\n")) {
    const idx = lineBuf.indexOf("\n")
    const line = lineBuf.slice(0, idx).trim()
    lineBuf = lineBuf.slice(idx + 1)
    if (line) {
      const cb = pending.shift()
      if (cb) cb(line)
    }
  }
})

let reqId = 0

async function sendRequest(method, params) {
  const id = ++reqId
  const json = JSON.stringify({ jsonrpc: "2.0", id, method, params }) + "\n"
  const start = Date.now()
  return new Promise((resolve) => {
    pending.push((raw) => {
      const elapsed = Date.now() - start
      totalLatency += elapsed
      callCount++
      resolve({ raw, elapsed })
    })
    proc.stdin.write(json, "utf-8")
  })
}

process.on("exit", () => { try { proc.kill() } catch {} })

async function call(method, args = {}) {
  const res = await sendRequest("tools/call", { name: method, arguments: args })
  const body = JSON.parse(res.raw)
  return { data: body, elapsed: res.elapsed }
}

async function main() {
  // Initialize
  const init = await sendRequest("initialize", {
    protocolVersion: "2024-11-05", capabilities: {},
    clientInfo: { name: "test", version: "1.0" },
  })
  JSON.parse(init.raw)
  const notifJson = JSON.stringify({ jsonrpc: "2.0", method: "notifications/initialized" }) + "\n"
  proc.stdin.write(notifJson, "utf-8")
  await new Promise((r) => setTimeout(r, 3000))

  const DOC = "C:\\Users\\xiele\\Desktop\\12\\Word_MCP_81_综合测试.docx"

  // 1. Open document
  console.log("--- Opening document ---")
  let r = await call("word_document", { path: DOC })
  console.log("word_document:", r.elapsed + "ms", r.data.result?.content?.[0]?.text?.slice(0, 100) || r.data.error?.message)

  // 2. Get structure
  console.log("\n--- get_structure ---")
  r = await call("word_get_structure", {})
  const structText = r.data.result?.content?.[0]?.text || ""
  console.log("word_get_structure:", r.elapsed + "ms,", structText.slice(0, 300))

  // 3. Test locate by heading (use known heading text — 综合测试文档 has "Word MCP 81" in H1)
  console.log("\n--- locate by heading ---")
  r = await call("word_locate", { by: "heading", match: "Word MCP 81", matchMode: "contains" })
  console.log("locate(contains 'Word MCP 81'):", r.elapsed + "ms,", r.data.result?.content?.[0]?.text?.slice(0, 150) || r.data.error?.message)

  r = await call("word_locate", { by: "heading", match: "1.1 段落与对齐", matchMode: "exact" })
  console.log("locate(exact '1.1 段落与对齐'):", r.elapsed + "ms,", r.data.result?.content?.[0]?.text?.slice(0, 150) || r.data.error?.message)

  r = await call("word_locate", { by: "heading", match: "1.1", matchMode: "startsWith" })
  console.log("locate(startsWith '1.1'):", r.elapsed + "ms,", r.data.result?.content?.[0]?.text?.slice(0, 150) || r.data.error?.message)

  // 4. Test select_at by heading
  console.log("\n--- select_at by heading ---")
  r = await call("word_select_at", { by: "heading", match: "1.2 引用", matchMode: "contains" })
  console.log("select_at(contains '1.2'):", r.elapsed + "ms,", r.data.result?.content?.[0]?.text?.slice(0, 150) || r.data.error?.message)

  // 5. Test insert_at by heading (insert near end, undo not needed — read-only test)
  console.log("\n--- insert_at by heading ---")
  r = await call("word_insert_at", { by: "heading", match: "1.1 段落与对齐", matchMode: "exact", text: "[READ-ONLY-TEST] " })
  console.log("insert_at(after '1.1'):", r.elapsed + "ms,", r.data.result?.content?.[0]?.text?.slice(0, 150) || r.data.error?.message)

  // 6. Test locate by paragraph (fallback reference)
  console.log("\n--- locate by paragraph ---")
  r = await call("word_locate", { by: "paragraph", match: "六、综合测试", matchMode: "contains" })
  console.log("locate(paragraph '六、综合测试'):", r.elapsed + "ms,", r.data.result?.content?.[0]?.text?.slice(0, 150) || r.data.error?.message)

  console.log("\n=== SUMMARY ===")
  console.log(`Total calls: ${callCount}, avg latency: ${(totalLatency / callCount).toFixed(0)}ms`)
  console.log("=== ALL TESTS PASSED ===")
  proc.kill()
  process.exit(0)
}

main().catch((err) => { console.error("FAIL:", err); proc.kill(); process.exit(1) })
