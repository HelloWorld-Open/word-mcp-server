import { spawn } from "node:child_process"
import { resolve, dirname } from "node:path"
import { fileURLToPath } from "node:url"

const __filename = fileURLToPath(import.meta.url)
const __dirname = dirname(__filename)

const CHILD_SCRIPT = resolve(__dirname, "child.js")
const WATCHDOG_TIMEOUT_MS = parseInt(process.env.WATCHDOG_TIMEOUT_MS ?? "30000", 10)
const WATCHDOG_INTERVAL_MS = parseInt(process.env.WATCHDOG_INTERVAL_MS ?? "5000", 10)

function log(...args: unknown[]) {
  console.error("[parent]", ...args)
}

class Watchdog {
  private child: ReturnType<typeof spawn> | null = null
  private lastChildOutput = Date.now()
  private restarting = false
  private shutdown = false
  private stdinHandler: ((chunk: Buffer) => void) | null = null
  private generation = 0

  start(): void {
    this.spawnChild()
    this.startWatchdog()
    this.setupSignalHandlers()
  }

  private spawnChild(): void {
    if (this.restarting) return
    this.generation++
    const myGeneration = this.generation
    log(`Spawning child (generation ${myGeneration}): ${CHILD_SCRIPT}`)

    this.lastChildOutput = Date.now()
    this.child = spawn("node", [CHILD_SCRIPT], {
      stdio: ["pipe", "pipe", "pipe"],
      windowsHide: false,
    })

    const child = this.child
    const childStdin = child.stdin!
    const childStdout = child.stdout!
    const childStderr = child.stderr!

    // 清理旧的 stdin listener (restart 时)
    const ps = process.stdin!
    if (this.stdinHandler) ps.removeListener("data", this.stdinHandler)

    this.stdinHandler = (chunk: Buffer) => {
      try { childStdin.write(chunk) } catch { /* child may be dead */ }
    }
    ps.on("data", this.stdinHandler)

    // 子进程 stdout → 父进程 stdout (MCP 响应)
    childStdout.on("data", (chunk: Buffer) => {
      this.lastChildOutput = Date.now()
      process.stdout.write(chunk)
    })

    // 子进程 stderr → 父进程 stderr (日志 + 心跳检测)
    childStderr.on("data", (chunk: Buffer) => {
      this.lastChildOutput = Date.now()
      process.stderr.write(chunk)
    })

    // 子进程退出处理
    child.on("exit", (code, signal) => {
      log(`Child exited (code: ${code}, signal: ${signal})`)
      if (this.generation !== myGeneration) return
      if (!this.shutdown) {
        log("Restarting child after unexpected exit...")
        this.spawnChild()
      }
    })

    child.on("error", (err) => {
      log(`Child error: ${err.message}`)
    })

    log("Child process spawned")
  }

  private startWatchdog(): void {
    setInterval(() => {
      if (this.shutdown || !this.child) return
      if (this.child.killed) return

      const idle = Date.now() - this.lastChildOutput
      if (idle > WATCHDOG_TIMEOUT_MS) {
        log(`Watchdog: no child output for ${idle}ms, killing and restarting...`)
        this.restarting = true
        this.child.kill("SIGTERM")
        // 等待进程退出
        setTimeout(() => {
          this.restarting = false
          this.spawnChild()
        }, 2000)
      }
    }, WATCHDOG_INTERVAL_MS)
  }

  private setupSignalHandlers(): void {
    const cleanup = () => {
      this.shutdown = true
      log("Shutting down parent...")
      if (this.child && !this.child.killed) {
        this.child.kill("SIGTERM")
      }
      process.exit(0)
    }

    process.on("SIGINT", cleanup)
    process.on("SIGTERM", cleanup)
    process.on("unhandledRejection", (reason) => {
      log("Unhandled rejection:", reason)
    })
  }
}

const watchdog = new Watchdog()
watchdog.start()
