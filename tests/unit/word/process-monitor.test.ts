import { describe, it, expect, vi, beforeEach, afterEach } from "vitest"
import { ProcessMonitor } from "../../../src/word/process-monitor.js"

describe("ProcessMonitor", () => {
  let onStatusChange: ReturnType<typeof vi.fn>
  let checkFn: ReturnType<typeof vi.fn>

  beforeEach(() => {
    vi.useFakeTimers()
    onStatusChange = vi.fn()
    checkFn = vi.fn((cb: (err: Error | null, found: boolean) => void) => cb(null, true))
  })

  afterEach(() => {
    vi.useRealTimers()
  })

  it("starts monitoring and calls checkFn at interval", () => {
    const pm = new ProcessMonitor(100, { checkFn })
    pm.onAliveChange(onStatusChange)
    pm.start()
    expect(checkFn).toHaveBeenCalledTimes(1)
    vi.advanceTimersByTime(100)
    expect(checkFn).toHaveBeenCalledTimes(2)
    pm.stop()
  })

  it("notifies alive=true when process is found", () => {
    checkFn = vi.fn((cb: (err: Error | null, found: boolean) => void) => cb(null, true))
    const pm = new ProcessMonitor(100, { checkFn })
    pm.onAliveChange(onStatusChange)
    pm.start()
    expect(onStatusChange).toHaveBeenCalledWith(true)
    pm.stop()
  })

  it("notifies alive=false when process is not found", () => {
    checkFn = vi.fn((cb: (err: Error | null, found: boolean) => void) => cb(null, false))
    const pm = new ProcessMonitor(100, { checkFn })
    pm.onAliveChange(onStatusChange)
    pm.start()
    expect(onStatusChange).toHaveBeenCalledWith(false)
    pm.stop()
  })

  it("ignores errors from checkFn", () => {
    checkFn = vi.fn((cb: (err: Error | null, found: boolean) => void) => cb(new Error("fail"), false))
    const pm = new ProcessMonitor(100, { checkFn })
    pm.onAliveChange(onStatusChange)
    pm.start()
    expect(onStatusChange).not.toHaveBeenCalled()
    pm.stop()
  })

  it("stop clears interval timer", () => {
    const pm = new ProcessMonitor(100, { checkFn })
    pm.start()
    pm.stop()
    checkFn.mockClear()
    vi.advanceTimersByTime(200)
    expect(checkFn).not.toHaveBeenCalled()
  })

  it("handles double start safely", () => {
    const pm = new ProcessMonitor(100, { checkFn })
    pm.start()
    pm.start()
    pm.stop()
  })

  it("handles double stop safely", () => {
    const pm = new ProcessMonitor(100, { checkFn })
    pm.start()
    pm.stop()
    pm.stop()
  })

  it("restarts monitoring after stop", () => {
    const pm = new ProcessMonitor(100, { checkFn })
    pm.start()
    pm.stop()
    checkFn.mockClear()
    pm.start()
    expect(checkFn).toHaveBeenCalledTimes(1)
    pm.stop()
  })
})
