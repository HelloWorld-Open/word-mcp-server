import { describe, it, expect, vi, beforeEach, afterEach } from "vitest"
import { SessionPathMachine } from "../../../src/server/session-path-machine.js"

describe("SessionPathMachine", () => {
  let machine: SessionPathMachine

  beforeEach(() => {
    vi.useFakeTimers()
    machine = new SessionPathMachine()
  })

  afterEach(() => {
    vi.useRealTimers()
  })

  describe("initial state", () => {
    it("starts in idle state", () => {
      expect(machine.currentPath).toBe("idle")
      expect(machine.isStreamingActive).toBe(false)
    })
  })

  describe("acquireStreamLock", () => {
    it("transitions to streaming from idle", () => {
      const result = machine.acquireStreamLock("word_stream_start")
      expect(result).toBeNull()
      expect(machine.currentPath).toBe("streaming")
      expect(machine.isStreamingActive).toBe(true)
    })

    it("rejects when in editing mode", () => {
      machine.enterEditMode()
      const result = machine.acquireStreamLock("word_stream_start")
      expect(result).toContain("[EDIT_MODE]")
      expect(machine.currentPath).toBe("editing")
    })

    it("rejects when already streaming", () => {
      machine.acquireStreamLock("word_stream_start")
      const result = machine.acquireStreamLock("word_stream_block")
      expect(result).toContain("[STREAMING]")
      expect(machine.currentPath).toBe("streaming")
    })

    it("starts streaming watchdog on acquire", () => {
      machine.acquireStreamLock("word_stream_start")
      vi.advanceTimersByTime(599_999)
      expect(machine.currentPath).toBe("streaming")
      vi.advanceTimersByTime(1)
      expect(machine.currentPath).toBe("idle")
    })
  })

  describe("releaseStreamLock", () => {
    it("transitions to idle from streaming", () => {
      machine.acquireStreamLock("word_stream_start")
      machine.releaseStreamLock()
      expect(machine.currentPath).toBe("idle")
      expect(machine.isStreamingActive).toBe(false)
    })

    it("is a no-op when in idle", () => {
      machine.releaseStreamLock()
      expect(machine.currentPath).toBe("idle")
    })

    it("stops streaming watchdog on release", () => {
      machine.acquireStreamLock("word_stream_start")
      machine.releaseStreamLock()
      vi.advanceTimersByTime(600_000)
      expect(machine.currentPath).toBe("idle")
    })
  })

  describe("enterEditMode / exitEditMode", () => {
    it("transitions to editing from idle", () => {
      machine.enterEditMode()
      expect(machine.currentPath).toBe("editing")
      expect(machine.isStreamingActive).toBe(false)
    })

    it("is a no-op when in streaming", () => {
      machine.acquireStreamLock("word_stream_start")
      machine.enterEditMode()
      expect(machine.currentPath).toBe("streaming")
    })

    it("exitEditMode transitions to idle from editing", () => {
      machine.enterEditMode()
      machine.exitEditMode()
      expect(machine.currentPath).toBe("idle")
    })

    it("exitEditMode is a no-op when not editing", () => {
      machine.exitEditMode()
      expect(machine.currentPath).toBe("idle")
    })

    it("exitEditMode is a no-op when in streaming", () => {
      machine.acquireStreamLock("word_stream_start")
      machine.exitEditMode()
      expect(machine.currentPath).toBe("streaming")
    })
  })

  describe("refreshWatchdog", () => {
    it("resets streaming watchdog timer", () => {
      machine.acquireStreamLock("word_stream_start")
      vi.advanceTimersByTime(300_000)
      machine.refreshWatchdog()
      vi.advanceTimersByTime(300_000)
      expect(machine.currentPath).toBe("streaming")
      vi.advanceTimersByTime(300_000)
      expect(machine.currentPath).toBe("idle")
    })

    it("is a no-op when not streaming", () => {
      machine.refreshWatchdog()
      expect(machine.currentPath).toBe("idle")
    })
  })

  describe("setOnLog", () => {
    it("calls log handler on state transitions", () => {
      const logFn = vi.fn()
      machine.setOnLog(logFn)
      machine.acquireStreamLock("word_stream_start")
      expect(logFn).toHaveBeenCalledWith("info", expect.stringContaining("Stream lock acquired"))
    })

    it("default setOnLog(null) does not crash", () => {
      machine.setOnLog(null as any)
      machine.acquireStreamLock("word_stream_start")
      expect(machine.currentPath).toBe("streaming")
    })
  })
})
