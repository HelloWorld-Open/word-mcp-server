import { describe, it, expect, vi, beforeEach, afterEach } from "vitest"
import { SessionDirector } from "../../../src/server/session-director.js"

function createMockSession(overrides: Record<string, any> = {}) {
  return {
    isAlive: vi.fn().mockReturnValue(true),
    isUnhealthy: vi.fn().mockReturnValue(false),
    isBusy: vi.fn().mockReturnValue(false),
    setBusy: vi.fn(),
    recover: vi.fn().mockResolvedValue(undefined),
    markHealthy: vi.fn(),
    getLastCallStart: vi.fn().mockReturnValue(0),
    activeDoc: { Name: "doc.docx" } as any,
    activeDocPath: "C:\\doc.docx" as string | null,
    application: null as any,
    setActiveDoc: vi.fn(),
    setActiveDocPath: vi.fn(),
    lockPrintView: vi.fn(),
    tryAdoptActiveDoc: vi.fn().mockReturnValue(false),
    ...overrides,
  }
}

function createMockPositionMap() {
  return { markDirty: vi.fn(), scheduleRefresh: vi.fn() }
}

function createMockAppManager() {
  return { adoptDocument: vi.fn() }
}

describe("SessionDirector", () => {
  let mockSession: ReturnType<typeof createMockSession>

  beforeEach(() => {
    mockSession = createMockSession()
  })

  describe("constructor", () => {
    it("starts in idle path", () => {
      const d = new SessionDirector(mockSession, createMockPositionMap(), createMockAppManager())
      expect(d.currentPath).toBe("idle")
      expect(d.isStreamingActive).toBe(false)
    })

    it("stores session reference", () => {
      const d = new SessionDirector(mockSession, createMockPositionMap(), createMockAppManager())
      expect(d.session).toBe(mockSession)
    })

    it("stores positionMap reference", () => {
      const posMap = createMockPositionMap()
      const d = new SessionDirector(mockSession, posMap)
      expect(d.positionMap).toBe(posMap)
    })

    it("creates a fresh CircuitBreaker", () => {
      const d = new SessionDirector(mockSession)
      expect(d.circuitBreaker.getState()).toBe("CLOSED")
    })

    it("accepts null session", () => {
      const d = new SessionDirector(null, null)
      expect(d.session).toBeNull()
      expect(d.positionMap).toBeNull()
    })
  })

  describe("setOnLog", () => {
    it("sets log handler without error", () => {
      const d = new SessionDirector(mockSession)
      d.setOnLog(vi.fn())
    })
  })

  describe("path delegation methods", () => {
    let d: SessionDirector

    beforeEach(() => {
      d = new SessionDirector(mockSession)
    })

    it("delegates acquireStreamLock", () => {
      const result = d.acquireStreamLock("word_stream_start")
      expect(result).toBeNull()
      expect(d.isStreamingActive).toBe(true)
      expect(d.currentPath).toBe("streaming")
    })

    it("delegates releaseStreamLock", () => {
      d.acquireStreamLock("word_stream_start")
      d.releaseStreamLock()
      expect(d.isStreamingActive).toBe(false)
      expect(d.currentPath).toBe("idle")
    })

    it("delegates enterEditMode", () => {
      d.enterEditMode()
      expect(d.currentPath).toBe("editing")
    })

    it("delegates exitEditMode", () => {
      d.enterEditMode()
      d.exitEditMode()
      expect(d.currentPath).toBe("idle")
    })

    it("delegates refreshWatchdog", () => {
      d.acquireStreamLock("word_stream_start")
      d.refreshWatchdog()
      expect(d.isStreamingActive).toBe(true)
    })
  })

  describe("precheck", () => {
    it("returns ok when all conditions pass", async () => {
      const d = new SessionDirector(mockSession, createMockPositionMap())
      const result = await d.precheck("word_type_text")
      expect(result).toEqual({ ok: true })
    })

    it("blocks streaming tools during streaming", async () => {
      const d = new SessionDirector(mockSession)
      d.acquireStreamLock("word_stream_start")
      const result = await d.precheck("word_document")
      expect(result).toEqual({ ok: false, error: expect.stringContaining("[STREAMING]") })
    })

    it("allows non-blocked tools during streaming", async () => {
      const d = new SessionDirector(mockSession)
      d.acquireStreamLock("word_stream_start")
      const result = await d.precheck("word_stream_block")
      expect(result).toEqual({ ok: true })
    })

    it("blocks word_stream_start during editing", async () => {
      const d = new SessionDirector(mockSession)
      d.enterEditMode()
      const result = await d.precheck("word_stream_start")
      expect(result).toEqual({ ok: false, error: expect.stringContaining("[EDIT_MODE]") })
    })

    it("clears stale busy flag", async () => {
      const busySession = createMockSession({
        isBusy: vi.fn().mockReturnValue(true),
      })
      const d = new SessionDirector(busySession)
      await d.precheck("word_type_text")
      expect(busySession.setBusy).toHaveBeenCalledWith(false)
    })

    it("recovers unhealthy session", async () => {
      const sickSession = createMockSession({
        isUnhealthy: vi.fn().mockReturnValue(true),
      })
      const d = new SessionDirector(sickSession)
      await d.precheck("word_type_text")
      expect(sickSession.recover).toHaveBeenCalled()
      expect(d.circuitBreaker.getState()).toBe("CLOSED")
    })

    it("handles recovery failure gracefully", async () => {
      const failedSession = createMockSession({
        isAlive: vi.fn().mockReturnValue(false),
        recover: vi.fn().mockRejectedValue(new Error("recovery blew up")),
      })
      const d = new SessionDirector(failedSession)
      // Should not throw
      const result = await d.precheck("word_type_text")
      expect(result).toEqual({ ok: true })
    })

    it("rejects when circuit breaker is open and session cannot recover", async () => {
      const deadSession = createMockSession({
        isAlive: vi.fn().mockReturnValue(false),
        isUnhealthy: vi.fn().mockReturnValue(true),
        recover: vi.fn().mockRejectedValue(new Error("recovery failed")),
      })
      const d2 = new SessionDirector(deadSession)
      d2.circuitBreaker.onFailure()
      d2.circuitBreaker.onFailure()
      d2.circuitBreaker.onFailure()

      const result = await d2.precheck("word_type_text")
      expect(result).toEqual({ ok: false, error: expect.stringContaining("Circuit breaker is OPEN") })
    })

    it("forces reset circuit breaker when session is alive and not unhealthy", async () => {
      let checkCalled = false
      const cb = {
        check: vi.fn(() => {
          checkCalled = true
          throw new Error("open")
        }),
        forceReset: vi.fn(),
      } as any
      const d = new SessionDirector(mockSession)
      ;(d as any)._circuitBreaker = cb

      await d.precheck("word_type_text")
      expect(cb.forceReset).toHaveBeenCalled()
    })

    describe("precondition DOC", () => {
      it("approves when activeDoc is set", async () => {
        const sessionWithDoc = createMockSession({ activeDoc: { Name: "test.docx" } })
        const d = new SessionDirector(sessionWithDoc)
        const result = await d.precheck("word_type_text", "DOC")
        expect(result).toEqual({ ok: true })
      })

      it("rejects when no activeDoc and no auto-detect", async () => {
        const d = new SessionDirector(createMockSession({ activeDoc: null as any }))
        const result = await d.precheck("word_type_text", "DOC")
        expect(result).toEqual({ ok: false, error: expect.stringContaining("[NO_DOCUMENT]") })
      })

      it("auto-adopts ActiveDocument from session if available", async () => {
        const activeDoc = { FullName: "C:\\test\\doc.docx" }
        const adoptableSession = createMockSession({
          activeDoc: null,
          activeDocPath: null,
          tryAdoptActiveDoc: vi.fn(() => {
            adoptableSession.activeDoc = activeDoc
            adoptableSession.activeDocPath = "C:\\test\\doc.docx"
            return true
          }),
        })
        const appMgr = createMockAppManager()
        const d = new SessionDirector(adoptableSession, null, appMgr)
        const result = await d.precheck("word_type_text", "DOC")
        expect(result).toEqual({ ok: true })
        expect(adoptableSession.tryAdoptActiveDoc).toHaveBeenCalled()
        expect(appMgr.adoptDocument).toHaveBeenCalledWith("C:\\test\\doc.docx", activeDoc)
      })

      it("handles adoptDocument failure gracefully", async () => {
        let sessionRef: any
        const adoptableSession = createMockSession({
          activeDoc: null,
          activeDocPath: null,
          tryAdoptActiveDoc: vi.fn(() => {
            sessionRef.activeDoc = { FullName: "C:\\test\\doc.docx" }
            sessionRef.activeDocPath = "C:\\test\\doc.docx"
            return true
          }),
        })
        sessionRef = adoptableSession
        const failingAppMgr = { adoptDocument: vi.fn(() => { throw new Error("adopt failed") }) } as any
        const d = new SessionDirector(adoptableSession, null, failingAppMgr)
        const result = await d.precheck("word_type_text", "DOC")
        expect(result).toEqual({ ok: true })
      })
    })

    describe("precondition NO_DOC", () => {
      it("approves when no activeDoc", async () => {
        const d = new SessionDirector(createMockSession({ activeDoc: null as any }))
        const result = await d.precheck("word_stream_start", "NO_DOC")
        expect(result).toEqual({ ok: true })
      })

      it("rejects when activeDoc exists", async () => {
        const sessionWithDoc = createMockSession({ activeDoc: { Name: "test.docx" } })
        const d = new SessionDirector(sessionWithDoc)
        const result = await d.precheck("word_stream_start", "NO_DOC")
        expect(result).toEqual({ ok: false, error: expect.stringContaining("[DOC_ACTIVE]") })
      })
    })

    describe("precondition array", () => {
      it("handles empty array as no precondition", async () => {
        const d = new SessionDirector(mockSession)
        const result = await d.precheck("word_type_text", [])
        expect(result).toEqual({ ok: true })
      })

      it("uses first element for multi-precondition array", async () => {
        const d = new SessionDirector(createMockSession({ activeDoc: null as any }))
        const result = await d.precheck("word_type_text", ["NO_DOC"])
        expect(result).toEqual({ ok: true })
      })
    })
  })

  describe("captureStatusSuffix", () => {
    it("returns empty string when session is null", () => {
      const d = new SessionDirector(null)
      expect(d.captureStatusSuffix()).toBe("")
    })

    it("returns doc: none when no activeDoc", () => {
      const d = new SessionDirector(createMockSession({ activeDoc: null as any }))
      const suffix = d.captureStatusSuffix()
      expect(suffix).toContain("doc: none")
    })

    it("returns doc: untitled when no activeDocPath", () => {
      const d = new SessionDirector(createMockSession({ activeDoc: { Name: "Untitled" } as any, activeDocPath: null }))
      const suffix = d.captureStatusSuffix()
      expect(suffix).toContain("doc: untitled")
    })

    it("returns filename from activeDocPath", () => {
      const sessionWithPath = createMockSession({
        activeDoc: { Name: "report.docx" },
        activeDocPath: "C:\\docs\\report.docx",
      })
      const d = new SessionDirector(sessionWithPath)
      const suffix = d.captureStatusSuffix()
      expect(suffix).toContain('doc: "report.docx"')
    })

    it("includes [stream active] when streaming", () => {
      const d = new SessionDirector(mockSession)
      d.acquireStreamLock("word_stream_start")
      const suffix = d.captureStatusSuffix()
      expect(suffix).toContain("[stream active]")
    })

    it("includes [edit mode] when editing", () => {
      const d = new SessionDirector(mockSession)
      d.enterEditMode()
      const suffix = d.captureStatusSuffix()
      expect(suffix).toContain("[edit mode]")
    })

    it("gracefully handles errors", () => {
      const base = createMockSession()
      delete (base as any).activeDoc
      delete (base as any).activeDocPath
      const throwingSession = Object.defineProperty(base, "activeDoc", {
        get: () => { throw new Error("COM error") },
      })
      const d = new SessionDirector(throwingSession)
      expect(d.captureStatusSuffix()).toBe("")
    })
  })

  describe("markDirtyIfNeeded", () => {
    it("marks position map dirty for non-read-only tools", () => {
      const posMap = createMockPositionMap()
      const d = new SessionDirector(mockSession, posMap)
      d.markDirtyIfNeeded("word_type_text")
      expect(posMap.markDirty).toHaveBeenCalled()
    })

    it("skips markDirty for read-only tools", () => {
      const posMap = createMockPositionMap()
      const d = new SessionDirector(mockSession, posMap)
      d.markDirtyIfNeeded("word_get_text")
      expect(posMap.markDirty).not.toHaveBeenCalled()
    })

    it("does not crash when positionMap is null", () => {
      const d = new SessionDirector(mockSession, null)
      d.markDirtyIfNeeded("word_type_text")
    })
  })

  describe("schedulePositionRefresh", () => {
    it("delegates to positionMap.scheduleRefresh", () => {
      const posMap = createMockPositionMap()
      const d = new SessionDirector(mockSession, posMap)
      d.schedulePositionRefresh()
      expect(posMap.scheduleRefresh).toHaveBeenCalled()
    })

    it("does not crash when positionMap is null", () => {
      const d = new SessionDirector(mockSession, null)
      d.schedulePositionRefresh()
    })
  })

  describe("recoverSession", () => {
    it("recovers session and resets circuit breaker", async () => {
      const d = new SessionDirector(mockSession)
      d.circuitBreaker.onFailure()
      d.circuitBreaker.onFailure()
      d.circuitBreaker.onFailure()
      expect(d.circuitBreaker.isOpen()).toBe(true)

      await d.recoverSession()

      expect(mockSession.recover).toHaveBeenCalled()
      expect(d.circuitBreaker.isOpen()).toBe(false)
    })

    it("does nothing when session is null", async () => {
      const d = new SessionDirector(null)
      await d.recoverSession()
    })
  })

  describe("watchdog", () => {
    beforeEach(() => { vi.useFakeTimers() })
    afterEach(() => { vi.useRealTimers() })

    it("startWatchdog and stopWatchdog cycle cleanly", () => {
      const d = new SessionDirector(mockSession)
      d.startWatchdog()
      d.stopWatchdog()
    })

    it("startWatchdog is idempotent", () => {
      const d = new SessionDirector(mockSession)
      d.startWatchdog()
      d.startWatchdog()
      d.stopWatchdog()
    })

    it("stopWatchdog when not started does not crash", () => {
      const d = new SessionDirector(mockSession)
      d.stopWatchdog()
    })

    it("watchdog recovers unhealthy session", async () => {
      const unhealthySession = createMockSession({
        isUnhealthy: vi.fn().mockReturnValue(true),
        recover: vi.fn().mockResolvedValue(undefined),
      })
      const d = new SessionDirector(unhealthySession)
      d.startWatchdog()

      vi.advanceTimersByTime(5000)
      await vi.waitFor(() => {
        expect(unhealthySession.recover).toHaveBeenCalled()
      })
      d.stopWatchdog()
    })

    it("watchdog recovers hung COM calls", async () => {
      const hungSession = createMockSession({
        getLastCallStart: vi.fn().mockReturnValue(Date.now() - 60000),
        recover: vi.fn().mockResolvedValue(undefined),
      })
      const d = new SessionDirector(hungSession)
      d.startWatchdog()

      vi.advanceTimersByTime(5000)
      await vi.waitFor(() => {
        expect(hungSession.recover).toHaveBeenCalled()
      })
      d.stopWatchdog()
    })
  })
})
