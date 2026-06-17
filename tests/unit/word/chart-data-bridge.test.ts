import { describe, it, expect, vi, beforeEach, afterEach } from "vitest"
import { ChartDataBridge } from "../../../src/word/chart-data-bridge.js"

describe("ChartDataBridge", () => {
  let messageCb: ((msg: unknown) => void) | undefined
  let mockWorker: ReturnType<typeof vi.fn> & {
    send: ReturnType<typeof vi.fn>
    on: ReturnType<typeof vi.fn>
    kill: ReturnType<typeof vi.fn>
    killed: boolean
    stdout: { on: ReturnType<typeof vi.fn>; removeAllListeners: ReturnType<typeof vi.fn> }
  }
  let mockFork: ReturnType<typeof vi.fn>
  let bridge: ChartDataBridge

  function makeWorker(): typeof mockWorker {
    messageCb = undefined
    const worker = {
      send: vi.fn((msg: unknown) => {
        const { id } = msg as { id: number }
        setImmediate(() => messageCb?.({ id, result: { ok: true, series: 1 } }))
      }),
      on: vi.fn((event: string, cb: Function) => {
        if (event === "message") messageCb = cb as (msg: unknown) => void
        return worker
      }),
      kill: vi.fn(),
      killed: false,
      stdout: {
        on: vi.fn(),
        removeAllListeners: vi.fn(),
      },
    }
    return worker as any
  }

  beforeEach(() => {
    mockWorker = makeWorker()
    mockFork = vi.fn(() => mockWorker)
    bridge = new ChartDataBridge({ forkFn: mockFork as any })
  })

  afterEach(() => {
    bridge.dispose()
  })

  it("creates worker on first setChartData call", async () => {
    const promise = bridge.setChartData("doc", 1, [["A", 1]])
    await expect(promise).resolves.toEqual({ ok: true, series: 1 })
    expect(mockFork).toHaveBeenCalledTimes(1)
  })

  it("reuses existing worker", async () => {
    await bridge.setChartData("doc", 1, [["A", 1]])
    mockFork.mockClear()
    await bridge.setChartData("doc", 2, [["B", 2]])
    expect(mockFork).not.toHaveBeenCalled()
  })

  it("sends message with docName, inlineIndex and data", async () => {
    const data: (string | number)[][] = [["A", "B"], [1, 2]]
    await bridge.setChartData("myDoc", 3, data)
    expect(mockWorker.send).toHaveBeenCalledWith(
      expect.objectContaining({
        params: { docName: "myDoc", inlineIndex: 3, data },
      }),
    )
  })

  it("resolves with ok:true on success response", async () => {
    const result = await bridge.setChartData("doc", 1, [["X", 1]])
    expect(result).toEqual({ ok: true, series: 1 })
  })

  it("resolves with ok:false on worker error response", async () => {
    mockWorker.send = vi.fn((msg: unknown) => {
      const { id } = msg as { id: number }
      setImmediate(() => messageCb?.({ id, error: "FAILED" }))
    })
    const result = await bridge.setChartData("doc", 1, [["A", 1]])
    expect(result).toEqual({ ok: false, series: 1 })
  })

  it("dispose sends shutdown message and removes listeners", async () => {
    await bridge.setChartData("doc", 1, [["A", 1]])
    bridge.dispose()
    expect(mockWorker.send).toHaveBeenCalledWith({ id: -1, params: null })
    expect(mockWorker.stdout.removeAllListeners).toHaveBeenCalled()
  })

  it("creates new worker after dispose", async () => {
    await bridge.setChartData("doc", 1, [["A", 1]])
    bridge.dispose()
    mockFork.mockClear()
    mockWorker.killed = true
    const w2 = makeWorker()
    mockFork.mockReturnValue(w2)
    await bridge.setChartData("doc", 2, [["B", 2]])
    expect(mockFork).toHaveBeenCalledTimes(1)
    expect(w2.send).toHaveBeenCalled()
  })

  it("shares one worker across multiple concurrent calls", async () => {
    let callCount = 0
    mockWorker.send = vi.fn((msg: unknown) => {
      const { id } = msg as { id: number }
      callCount++
      setImmediate(() => messageCb?.({ id, result: { ok: true, series: callCount } }))
    })
    const p1 = bridge.setChartData("doc", 1, [["A", 1]])
    const p2 = bridge.setChartData("doc", 2, [["B", 2]])
    await expect(Promise.all([p1, p2])).resolves.toHaveLength(2)
    expect(mockFork).toHaveBeenCalledTimes(1)
  })

  it("returns ok:false when ensureWorker throws (throttled)", async () => {
    bridge.dispose()
    const badFork = vi.fn(() => { throw new Error("fail") })
    const bridge2 = new ChartDataBridge({ forkFn: badFork as any })
    const result = await bridge2.setChartData("doc", 1, [["A", 1]])
    expect(result).toEqual({ ok: false, series: 1 })
    bridge2.dispose()
  })
})
