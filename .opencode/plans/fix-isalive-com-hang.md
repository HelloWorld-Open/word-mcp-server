# 修复：isAlive() 裸 COM Version 调用导致 30s 挂起

## 问题
`word_insert_at(by:"cursor")` 在空白文档上超时 30s。
每次工具调用都经过 `session.application` → `ensureAlive()` → `isAlive()` → `this.word.Version`，
这个 COM 属性访问是裸调用（无 comCall 包裹），当 Word 显示模态对话框时阻塞 30s。

## 根因
`session.ts:170-171` 中 `this.word.Version` 是裸 COM 调用。虽然前面有 `monitor.isAlive()` 进程检查，
但如果进程活着但 COM 被对话框阻塞，仍然会挂起 30s。

## 修复（1 处修改）

**`src/word/session.ts:158-182`**

```diff
  isAlive(): boolean {
    if (!this.word) return false

    // 进程级检查（无 COM 调用，避免死进程 COM 挂起 30-60s）
    if (!this.monitor.isAlive()) {
      this.onLog?.("warn", "Process-level check: WINWORD.EXE not found — stale COM proxy")
      this.word = null
      this._unhealthy = true
      return false
    }

+   this._unhealthy = false
+   return true
-
-    // COM 存活检查: Version 属性（~5ms，不比 tasklist 慢且不依赖外部进程）
-    try {
-      const v = (this.word as Record<string, unknown>).Version
-      if (v === undefined) {
-        this.onLog?.("warn", "COM Version returned undefined — stale COM state")
-        throw new Error("stale com")
-      }
-      this._unhealthy = false
-      return true
-    } catch {
-      this.word = null
-      this._unhealthy = true
-      return false
-    }
  }
```

## 理由
- `monitor.isAlive()` 每秒用 `tasklist` 检查一次，准确可靠
- 进程活着 → COM 连接大概率可用，不需要额外的 `Version` 探针
- 移除后每次 `isAlive()` 省 1 次 COM 调用（~5ms），路径上更轻量
- 极端情况（进程活着但 COM 断开），后续真实 COM 调用会在被 `comCall` 包裹的地方失败，产生有意义的错误

## 验证
- `npx tsc --noEmit` — 无错误
- `npm test` — 110 测试通过
- `word_document()` → 创建新文档 → `word_insert_at(by:"cursor")` 不再超时
