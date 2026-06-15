# 修复：流锁泄漏导致 word_document 被阻挡

## 问题
`word_stream_start` 后如果不调用 `word_stream_end`，`_currentPath` 卡在 `"streaming"`，
后续 `word_document` 被 precheck 拦截。

## 根因
1. `word_close` 在 `STREAM_BLOCKED_TOOLS` 中，流式会话中无法调用
2. `word_close` handler 只调了 `exitEditMode()`（处理 `"editing"`），没调 `releaseStreamLock()`（处理 `"streaming"`）

## 修复（2 处修改）

### 1. `src/server/session-director.ts:23`
```diff
-  "word_document", "word_open", "word_close", "word_quit",
+  "word_document", "word_open", "word_quit",
```
把 `word_close` 从 `STREAM_BLOCKED_TOOLS` 移除，允许流式会话中关闭文档。

### 2. `src/server/tools/document.ts:129-134`
```diff
  mcpCall(security, context, "word_close", async ({ saveChanges }) => {
    await appManager.closeDocument(saveChanges)
    context.director?.exitEditMode()
+   context.director?.releaseStreamLock()
    const saved = saveChanges ?? false
    return ...
  }),
```
`word_close` 后调用 `releaseStreamLock()`，无论当前是 `"editing"` 还是 `"streaming"` 都回到 `"idle"`。

## 验证
- `npx tsc --noEmit` — 无错误
- `npm test` — 110 测试通过
- 新建文档: `word_document()` → 创建空白无标题文档
- 流式退出: `word_stream_start` → `word_close` → `word_document()` 可用
