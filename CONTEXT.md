# 领域上下文

## 核心实体

**WordSession** — 一条到 WINWORD.EXE 的 COM 连接。管理 Word.Application 的生命周期（启动/退出）、健康状态追踪（isAlive/healthCheck/recover）以及当前活跃文档指针。所有字处理模块均依赖于此会话。

**Document** — 由一个已解析路径追踪的 Word 文档。在 DocumentRegistry 中具有活跃状态（会话的 activeDoc）、路径和保存状态。

**Selection** — 当前文档中的光标位置/文本选区。所有编辑操作都通过选区进行。

**PositionMap** — 文档标题位置和表格位置的进程内缓存，由单次 COM 批量读取构建。支持 O(1) 查找"转至标题 X"操作。

## 核心值对象

**DocumentInfo** — wordCount、paragraphCount、pageCount、characterCount、sectionCount、saved

**HeadingEntry** — level、text、paragraphIndex

**Locator** — 使用 matchModes 和偏移量的 heading/paragraph/table/bookmark 定位器的判别联合

**ResolvedPosition** — found、paragraphIndex、headingContext、tableIndex

**WordMcpError** — code、recoverable 标志、recoveryHint 字符串

**Document States** — NO_WORD、NO_DOC、DOC_ACTIVE、DIALOG

## 模块

**CursorPosition** — 光标状态管理器。确保对 `ensureMainBody()` 的每次调用后，选区位于文档主正文中（而非页眉/页脚/表格/形状）。在后几类上下文之间转换时，通过标记追踪位置。在 `WordBase` 中引用，并发包给所有编辑器类。

**ChartDataBridge** — 图表数据处理子进程管理器。通过 IPC 派生一个独立的 Node.js 子进程（`chart-data-worker.js`）以设置图表数据。管理空闲超时（60s auto-kill）和每条操作的超时（15s）。在 `WordMediaEditor` 中引用，并通过 `IChartDataBridge` 接缝注入。

**WordStreamWriter** — 流式文档写入引擎。支持在单个 COM 会话中分块写入 Markdown 内容，内容在 Word 窗口中即时呈现。配合 `word_stream_start/block/end` 工具链，实现零样板代码的文档生成。自动管理样式继承、光标定位和批处理刷新。

**SessionDirector** — 会话编排器。管理 `word_document`/`word_create`/`word_get_status` 等文档级操作的路由，协调 WordSession 与 DocumentRegistry 之间的交互。在 `server/` 层提供统一的文档入口抽象。

**ServerContext** — 一组 MCP 工具管道所需的共享会话依赖项。包含对 `IWordSession`、`PositionMap`、`DocumentRegistry` 的引用。不执行可能被隐藏的模块级全局连接；所有工具注册函数均将其作为显式参数接收。

**ProcessMonitor** — 子进程健康看门狗。监控 `chart-data-worker` 等派生进程的生命周期，管理空闲超时和异常终止后的资源清理。在 `word/` 层提供进程级保障。

**StreamingMarkdownWriter** — 流式文档写入引擎（`word/word-stream-writer.ts`）。支持在单个 COM 会话中分块写入 Markdown 内容，内容在 Word 窗口中即时呈现。配合 `word_stream_start/block/end` 工具链，实现零样板代码的文档生成。自动管理样式继承、光标定位和批处理刷新。支持 `baseStyleProfile` 在文档创建时预配置样式定义（字体、段落），markdown 写入时自动继承。

**COM Proxy** — 类型化的 COM 对象包装层（`word/com-proxy/`）。`DocumentProxy` / `SelectionProxy` / `RangeProxy` 分别封装最常用的三个 COM 对象，消除全库 `Record<string, unknown>` 类型不安全。在 Session 层通过 `getDocProxy()` / `getSelectionProxy()` / `wrapRange()` 懒加载工厂构造。子对象（Find、Font、Shading 等）保持 Record 逃生口，待后续细分。

## 架构概览

```
┌─────────────────────────────────────────────────────────────────┐
│                        MCP Client (stdio)                        │
└──────────────────────────────┬──────────────────────────────────┘
                               │
                    ┌──────────▼──────────┐
                    │     parent.js        │  ← 看门狗进程，监控/重启子进程
                    │  (Watchdog + spawn)  │
                    └──────────┬──────────┘
                               │
                    ┌──────────▼──────────┐
                    │      child.js        │  ← MCP Server 主进程
                    │  (createServer)      │
                    └──────────┬──────────┘
                               │
              ┌────────────────┼────────────────┐
              │                │                │
     ┌────────▼──────┐ ┌──────▼───────┐ ┌──────▼────────┐
     │ Server Layer  │ │  Word Layer  │ │ Security Layer│
     │               │ │              │ │               │
     │ SessionDirector│ │ WordSession  │ │ SecurityManager│
     │ ServerContext │ │ COM Proxies  │ │ PathSanitizer │
     │ Tool Modules  │ │ ContentWriter│ │ RateLimiter   │
     │ Prompts       │ │ Stream Writer│ │ Audit         │
     │ Resources     │ │ PositionMap  │ │               │
     └───────────────┘ └──────┬───────┘ └───────────────┘
                              │
                    ┌─────────▼─────────┐
                    │   WINWORD.EXE     │  ← COM 自动化
                    │   (winax addon)   │
                    └───────────────────┘
```

## 数据流

1. **MCP 客户端** 通过 stdio 发送 JSON-RPC 请求
2. **parent.js** 看门狗接收请求，转发给 child.js
3. **child.js** 中的 `createServer` 初始化所有模块
4. **SessionDirector** 执行前置检查（状态机、速率限制、电路断路器）
5. **工具处理器** 调用 Word 层模块执行实际操作
6. **COM Proxy** 层将操作转换为 Word COM API 调用
7. 结果通过 stdio 返回给 MCP 客户端

## 关键设计模式

### 流式文档写入
```
word_stream_start({baseStyleProfile}) → word_stream_block(×N) → word_stream_end()
```
- 一次性预配样式，后续内容自动继承
- 批处理刷新（200ms 批次），避免 COM 调用过载
- Markdown 解析 → 分段渲染 → Word 即时呈现

### 工具注册管道
```
createRegTool → mcpCall 中间件:
  速率限制 → 前置检查 → ensureReady → 超时 → 处理器 → 审计 → 电路断路器 → 恢复
```

### 光标管理
- `ContextSanitizer` 确保编辑前光标在主正文
- 页眉/页脚/表格/水印操作后自动恢复光标位置
- 使用 `wasInNonBody` 标记追踪上下文转换

### 会话恢复
- `WordSession.comCall()` 包装所有 COM 调用
- 瞬态错误（RPC_E_CALL_REJECTED 等）自动标记不健康
- `SessionDirector` 看门狗检测挂起的 COM 调用并触发恢复
- 恢复流程：杀进程 → 清理 → 重建 COM 会话

### 位置缓存
- `PositionMap` 缓存标题、表格、段落位置
- 使用二进制搜索实现 O(1) 查找
- 文档修改后自动标记脏数据，延迟刷新

## 安全机制

- **路径消毒** — 防止目录遍历攻击，限制允许访问的目录
- **速率限制** — 防止工具调用过载
- **Zod 验证** — 所有输入参数类型安全验证
- **宏禁用** — `AutomationSecurity = 3`（高安全级别）
- **审计日志** — 记录所有工具调用和错误
