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
