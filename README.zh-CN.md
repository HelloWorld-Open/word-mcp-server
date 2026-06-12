> [English](./README.md) | [中文](./README.zh-CN.md)

# Word MCP Server

通过 MCP 协议实时操控 Microsoft Word 的服务器。Word 以可见窗口运行，所有操作实时可见。

```
此工具仅供已有 Word 授权的用户合法使用。
不包含、不破解、不分发 Microsoft Office 软件。
```

---

## 功能一览

- **110+ 个 MCP 工具** — 覆盖文档生命周期、内容输入、格式编排、表格、图表、媒体、文档结构、选区与剪贴板、高层 Manager 封装、语义定位、变量替换
- **实时可见** — Word 窗口在前台打开，每条指令的效果即时呈现
- **Ctrl+Z 友好** — 所有操作通过 COM API 执行，Word 撤销栈正常工作
- **断线不丢** — MCP 断开连接后，Word 窗口**保持打开**，已编辑文档不会丢失
- **自动备份** — 每次保存自动创建 `.bak` 备份文件
- **图表数据安全写入** — 通过子进程隔离 + 超时保底，避免 Excel COM 卡死主进程
- **格式自动推导** — 根据文件扩展名自动选择保存格式（docx/pdf/rtf/txt 等）

---

## 环境要求

| 依赖 | 版本要求 |
|------|---------|
| 操作系统 | Windows（必须，因依赖 Word COM 组件） |
| Microsoft Word | Office 2016 或更高版本 |
| Node.js | 18+ |
| MSVC Build Tools | 仅在首次 `npm install` 时编译 winax 需要 |

## 快速开始

### 1. 安装构建

```bash
cd word-mcp-server
npm install
npm run build
```

> **注意**：`npm install` 时需要 MSVC Build Tools 编译 winax 原生模块。
> 如报错请安装 Visual Studio 2022 Build Tools（选择"使用 C++ 的桌面开发"工作负载）。

### 2. 配置 MCP 客户端

```json
{
  "mcpServers": {
    "word": {
      "command": "node",
      "args": ["D:\\path\\to\\word-mcp-server\\build\\parent.js"]
    }
  }
}
```

### 3. 启动

客户端连接后，第一次调用任意工具时 Word 会自动启动。

---

## 工具参考（110+ 个）

### 📂 文档生命周期

| 工具 | 描述 | 必需参数 |
|------|------|---------|
| `word_get_status` | 查询当前 Word/文档状态（NO_WORD/NO_DOC/DOC_ACTIVE/DIALOG） | — |
| `word_document` | 通用文档入口：按路径打开/创建，不带路径则用当前文档 | `path`（可选） |
| `word_create` | 创建新文档（可见） | — |
| `word_create_from_template` | 从 .dotx 模板创建新文档（继承样式/页眉页脚/水印） | `path` |
| `word_open` | 打开已有 .docx 文件 | `path` |
| `word_save` | 保存当前文档（自动 .bak 备份） | — |
| `word_save_as` | 另存为新文件（自动适配格式） | `path` |
| `word_close` | 关闭当前文档（可选是否保存） | — |
| `word_get_info` | 获取文档统计信息（字数/页数/段落等） | — |
| `word_get_structure` | 获取文档标题大纲（含段落索引） | — |
| `word_quit` | 退出 Microsoft Word（不丢失已保存文档） | — |

### ✏️ 内容输入

| 工具 | 描述 | 必需参数 |
|------|------|---------|
| `word_type_text` | 在光标位置输入文字（smooth/instant 两种模式） | `text` |
| `word_insert_paragraph` | 插入段落换行 | — |
| `word_insert_page_break` | 插入分页符 | — |
| `word_insert_horizontal_line` | 插入水平分割线 | — |
| `word_insert_list` | 插入项目符号或编号列表 | `type`, `items` |
| `word_set_columns` | 设置当前节的分栏数（如报纸排版） | `count` |
| `word_insert_section_break` | 插入分节符（nextPage/continuous/evenPage/oddPage） | — |
| `word_insert_file` | 在光标位置插入另一个 .docx 文件的内容（合并文档） | `path` |
| `word_backspace` | 删除光标前字符（模拟退格键） | — |
| `word_write_markdown` | 使用 Markdown 语法写入格式化内容（标题/粗体/斜体/删除线/代码/链接/列表/表格/引用/分隔线） | `markdown` |
| `word_get_text` | 读取全文文本 | `maxLength`（可选） |
| `word_get_paragraph` | 读取指定段落文本 | `index` |
| `word_get_table_data` | 提取表格内容为结构化数据（行列网格） | `index`（可选，默认1） |
| `word_get_comments` | 列出所有批注及作者信息 | — |
| `word_get_bookmarks` | 列出所有书签 | — |
| `word_get_lists` | 列出所有项目/编号列表（含层级） | — |
| `word_get_sections` | 列出所有节及页面设置（方向/栏数/纸张大小） | — |
| `word_export_to_pdf` | 将文档导出为 PDF | `path`（可选） |

### 🔗 链接、引用与批注

| 工具 | 描述 | 必需参数 |
|------|------|---------|
| `word_add_hyperlink` | 添加超链接（支持 screenTip 和 subAddress） | `text`, `address` |
| `word_add_footnote` | 添加脚注（Word 自动编号） | `text` |
| `word_add_comment` | 添加批注 | `text` |

### 🔍 查找、选区与导航

| 工具 | 描述 | 必需参数 |
|------|------|---------|
| `word_find_text` | 查找文本并将光标定位到匹配处 | `findText` |
| `word_find_replace` | 查找替换文本（支持大小写/全词/全部替换） | `findText`, `replaceWith` |
| `word_go_to` | 跳转到指定位置（页/节/行/书签/末尾） | — |
| `word_go_to_paragraph` | 按段落索引跳转（1-based） | `index` |
| `word_select_all` | 全选文档内容 | — |
| `word_select_text` | 按字符范围选择文本 | `start`, `length` |
| `word_select_current_word` | 选择光标所在词 | — |
| `word_select_current_paragraph` | 选择光标所在段落 | — |
| `word_delete` | 删除当前选中内容（模拟 Delete 键） | — |

### 📋 剪贴板与撤销

| 工具 | 描述 | 必需参数 |
|------|------|---------|
| `word_copy` | 复制选中内容到剪贴板 | — |
| `word_cut` | 剪切选中内容到剪贴板 | — |
| `word_paste` | 从剪贴板粘贴（在光标位置） | — |
| `word_undo` | 撤销上一步操作 | — |
| `word_redo` | 重做已撤销的操作 | — |
| `word_get_cursor_info` | 获取光标位置/选区信息 | — |

### 🎨 格式编排

| 工具 | 描述 | 可设参数 |
|------|------|---------|
| `word_set_font` | 设置字体 | name, size, bold, italic, underline, color, strikethrough, superscript, subscript |
| `word_set_paragraph` | 设置段落格式 | alignment, indent, spacing, lineSpacing |
| `word_apply_style` | 应用命名样式（如"Heading 1"） | `styleName` |
| `word_set_page_setup` | 页面设置 | margins, orientation, pageSize |
| `word_set_properties` | 文档属性 | title, author, subject, keywords, comments |
| `word_list_styles` | 列出文档中的可用样式 | — |

### 📊 表格

| 工具 | 描述 | 必需参数 |
|------|------|---------|
| `word_insert_table` | 插入表格（可选填充数据） | `rows`, `columns` |
| `word_edit_cell` | 修改单元格内容 | `row`, `column`, `text` |
| `word_edit_cells` | 批量填充多个单元格（二维数组） | `data` |
| `word_add_table_row` | 添加行（可选填充数据） | — |
| `word_delete_table_row` | 删除行 | `rowIndex` |
| `word_add_table_column` | 添加列（在某列左侧追加或在末尾追加） | — |
| `word_delete_table_column` | 删除列 | `column` |
| `word_set_table_borders` | 设置表格边框样式/颜色/宽度 | — |
| `word_set_table_shading` | 设置表格或首行背景色 | `color` |
| `word_merge_table_cells` | 合并指定范围的单元格 | `rowStart`, `colStart`, `rowEnd`, `colEnd` |
| `word_set_column_width` | 设置列宽（单位：磅） | `column`, `width` |
| `word_set_row_height` | 设置行高（单位：磅） | `row`, `height` |
| `word_set_cell_font` | 设置指定单元格的字体格式 | `row`, `column` |
| `word_set_cell_vertical_alignment` | 设置单元格垂直对齐（上/中/下） | `row`, `column`, `alignment` |
| `word_apply_table_style` | 应用内置 Word 表格样式 | `styleName` |
| `word_table_to_text` | 将表格转换为纯文本（移除表格结构） | — |
| `word_text_to_table` | 将选中的分隔文本转换为表格 | — |

### 🖼️ 媒体

| 工具 | 描述 | 必需参数 |
|------|------|---------|
| `word_insert_image` | 插入图片文件 | `imagePath` |
| `word_insert_chart` | 插入图表（column/bar/line/pie/area，子进程隔离写入数据） | `type`, `data` |
| `word_insert_textbox` | 插入文本框（支持位置/大小/方向） | `text` |

### 📄 文档结构

| 工具 | 描述 | 必需参数 |
|------|------|---------|
| `word_set_header` | 设置页眉 | `text` |
| `word_set_footer` | 设置页脚 | `text` |
| `word_set_page_numbers` | 插入页码 | `target`（header/footer） |
| `word_insert_toc` | 插入目录 | — |
| `word_add_bookmark` | 添加书签 | `name` |
| `word_set_watermark` | 设置或移除水印（如"DRAFT"） | `text` |

### ⚡ 批处理模式

| 工具 | 描述 | 必需参数 |
|------|------|---------|
| `word_batch_start` | 开启批处理模式：禁用 Word 重绘，操作加速 3-10 倍 | — |
| `word_batch_end` | 结束批处理模式：恢复重绘并刷新窗口 | — |

### 🏗️ Manager 层（高层文档建造者）

| 工具 | 描述 |
|------|------|
| `word_mgr_create_document` | 创建文档（自动关闭上一个 + 页面设置） |
| `word_mgr_write_content` | 写入 Markdown（自动管理光标和 ScreenUpdating） |
| `word_mgr_apply_heading` | 在文档末尾应用标题样式 |
| `word_mgr_set_header` / `word_mgr_set_footer` | 设置页眉/页脚（自动返回正文） |
| `word_mgr_set_page_numbers` | 插入页码（自动返回正文） |
| `word_mgr_insert_table/chart/image/list/textbox` | 插入元素（自动段落分隔） |
| `word_mgr_add_bookmark/comment/footnote/hyperlink` | 添加标记（自动深度光标复位） |
| `word_mgr_insert_section_break` | 分节符（自动返回正文） |
| `word_mgr_format_page` | 一键页面布局设置 |
| `word_mgr_set_watermark` | 设置水印（自动返回正文） |
| `word_mgr_save` | 保存（可选同时导出 PDF） |

### 🧭 语义定位

| 工具 | 描述 |
|------|------|
| `word_locate` | 解析标题/段落/表格/书签位置（只读） |
| `word_select_at` | 按语义位置移动光标 |
| `word_insert_at` | 在语义位置写入文本 |
| `word_write_markdown_at` | 在语义位置写入 Markdown |
| `word_edit_cell_at` | 在语义位置编辑表格单元格 |

### 🔄 变量替换

| 工具 | 描述 |
|------|------|
| `word_replace_variables` | 替换文档中的 {{placeholder}} 变量 |

## 使用示例

### 用 Markdown 写入格式化内容

使用 `word_write_markdown` 一次性写入带格式的内容：

``````md
# 章节标题
**粗体** 和 *斜体* 文本，包含 `行内代码`、~~删除线~~ 和 [链接](https://example.com)。

- 一级项目
  - 嵌套项目（带 **粗体**）
- 另一个项目

1. 第一步
2. 第二步

```
代码块中的预格式化文本
```

| 名称 | 数值 |
|------|------|
| Alpha | 100 |
| Beta | 200 |

> 引用文本

---
``````

### 辅助提示词

| 提示词 | 功能 |
|--------|------|
| `create_report` | 生成创建结构化报告的步骤指南 |
| `format_document` | 生成格式化现有文档的步骤指南 |

---

## 架构

```
src/
├── index.ts                          # 入口 — McpServer + stdio transport
├── parent.ts                         # 父进程看门狗（30s 超时，自动重启子进程）
├── child.ts                          # MCP 服务器进程（由看门狗父进程启动）
├── manager/                          # 文档建造者模式（高层 API）
│   ├── types.ts                      # 11 个复合操作参数接口
│   └── word-manager.ts               # 文档建造者（自动光标复位、上下文切换）
├── word/                             # Word COM 适配层（通过 winax 原生模块）
│   ├── session.ts                    # Winax COM 会话管理（惰性启动/自动恢复）
│   ├── types.ts                      # 共享类型定义
│   ├── application.ts                # 文档生命周期（创建/打开/保存/关闭/退出）
│   ├── document.ts                   # 文档信息、路径扫描
│   ├── document-registry.ts          # 文档注册表（路径→doc 映射/去重/外部关闭检测）
│   ├── process-monitor.ts            # WINWORD.EXE 进程健康监控
│   ├── word-base.ts                  # Word COM 操作抽象基类
│   ├── word-text-editor.ts           # 文本编辑操作（输入/查找/替换等）
│   ├── word-document-structure.ts    # 文档结构操作（页眉/页脚/目录/书签/水印）
│   ├── word-table-editor.ts          # 表格编辑操作（单元格/行/列/合并）
│   ├── word-media-editor.ts          # 媒体操作（图片/图表/文本框）
│   ├── formatting.ts                 # 格式操作（字体/段落/页面/样式）
│   ├── position-map.ts               # 语义定位系统（二分查找 + 脏标记刷新）
│   ├── variable-replacer.ts          # {{placeholder}} 变量替换引擎
│   ├── word-markdown.ts              # Markdown → Word 转换器（480 行，自研解析引擎）
│   └── chart-data-worker.ts          # 子进程：隔离设置图表数据（15s 超时）
├── security/                         # 安全与策略层（4 层纵深防御）
│   ├── errors.ts                     # 分层错误体系（6 种类型 + 脱敏处理）
│   ├── path-sanitizer.ts             # 路径消毒（7 种攻击向量防护）
│   ├── policy.ts                     # 安全策略配置 + SecurityManager
│   ├── rate-limiter.ts               # 滑动窗口速率限制
│   └── audit.ts                      # 操作审计日志
└── server/                           # MCP 协议层
    ├── tools/                        # 110+ 个工具注册（12 个模块）
    │   ├── helper.ts                 # mcpCall 封装：超时/审计/限流/错误处理
    │   ├── document.ts               # 11 个文档生命周期工具
    │   ├── content.ts                # 26 个内容/选区/剪贴板工具
    │   ├── formatting.ts             # 6 个格式编排工具
    │   ├── tables.ts                 # 17 个表格工具
    │   ├── media.ts                  # 3 个媒体工具
    │   ├── structure.ts              # 7 个文档结构工具
    │   ├── reader.ts                 # 8 个读取工具
    │   ├── batch.ts                  # 2 个批处理模式工具
    │   ├── markdown.ts               # 1 个 Markdown 辅助工具
    │   ├── semantic.ts               # 5 个语义定位工具（locate/select_at/insert_at 等）
    │   ├── variable.ts               # 1 个变量替换工具
    │   └── manager.ts                # 11 个高层 Manager 封装工具
    └── prompts/                      # 提示词注册
```

### 技术栈

- **COM 通信**：通过 `winax` 原生 Node.js 加载项直接调用 Word COM 接口
- **MCP 协议**：`@modelcontextprotocol/sdk` — StdioServerTransport
- **参数校验**：Zod 运行时校验 + JSON Schema

### 通信流程

```
MCP Client (Claude Desktop 等)
    │  JSON-RPC over stdio
    ▼
build/parent.js  (看门狗 — 30s 超时，崩溃自动重启)
    │  spawn + 管道转发 stdin/stdout
    ▼
build/child.js   (McpServer)
    │  winax COM Automation（原生调用）
    ├─ 主线程：文档/内容/格式/表格/媒体/结构/读取/批处理/Markdown/语义/变量/Manager 等 110+ 个工具
    └─ 子进程：图表数据设置（fork + 15s 超时隔离）
            │
            ▼
        WINWORD.EXE  (可见窗口)
```

> 不再依赖 PowerShell 桥接，直接通过 winax 调用 Word COM，延迟更低、稳定性更高。

---

## 安全机制（4 层防护）

### 第 1 层 — 路径消毒

所有文件路径输入经过严格检查：

- **目录穿越防护** — 拦截 `..` 和 `~` 路径
- **系统目录封锁** — `C:\Windows\` 等系统目录只读封锁
- **网络路径禁用** — `\\server\share` 格式默认拒绝
- **设备路径拦截** — `\\.\`, `\\?\` 等设备路径禁止访问
- **备用数据流拦截** — 包含 `:` 的 ADS 路径禁止访问
- **目录白名单** — 可选限制只允许访问特定目录
- **最大文件大小** — 超过 50MB 的文件拒绝操作

### 第 2 层 — 宏保护

- 启动时强制设置 `AutomationSecurity = 3`（**msoAutomationSecurityForceDisable**）
- 所有通过本服务器打开的文档，其中的宏和 ActiveX 控件**被强制禁用**
- 底层 COM 层始终阻止宏执行，不受策略配置影响

### 第 3 层 — 参数校验

- **所有工具输入** — 通过 Zod schema 校验（20+ 个 schema）
- **文本参数** — 限制最大长度（maxLength）
- **数值参数** — 限制取值范围（如 fontSize: 1–1638, rows: 1–500）
- **速率限制** — 可选限制单位时间调用频率（防误操作）

### 第 4 层 — 错误脱敏

- 所有暴露给 MCP 客户端的错误消息中，**文件路径被替换为占位符**
- 内部实现细节（COM 错误码、内存地址等）不会泄露
- 用户获得的错误信息包含操作建议（可恢复/不可恢复）

---

## 行为说明

| 场景 | 行为 |
|------|------|
| 首次调用工具 | Word 自动启动，窗口可见 |
| 调用 `word_quit` | Word 正常退出 |
| MCP 连接断开 | Word 窗口**保持打开**，不保存不退出，已编辑内容不丢失 |
| 保存未命名文档 | 自动保存到系统临时目录 `%TEMP%\word-mcp-doc.docx` |
| 手动关掉 Word 窗口 | 下次工具调用会自动重新打开 Word |
| 操作超时 | 默认 30 秒，可通过策略配置调整 |
| 图表设置数据 | 子进程隔离执行，15s 超时保底，不阻塞主线程 |

---

## 开发

```bash
# 安装依赖
npm install

# 构建
npm run build

# 开发模式（增量编译）
npx tsc --watch

# 运行测试
npm test

# 端到端手动测试（需要 Word 已安装）
node test-mcp.mjs
```

### 测试

使用 vitest 框架，共 **10 个测试文件**（9 单元 + 1 集成）— **116 个测试，全部通过** ✅

| 分类 | 测试文件 | 覆盖情况 |
|------|---------|---------|
| 安全层 | `errors`, `path-sanitizer`, `policy`, `rate-limiter`, `audit` | 5 文件，30 测试 |
| COM 会话 | `session` | 13 测试（mock winax） |
| Markdown 解析器 | `markdown` | 33 测试（纯函数，零 mock） |
| 文本编辑器 | `text-editor` | 16 测试（分段逻辑 + COM mock） |
| 语义定位 | `position-map` | 21 测试（定位算法 + 状态管理） |
| 集成测试 | `smoke` | 3 测试（MCP 协议 + 工具清单） |

单元测试 mock 了 winax 层，不依赖真实 Word 环境。集成测试需要 Word 安装但不需要运行，仅验证协议层。

---

## 常见问题

**Q: Word 启动报错「拒绝访问」？**
A: 以管理员身份运行 MCP 客户端。部分 COM 操作需要管理员权限。

**Q: 工具调用后 Word 没反应？**
A: 首次调用需等待 Word 启动（约 1–3 秒）。如果 30 秒无响应，检查任务管理器中有无 `WINWORD.EXE` 僵尸进程。

**Q: 保存时提示路径不允许？**
A: 检查路径是否在允许的目录白名单内，或包含系统目录关键字。

**Q: 能同时操作多个文档吗？**
A: 当前设计为单文档模式（`activeDoc`），每次操作影响当前激活的文档。`word_open` 或 `word_create` 会切换激活文档。

**Q: 打开包含宏的文档安全吗？**
A: 宏已被 COM 层强制禁用，无论文档中是否携带宏代码。

**Q: 插入图表卡死怎么办？**
A: 图表数据设置在子进程中执行（15 秒超时）。如果超时，图表仍会创建（带默认数据），不影响主线程。通常卡死原因是上次操作残留的僵尸 Excel 进程，`insertChart` 会自动检测并清理。

**Q: `npm install` 报错缺少 MSVC？**
A: 安装 Visual Studio 2022 Build Tools，选择"使用 C++ 的桌面开发"工作负载。

---

## 搭配相关 Skills 使用（可选）

Word MCP 可独立运行。如需让 AI 输出风格更统一（如固定排版规范），
可在项目中创建规则文件（如 `.opencode/skills/` 或其他 Agent 支持的方式），
让 AI 在生成文档时自动遵循。是否使用取决于你的需要，MCP 本身功能不依赖此配置。

---

## 相关项目

- [Model Context Protocol](https://modelcontextprotocol.io) — MCP 协议规范
- [@modelcontextprotocol/sdk](https://github.com/modelcontextprotocol/typescript-sdk) — 官方 TypeScript SDK
- [winax](https://github.com/demchenkoe/winax) — Node.js Windows COM 原生模块

## 许可

MIT
