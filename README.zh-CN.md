> [English](./README.md) | [中文](./README.zh-CN.md)

# Word MCP Server

[![CI](https://github.com/HelloWorld-Open/word-mcp-server/actions/workflows/ci.yml/badge.svg)](https://github.com/HelloWorld-Open/word-mcp-server/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Node](https://img.shields.io/badge/node-%3E%3D20-brightgreen)](https://nodejs.org)
[![Platform](https://img.shields.io/badge/platform-Windows-blue)]()

通过 MCP 协议实时操控 Microsoft Word 的服务器。Word 以可见窗口运行，所有操作即时呈现。

```
此工具仅供已有 Word 授权的用户合法使用。
不包含、不破解、不分发 Microsoft Office 软件。
```

---

## 功能一览

- **110+ 个 MCP 工具** — 文档生命周期、内容编排、格式排版、表格、图表、媒体、结构、剪贴板、Manager 高层 API、语义定位、变量替换
- **实时可见** — Word 在前台打开，每条操作即时呈现
- **Ctrl+Z 友好** — 全部操作走 COM API，撤销栈正常
- **断线不丢** — MCP 断开后 Word 窗口保持打开，编辑不丢失
- **自动备份** — 每次保存自动创建 `.bak`
- **图表数据隔离** — 子进程 + 15s 超时，防止 Excel COM 挂死

## 快速开始

```bash
git clone https://github.com/HelloWorld-Open/word-mcp-server.git
cd word-mcp-server
npm install && npm run build
```

添加到 MCP 客户端配置：

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

> `npm install` 时需要 MSVC Build Tools 编译 winax 原生模块。如报错请安装 Visual Studio 2022 Build Tools（选择"使用 C++ 的桌面开发"工作负载）。

## 使用示例

### Markdown 写入格式化内容

```md
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
```

### 批处理模式加速

```js
word_batch_start   // 加速 3-10 倍
// ... 多个操作 ...
word_batch_end     // 刷新窗口
```

### 高层 Manager API

```js
word_mgr_create_document    // 自动关闭 + 页面设置
word_mgr_write_content      // 自动管理光标和重绘
word_mgr_insert_table       // 自动段落分隔
word_mgr_add_bookmark       // 自动深度复位光标
word_mgr_save               // 保存 + 可选导出 PDF
```

## 完整工具参考

详见 **[TOOLS.zh-CN.md](./TOOLS.zh-CN.md)** — 110+ 个工具，涵盖 12 个模块。

## 架构

### 通信流程

```
MCP Client (Claude Desktop 等)
    │  JSON-RPC over stdio
    ▼
build/parent.js  (看门狗 — 30s 超时，崩溃重启)
    │  spawn + 管道转发
    ▼
build/child.js   (McpServer)
    │  winax COM Automation（原生调用）
    ├─ 主线程：110+ 个工具
    └─ 子进程：图表数据设置（fork + 15s 超时）
            │
            ▼
        WINWORD.EXE  (可见窗口)
```

### 源码结构

```
src/
├── index.ts       # 入口
├── parent.ts      # 看门狗父进程
├── child.ts       # MCP 服务器进程
├── manager/       # 高层 API（文档建造者模式）
├── word/          # COM 适配层（通过 winax）
├── security/      # 4 层安全防御
└── server/        # MCP 协议层（110+ 工具、提示词）
```

## 安全机制（4 层防御）

- **路径消毒** — 7 种攻击向量防护（目录穿越、设备路径、网络路径、跨盘符、系统目录、NTFS 数据流、白名单）
- **宏保护** — 启动时强制禁用宏（`AutomationSecurity = 3`）
- **参数校验** — 所有工具输入经 Zod 运行时校验
- **错误脱敏** — 路径替换为占位符，不泄露内部细节

## 测试

**116 个测试，全部通过** ✅ — 9 个单元测试 + 1 个集成测试。

单元测试 mock 了 winax 层，不需要真实 Word 环境。覆盖率：Markdown 解析器 33、语义定位 21、文本编辑器 16、会话 13、安全层 30。

```bash
npm test
```

## 许可

[MIT](./LICENSE)
