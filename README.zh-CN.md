> [English](./README.md) | [中文](./README.zh-CN.md)

# Word MCP Server

<p align="center">
  <video src="https://github.com/user-attachments/assets/0f8d144b-62bd-4ee7-9c01-c0f71cca05b6" autoplay loop muted playsinline width="300"></video>
</p>

<p align="center">
  <a href="https://github.com/HelloWorld-Open/word-mcp-server/actions/workflows/ci.yml">
    <img src="https://github.com/HelloWorld-Open/word-mcp-server/actions/workflows/ci.yml/badge.svg" alt="CI">
  </a>
  <a href="https://opensource.org/licenses/MIT">
    <img src="https://img.shields.io/badge/License-MIT-yellow.svg" alt="License: MIT">
  </a>
  <a href="https://nodejs.org">
    <img src="https://img.shields.io/badge/node-%3E%3D20-brightgreen" alt="Node">
  </a>
  <img src="https://img.shields.io/badge/platform-Windows-blue" alt="Platform">
</p>

## 是什么

让你的 AI 智能体直接操控 Microsoft Word。108 个工具，覆盖创建、编辑、排版、表格、图表、图片、书签、页眉页脚、批注等全流程。Word 以可见窗口运行，每条修改即时呈现，Ctrl+Z 全程有效。

> 需要已有 Microsoft Word 授权。不包含、不破解、不分发 Office 软件。

## 做什么

| 场景 | 描述 |
|------|------|
| 周报月报 | Agent 拉取数据，写报告，排版，保存，一气呵成 |
| 合同草拟 | 模板填入变量，生成正式合同，导出 PDF |
| 学术论文 | 撰写结构化内容，插入表格图表，格式化引用 |
| 自动化测试 | CI 流程中自动创建 Word 文档用于测试验证 |

## 兼容生态

| 生态 | 智能体 |
|------|--------|
| Claude | Claude Desktop、Claude Code |
| OpenAI | ChatGPT Desktop |
| IDE | Cursor、VS Code / GitHub Copilot、Trae |
| AI 终端 | OpenCode |
| 更多 | 任意 stdio MCP 客户端 |

## 为什么需要

- **108 个工具** — 文本、表格、图表、图片、页眉页脚、批注、书签，精准操控每个细节
- **流式 Markdown** — `start → block → end` 三步写文档，即时预览，字体/样式自动继承
- **Manager API** — 一条指令搞定复杂流程（页眉 → 页码 → 表格 → 水印），Agent 不用管光标位置
- **完全可撤销** — COM API 原生操作，搞砸了 Ctrl+Z 就行
- **安全内置** — 路径消毒、宏强制禁用、Zod 输入校验、限流、审计日志

---

### ⚡ 快速开始

```bash
git clone https://github.com/HelloWorld-Open/word-mcp-server.git
cd word-mcp-server
npm install && npm run build
```

> 需要 MSVC Build Tools 编译 winax 原生模块。如失败请安装 VS 2022 Build Tools（"使用 C++ 的桌面开发"工作负载）。

添加到 MCP 客户端配置：

```json
{
  "mcpServers": {
    "word": {
      "command": "node",
      "args": ["path\\to\\word-mcp-server\\build\\parent.js"]
    }
  }
}
```

然后对 AI 说：*"帮我写一份周报到 Word"*。

### 📖 更多

- [TOOLS.zh-CN.md](./TOOLS.zh-CN.md) — 108 工具速查
- [CONTEXT.md](./CONTEXT.md) — 架构与设计
- [.env.example](./.env.example) — 配置项
- [MIT License](./LICENSE)
