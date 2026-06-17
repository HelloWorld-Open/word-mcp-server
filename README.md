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

## What it is

Let your AI agent directly control Microsoft Word. 80 tools covering the full workflow — create, edit, format, tables, charts, images, bookmarks, headers, footers, comments. Word runs as a visible window. Every change appears instantly. Ctrl+Z works end-to-end.

> Requires a valid Microsoft Word license. Does not include, crack, or distribute Office software.

## What it can do

| Scenario | Description |
|----------|-------------|
| Reports | Agent pulls data, writes, formats, and saves in one flow |
| Contracts | Fill templates with variables, generate formal documents, export PDF |
| Papers | Structured writing, insert tables and charts, format citations |
| CI testing | Auto-create Word docs in pipelines for verification |

## What it works with

| Ecosystem | Agents |
|-----------|--------|
| Claude | Claude Desktop, Claude Code |
| OpenAI | ChatGPT Desktop |
| IDE | Cursor, VS Code / GitHub Copilot, Trae |
| AI Terminal | OpenCode |
| More | Any stdio MCP client |

## Why you need it

- **80 tools** — text, tables, charts, images, headers, footers, comments, bookmarks — precise control over every detail
- **Streaming Markdown** — `start → block → end` in 3 steps. Instant preview. Auto style inheritance.
- **Auto cursor management** — header/footer/table/marker operations reset cursor automatically. No manual position juggling needed.
- **Ctrl+Z friendly** — COM API native operations. Agent messes up? Just undo.
- **Security built-in** — path sanitization, macros disabled, Zod validation, rate limiting, audit logging

---

### ⚡ Quick Start

```bash
git clone https://github.com/HelloWorld-Open/word-mcp-server.git
cd word-mcp-server
npm install && npm run build
```

> Requires MSVC Build Tools for the winax native addon. If it fails, install VS 2022 Build Tools ("Desktop development with C++" workload).

Add to your MCP client config:

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

Then tell your agent: *"Write a weekly report in Word."*

### 📖 More

- [TOOLS.md](./TOOLS.md) — 80-tool quick reference
- [CONTEXT.md](./CONTEXT.md) — architecture & design
- [.env.example](./.env.example) — configuration
- [MIT License](./LICENSE)
