> [English](./README.md) | [中文](./README.zh-CN.md)

# Word MCP Server

[![CI](https://github.com/HelloWorld-Open/word-mcp-server/actions/workflows/ci.yml/badge.svg)](https://github.com/HelloWorld-Open/word-mcp-server/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Node](https://img.shields.io/badge/node-%3E%3D20-brightgreen)](https://nodejs.org)
[![Platform](https://img.shields.io/badge/platform-Windows-blue)]()

MCP server for real-time Microsoft Word document manipulation via COM automation. Word runs in a visible window — every change appears instantly.

> Requires a valid Microsoft Word license. Does not include, crack, or distribute Office software.

---

## Features

- **110+ MCP tools** — document lifecycle, content, formatting, tables, charts, media, structure, clipboard, Manager API, semantic navigation, variable replacement
- **Real-time visible** — Word opens in foreground, all changes immediately visible
- **Ctrl+Z friendly** — everything uses COM API, undo stack works normally
- **Disconnect-safe** — MCP disconnect keeps Word open, no lost edits
- **Auto backup** — `.bak` created before every save
- **Chart data isolation** — child process + 15s timeout prevents Excel COM freeze

## Quick Start

```bash
git clone https://github.com/HelloWorld-Open/word-mcp-server.git
cd word-mcp-server
npm install && npm run build
```

Add to your MCP client config:

```json
{
  "mcpServers": {
    "word": {
      "command": "node",
      "args": ["C:\\path\\to\\word-mcp-server\\build\\parent.js"]
    }
  }
}
```

> `npm install` requires MSVC Build Tools to compile the winax native addon. If it fails, install Visual Studio 2022 Build Tools with the "Desktop development with C++" workload.

## Usage

### Write formatted content with Markdown

```md
# Section Title
Some **bold** and *italic* text with `inline code`, ~~strikethrough~~, and [a link](https://example.com).

- Top level item
  - Nested item with **bold**
- Another top item

1. First step
2. Second step

```
code block with pre-formatted text
```

| Name | Value |
|------|-------|
| Alpha | 100 |
| Beta | 200 |

> Notable quote here
---
```

### Batch mode for speed

```js
word_batch_start   // 3-10x faster
// ... multiple operations ...
word_batch_end     // refresh window
```

### High-level Manager API

```js
word_mgr_create_document    // auto close + page setup
word_mgr_write_content      // auto cursor + ScreenUpdating
word_mgr_insert_table       // auto paragraph separator
word_mgr_add_bookmark       // auto deep cursor reset
word_mgr_save               // save + optional PDF
```

## Tool Reference

See **[TOOLS.md](./TOOLS.md)** — 110+ tools across 12 modules.

## Architecture

### Communication Flow

```
MCP Client (Claude Desktop, etc.)
    │  JSON-RPC over stdio
    ▼
build/parent.js  (Watchdog — 30s timeout, auto-restart)
    │  spawn + pipe stdin/stdout
    ▼
build/child.js   (McpServer)
    │  winax COM Automation (native)
    ├─ Main thread: 110+ tools
    └─ Child process: chart data setting (fork + 15s timeout)
            │
            ▼
        WINWORD.EXE  (visible window)
```

### Code Layout

```
src/
├── index.ts       # Entry point
├── parent.ts      # Watchdog process
├── child.ts       # MCP server
├── manager/       # High-level API (document builder pattern)
├── word/          # COM adapter layer via winax
├── security/      # 4-layer defense (path, macro, validation, redaction)
└── server/        # MCP protocol layer (110+ tools, prompts)
```

## Security

- Path traversal detection (7 attack vectors: `..`, ADS, network paths, cross-drive, system dirs, device paths, whitelist)
- Macros forcibly disabled via `AutomationSecurity = 3`
- Input validation via Zod (all tool parameters at runtime)
- Rate limiting (sliding window, configurable)
- Error message sanitization (no internal paths/COM details leaked)

## Tests

**116 tests, all passing** ✅ — 9 unit + 1 integration.

Unit tests mock the winax layer (no real Word needed). Coverage: markdown parser (33), position map (21), text editor (16), session (13), security (30).

```bash
npm test
```

## License

[MIT](./LICENSE)
