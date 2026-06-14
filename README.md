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

Let your AI agent directly control Microsoft Word — create, edit, and format documents in real time, just like a human typing at the keyboard.

Word opens as a visible window. Every change appears instantly. Every edit stays in the undo stack.

> Requires a valid Microsoft Word license. Does not include, crack, or distribute Office software.

---

## 🤖 Compatible AI Agents

Works with any MCP-compatible client. Add the config and your agent can control Word immediately.

| Ecosystem | Agents |
|-----------|--------|
| **Claude** | Claude Desktop, Claude Code (CLI) |
| **IDE** | Cursor, VS Code + Continue / Cline |
| **AI Terminal** | OpenCode, Codex CLI |
| **Others** | Any stdio-based MCP client |

## 💼 Use Cases

| Scenario | What your agent can do |
|----------|------------------------|
| **Reports & memos** | Pull data, write formatted reports, save as `.docx` |
| **Contract drafting** | Fill templates with dynamic content, generate PDF |
| **Academic papers** | Write structured content, insert tables and charts |
| **Batch generation** | Process 100+ documents from JSON/CSV with one instruction |
| **Test automation** | Programmatically create Word docs for CI/QA pipelines |

## ⚡ Quick Start

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

Then ask your agent: _"Write a weekly report in Word"_ or _"Generate a contract from this template"_.

### 🎯 Try this — configure styles first, then write

`word_stream_start` accepts a `baseStyleProfile` parameter to pre-configure any built-in style. The AI sets fonts, sizes, spacing upfront — then writes clean Markdown. Styles are inherited automatically, zero per-paragraph overhead.

Give this to your AI agent:

> **Create a deep learning paper**
>
> First configure styles, then write the full paper:
>
> ```
> word_stream_start title:"Attention-Based Transformer for Image Classification" baseStyleProfile:{
>   "Normal": {"font": {"name": "Times New Roman", "size": 12}, "paragraph": {"lineSpacing": 22, "firstLineIndent": 0.74}},
>   "Heading 1": {"font": {"name": "Arial", "size": 16, "bold": true}, "paragraph": {"spaceBefore": 18, "spaceAfter": 6, "alignment": "center"}},
>   "Heading 2": {"font": {"name": "Arial", "size": 14, "bold": true}, "paragraph": {"spaceBefore": 12, "spaceAfter": 6}},
>   "Heading 3": {"font": {"name": "Arial", "size": 12, "bold": true, "italic": true}, "paragraph": {"spaceBefore": 6, "spaceAfter": 3}}
> }
> ```
>
> Then write:
> - **Title**: "Attention-Based Transformer for Image Classification"
> - **Abstract** (1 paragraph): Summarize that we propose a novel Vision Transformer variant with improved attention mechanism, achieving 92.5% top-1 accuracy on ImageNet with 40% fewer parameters than ViT-Base.
> - **1. Introduction** (2 paragraphs): Briefly motivate the need for efficient vision transformers; mention the quadratic complexity problem of standard self-attention and our proposed sparse attention approach.
> - **2. Method** (2 paragraphs + a table):
>   - Para 1: Describe the overall architecture — patch embedding, transformer encoder with sparse attention, classification head.
>   - Para 2: Explain the sparse attention mechanism — how we reduce O(n²) to O(n√n) using windowed + global token attention.
>   - Insert a table comparing our method with baselines:
>     Model | Parameters | Top-1 Acc | FLOPs
>     ViT-Base | 86M | 81.8% | 17.6G
>     Swin-T | 28M | 83.5% | 4.5G
>     Ours (Tiny) | 12M | 84.2% | 2.1G
>     Ours (Base) | 52M | 92.5% | 11.3G
> - **3. Results** (1 paragraph): Summarize key findings — our method outperforms both ViT and Swin Transformer across all model sizes, with significant efficiency gains.
> - **References**: Inline references like [1], [2] in text
>
> → Save as "Transformer_Image_Classification.docx" on desktop

The AI configures styles once at the start, then writes everything in plain Markdown — formatting (font, size, spacing, bold/italic) inherits automatically.

## ✨ Why Word MCP Server

Every feature is designed to make AI agents better at Word automation.

| Capability | Why it matters for AI |
|-----------|----------------------|
| **108 MCP tools** | Agent can precisely control every aspect — text, formatting, tables, charts, images, bookmarks, headers, footnotes, comments |
| **Real-time visible** | Watch your agent work. Spot mistakes immediately. Interrupt if needed. |
| **Manager API** | Agent completes complex tasks in one call: `set header → page numbers → insert table → bookmark` |
| **Ctrl+Z friendly** | All operations go through COM API — the undo stack works. If the agent makes a mistake, you can undo. |
| **Disconnect-safe** | MCP disconnect keeps Word open. No lost work, no surprises. |
| **Auto backup** | `.bak` file created before every save. Safe to experiment. |
| **Chart data isolation** | Separate child process + 15s timeout — prevents Excel COM from freezing Word |
| **Security built-in** | Path traversal protection, macros forcibly disabled, Zod input validation, rate limiting, audit logging, error sanitization |

## 🛠️ Usage

### Stream formatted content with Markdown (recommended)

```
word_stream_start title:"My Report"
word_stream_block text:"# Title\n\nContent here..."
word_stream_end
```

The streaming API creates a document, writes Markdown in real-time blocks, and saves. Supports full Markdown: headings, bold, italic, code, tables, lists, links, strikethrough, blockquotes, horizontal rules, and code blocks.

Use `baseStyleProfile` to pre-configure font and paragraph formatting for styles like `Normal`, `Heading 1`:

```
word_stream_start title:"Paper" baseStyleProfile:{"Normal":{"font":{"name":"SimSun","size":12},"paragraph":{"firstLineIndent":0.74}}}
```

### High-level Manager API

```
word_mgr_set_header text:"Report" alignment:"center"
word_mgr_set_page_numbers target:"footer"
word_mgr_insert_table rows:5 cols:3 data:[["A","B","C"],["1","2","3"]]
word_mgr_add_bookmark name:"section1"
```

Manager API handles cursor positioning, paragraph separation, and screen updating automatically.

## 📝 Built-in Prompts

Three MCP prompt templates help agents follow best-practice workflows:

| Prompt | Description |
|--------|-------------|
| `create_report` | Generate a step-by-step plan for creating a structured Word report (title, sections, style) |
| `format_document` | Get a guided workflow for formatting an existing document |
| `state_machine` | Explains the 4-state + 2-substate model and correct operation ordering |

## 📖 Tool Reference

See **[TOOLS.md](./TOOLS.md)** — 108 tools across 12 modules: document lifecycle, content editing, formatting, tables, charts, images, text boxes, structure, clipboard, Manager API, semantic navigation, and variable replacement.

## 🏗️ Architecture

### Communication Flow

```
MCP Client (Claude Desktop, Cursor, OpenCode, etc.)
    │  JSON-RPC over stdio
    ▼
build/parent.js  (Watchdog — 30s timeout, auto-restart)
    │  spawn + pipe stdin/stdout
    ▼
build/child.js   (McpServer)
    │  winax COM Automation (native)
    ├─ Main thread: 108 tools
    └─ Child process: chart data setting (fork + 15s timeout)
            │
            ▼
        WINWORD.EXE  (visible window)
```

### Code Layout

```
src/
├── index.ts            # Entry point
├── parent.ts           # Watchdog process (30s timeout, auto-restart)
├── child.ts            # MCP server process
├── server/             # MCP protocol layer
│   ├── create-server.ts
│   ├── server-context.ts      # Shared session dependencies
│   ├── session-director.ts    # Session orchestration (streaming lock, edit mode)
│   ├── tools/                 # 12 tool modules (108 tools)
│   │   ├── content.ts
│   │   ├── document.ts
│   │   ├── formatting.ts
│   │   ├── helper.ts
│   │   ├── manager.ts         # High-level Manager API
│   │   ├── media.ts
│   │   ├── reader.ts
│   │   ├── semantic.ts
│   │   ├── stream.ts          # Streaming document writer (recommended)
│   │   ├── structure.ts
│   │   ├── tables.ts
│   │   └── variable.ts
│   └── prompts/               # Built-in prompt templates (3 prompts)
│       ├── report-prompts.ts
│       └── state-machine.ts
├── word/              # Word COM automation core (19 modules)
│   ├── session.ts
│   ├── application.ts
│   ├── document.ts
│   ├── document-registry.ts
│   ├── word-base.ts
│   ├── word-text-editor.ts
│   ├── word-markdown.ts
│   ├── word-stream-writer.ts
│   ├── word-table-editor.ts
│   ├── word-media-editor.ts
│   ├── word-document-structure.ts
│   ├── formatting.ts
│   ├── cursor-position.ts
│   ├── position-map.ts
│   ├── variable-replacer.ts
│   ├── chart-data-bridge.ts
│   ├── chart-data-worker.ts
│   ├── process-monitor.ts
│   └── types.ts
└── security/          # 5-layer defense
    ├── path-sanitizer.ts
    ├── policy.ts
    ├── rate-limiter.ts
    ├── audit.ts
    └── errors.ts
```

## 🔒 Security (5-layer defense)

- Path traversal detection (7 attack vectors: `..`, ADS, network paths, cross-drive, system dirs, device paths, whitelist)
- Macros forcibly disabled via `AutomationSecurity = 3`
- Input validation via Zod (all tool parameters at runtime)
- Rate limiting (sliding window, configurable via `RATE_LIMIT_WINDOW_MS`/`RATE_LIMIT_MAX_CALLS`)
- Audit logging — every tool call logged with timestamp, duration, success/failure, redacted args
- Error message sanitization (no internal paths/COM details leaked)

## 🧪 Tests

**106 tests, all passing** ✅ — 9 unit + 1 integration.

Unit tests mock the winax layer (no real Word needed). Coverage: markdown parser (33), position map (21), session (13), text editor (6), security (30).

```bash
npm test
```

## ⚙️ Configuration

Configure via `.env` file in the project root. See `.env.example` for all options:

| Variable | Default | Description |
|----------|---------|-------------|
| `ALLOWED_DIRECTORIES` | *(unrestricted)* | Semicolon-separated directory whitelist |
| `ALLOW_NETWORK_PATHS` | `false` | Allow UNC network paths (`\\server\share`) |
| `MAX_FILE_SIZE` | `52428800` | Max file size in bytes |
| `MAX_TEXT_LENGTH` | `1000000` | Max text input length in chars |
| `OPERATION_TIMEOUT_MS` | `30000` | Tool operation timeout in ms |
| `RATE_LIMIT_WINDOW_MS` | `5000` | Rate limiter sliding window in ms |
| `RATE_LIMIT_MAX_CALLS` | `30` | Max calls per window |
| `WATCHDOG_TIMEOUT_MS` | `30000` | Child process silent timeout in ms |
| `WATCHDOG_INTERVAL_MS` | `5000` | Watchdog health check interval in ms |
| `CHART_OP_TIMEOUT` | `15000` | Chart data operation timeout in ms |
| `CHART_WORKER_IDLE_TIMEOUT` | `60000` | Chart worker idle timeout in ms |

## 📄 License

[MIT](./LICENSE)
