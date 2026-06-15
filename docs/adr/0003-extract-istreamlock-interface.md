# ADR-0003: Extract IStreamLock Interface to Resolve Layer Violation

**Status:** Accepted  
**Date:** 2026-06-15  
**Deciders:** Architecture Team  

## Context

`StreamingMarkdownWriter` in `src/word/` imported `SessionDirector` from `src/server/`. This created a **layer violation**: the `word/` module depended on the `server/` module, preventing independent reuse of the word layer.

The `SessionDirector` provides a stream-lock mechanism (3 methods) used by `StreamingMarkdownWriter`:
- `acquireStreamLock(toolName)`
- `releaseStreamLock()`
- `refreshWatchdog()`

### Architecture Dependency
```
server/ ──> word/  (correct)
word/   ──> server/  (wrong — layer violation)
```

## Decision

1. **Define `IStreamLock` interface** in `src/word/types.ts` with the 3 methods
2. **`SessionDirector implements IStreamLock`** — no code changes to `SessionDirector`
3. **`StreamingMarkdownWriter.director` type** changed from `SessionDirector` to `IStreamLock`
4. **Remove** `import type { SessionDirector } from "../server/session-director.js"` from `word-stream-writer.ts`
5. **Wiring in `create-server.ts` unchanged** — `SessionDirector` instance satisfies the interface

## Consequences

### Positive
- Layer violation eliminated: `word/` depends only on interfaces defined in `word/`
- `word/` module is now independently reusable
- `SessionDirector` unchanged — zero regression risk
- Interface is minimal (3 methods), easy to maintain

### Negative
- Minor indirection: interface defined in one file, implemented in another
- Two files to update if stream-lock API changes

### Neutral
- New `import type { IStreamLock }` in `session-director.ts` creates a `server/` → `word/types.ts` dependency (correct direction)
