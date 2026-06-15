# ADR-0001: Extract ICursorContext Interface for Cursor Management

**Status:** Accepted  
**Date:** 2026-06-15  
**Deciders:** Architecture Team  

## Context

The `ContextSanitizer` class managed cursor context (main body detection, table/shape/header-footer exit) and text sanitization. `WordBase` hardcoded `cursor = new ContextSanitizer(session)`, coupling all 8 editor subclasses to the concrete implementation. A dead-code file `cursor-position.ts` (112 lines) duplicated core cursor logic.

### Problems
- No abstraction boundary: all editors coupled to `ContextSanitizer` concrete class
- `resetParagraphStyle()` lived in the wrong module (cursor concern mixed with COM interaction)
- `cursor-position.ts` was dead code with zero callers

## Decision

1. **Define `ICursorContext` interface** in `context-sanitizer.ts` with methods: `ensureMainBody()`, `markInBody()`, `markSelectionRead()`, `goToEnd()`, `reset()`
2. **`ContextSanitizer implements ICursorContext`**
3. **`WordBase.cursor` type** changed from `ContextSanitizer` to `ICursorContext`
4. **`WordBase` constructor** accepts optional `cursor?: ICursorContext` for DI
5. **Move `resetParagraphStyle()` + `NORMAL_STYLES`** to `WordFormatting`
6. **Delete `cursor-position.ts`**
7. **`WordTextEditor` / `WordTableEditor`** accept optional `formatting?: WordFormatting` for `resetParagraphStyle()` calls

## Consequences

### Positive
- Editor classes depend on interface, not implementation — swap cursor strategy without source changes
- Dead code eliminated (~112 lines)
- `resetParagraphStyle()` correctly lives in formatting module

### Negative
- Increased constructor parameter count for `WordTextEditor`/`WordTableEditor` (now accept optional second arg)
- `create-server.ts` requires explicit wiring ordering

### Neutral
- Tests can inject mock `ICursorContext` — already supported by `createMockSession()` pattern
