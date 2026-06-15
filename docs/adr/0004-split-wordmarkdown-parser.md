# ADR-0004: Split WordMarkdown Parser from Renderer

**Status:** Accepted  
**Date:** 2026-06-15  
**Deciders:** Architecture Team  

## Context

`word-markdown.ts` (661 lines) contained both:
1. **Pure logic** — Markdown parser (`parseBlocks`, `parseInline`, `isLooseTableRow`) with zero COM dependencies
2. **COM rendering** — `writeBlocks`, `insertAtCursor`, `renderBlock`, `typeSeg`, `applyHeadingStyle`, `blockCost`

Mixing pure and impure code made the file hard to test (COM mocking required even for parser tests), hard to reason about, and violated the Single Responsibility Principle.

Tests accessed parser functions through `(md as any).parseBlocks.bind(md)` — bypassing type safety and requiring a full `WordMarkdown` instance (with COM session) to test pure text parsing.

## Decision

1. **Create `src/word/markdown-parser.ts`** (235 lines) with standalone exports:
   - Types: `InlineSegment`, `ListItem`, `Block`
   - Functions: `parseBlocks()`, `parseInline()`, `isLooseTableRow()`
2. **`WordMarkdown`** reduced to 429 lines:
   - Import types and functions from `markdown-parser`
   - Keep COM-dependent methods: `writeBlocks`, `insertAtCursor`, `renderBlock`, `blockCost`, `typeSeg`, `applyHeadingStyle`
3. **Tests updated** to import `parseBlocks`/`parseInline` directly from `markdown-parser` — no `(md as any)` hackery needed

## Consequences

### Positive
- Parser is pure logic — testable without COM mocks
- `WordMarkdown` reduced by ~35% (429 vs 661 lines)
- Tests import directly from parser module — type-safe, no `as any` casts
- Clear separation of concerns: parser (data transformation) vs renderer (COM interaction)

### Negative
- Additional file in the `word/` module
- `Block` type used by both files — import chain from `word-markdown.ts`

### Neutral
- `blockCost()` stayed in `WordMarkdown` (it's used by rendering batching logic)
- Test count unchanged (33 tests for parser, 1 for COM renderer)
