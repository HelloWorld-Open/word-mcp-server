# ADR-0002: Unified Mock Factory for Test Infrastructure

**Status:** Accepted  
**Date:** 2026-06-15  
**Deciders:** Architecture Team  

## Context

Three test files (`text-editor.test.ts`, `position-map.test.ts`, `markdown.test.ts`) each defined incompatible `createMockSession()` factories. A fourth existed in `session.test.ts` as `createMockWinax()`. If `IWordSession` interface changed, all four would need synchronous updates — a maintenance burden that frequently causes test rot.

### Problems
- 4 incompatible mock implementations of the same interface
- No shared type-checking across mocks
- `IWordSession` changes require fixing 4+ locations
- Some mocks were incomplete (missing newer interface methods)

## Decision

1. **Create `tests/unit/test-helpers.ts`** with:
   - `createMockSession()` — returns full `IWordSession` with `vi.fn()` stubs
   - `createMockDoc()` — returns standard mock Word document object
   - `createMockSel()` — returns standard mock Selection object
   - All accept optional `overrides` / `customDoc` / `customSel` for per-test customization
2. **Update 3 test files** to import from shared factory
3. **Delete 3 incompatible** `createMockSession()` implementations
4. **Session test** retains its own `createMockWinax()` (tests winax-specific path, not `IWordSession`)

## Consequences

### Positive
- Single source of truth for `IWordSession` mock shape
- Adding a method to `IWordSession` requires updating one factory
- Tests more readable with standard `createMockSession()` call
- Overrides pattern enables test-specific customization without copy-paste

### Negative
- None

### Neutral
- Integration tests (`smoke.test.ts`) not affected — they test the actual build
