# ADR-0005: Enhanced PositionMap Test Coverage with COM Mock Paths

**Status:** Accepted  
**Date:** 2026-06-15  
**Deciders:** Architecture Team  

## Context

`PositionMap` (480 lines) is the most algorithmically complex module in the `word/` layer (binary search, bilingual heading matching, table-range exclusion, async incremental refresh). However, its test suite completely bypassed the COM interaction path via `(pm as any)` property injection:

- `refresh()` — 0 tests exercising the actual `Range.Find → heading extraction → binary search → table filter` pipeline
- `ensureFresh()` — 0 tests for dirty detection, para count change, content end change, error fallback
- `scheduleRefresh()` — 0 tests for promise management
- `resolveTable()` — only tested empty-tables case
- `resolveBookmark()` — only tested empty-bookmarks case
- `getHeadings()` — untested
- `fetchActualParaCount()` / `paraCountMatches()` — untested
- `refreshContentOnly()` — untested

### Test Quality Before
- 28 tests for 480 lines (~6%)
- COM path: fully bypassed
- Pure logic (resolveHeading, getHeadingContext): well-covered

## Decision

1. **Build rich COM mock** with getter/setter-based `Range.Find` that returns heading positions for specific `Style` values
2. **Add 21 new tests** covering:
   - `refresh()`: heading extraction via Find, table-range exclusion, empty document, no headings
   - `ensureFresh()`: dirty flag, para count change, content end change, COM error fallback
   - `scheduleRefresh()`: initial schedule, deduplication while pending
   - `getHeadings()`: simple getter verification
   - `resolveTable()`: first/second occurrence, offset after table
   - `resolveBookmark()`: cached binary search, not-found, COM fallback path
   - `fetchActualParaCount()` / `paraCountMatches()`: direct COM delegation

## Consequences

### Positive
- 49 tests now cover both pure logic and COM-dependent paths
- Increased confidence in the module's correctness
- Mock pattern establishes template for testing other COM-dependent modules

### Negative
- Test infrastructure requires careful mock setup for `Range.Find` interactions
- Mock fidelity depends on understanding COM `Find.Execute` parameter semantics

### Neutral
- Existing pure-logic tests unchanged
- Mock helper (`makeRefreshableDoc`) reusable for future tests
