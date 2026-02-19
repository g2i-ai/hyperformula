# Google Sheets Compatibility Mode — Future Work

**Date:** 2026-02-19
**Related:** [Design Document](./2026-02-19-google-sheets-mode-design.md)

These are known behavioral differences between HyperFormula and Google Sheets that require parser-level or architectural changes. They are documented here for future evaluation.

## 1. Dependency Collection at Evaluation Time

**Current behavior:** HyperFormula collects dependencies statically from the AST during parsing. All cell references in a formula are registered as dependencies regardless of whether the code path is reachable at runtime.

**Google Sheets behavior:** Dependencies are collected during evaluation. `=IF(FALSE(), A1, 0)` does not create a dependency on `A1`, so self-referencing formulas guarded by conditionals do not produce `#CYCLE` errors.

**Example:**
```
A1: =IF(FALSE(), A1, 0)
```
- HyperFormula: `#CYCLE` error (A1 depends on A1 statically)
- Google Sheets: `0` (A1 is never evaluated in the FALSE branch)

**Why it's hard:** The entire `DependencyGraph` + `TopSort` (Tarjan SCC) architecture assumes static dependencies known at parse time. Switching to evaluation-time collection would require:
- A fundamentally different graph update strategy (dependencies discovered during evaluation)
- Handling cycles that only appear at runtime
- Likely significant performance implications for incremental recalculation

**Impact:** Rare in practice — only affects self-referencing formulas guarded by conditional branches.

## 2. OFFSET as a Runtime Function

**Current behavior:** `OFFSET` is resolved at parse time by `handleOffsetHeuristic` in `FormulaParser.ts`. It's a parser-level "macro" that transforms into a `CELL_REFERENCE` or `CELL_RANGE` AST node. The first argument must be a single cell reference. `OFFSET` does not exist as a runtime function (`['OFFSET', undefined]` in `FunctionRegistry`).

**Google Sheets behavior:** `OFFSET` is a runtime function that accepts a range as the first argument.

**Example:**
```
=OFFSET(A1:B1, 0, 0)
```
- HyperFormula: Parse error (`StaticOffsetError`)
- Google Sheets: Returns the range A1:B1

**Why it's hard:** Requires:
1. Implementing OFFSET as a real runtime function with `doesNotNeedArgumentsToBeComputed: true`
2. Special dependency collection (the resulting reference depends on runtime arguments)
3. Removing or relaxing `handleOffsetHeuristic` in the parser
4. Careful interaction with the lazy AST transformation system

**Impact:** Users must pass single cell references to OFFSET. Range-as-first-arg is a moderately common pattern in GSheets.

## 3. Named Expression Naming Rules

**Current behavior:** HyperFormula rejects named expression names that could be interpreted as cell references (case-insensitive). For example, `ProductPrice1` is invalid because it matches the pattern of a cell reference.

**Google Sheets / Excel behavior:** Names that look like cell references are allowed if the column part is at least 4 letters long. `ProductPrice1` is valid.

**Example:**
```javascript
hf.addNamedExpression('ProductPrice1', '=42')
```
- HyperFormula: Error (name looks like a cell reference)
- Google Sheets: Works fine

**Why it's hard:** The validation logic is in the parser/named-expression layer. The cell reference detection regex would need to be relaxed, and the parser would need to disambiguate between names and cell references based on column-length heuristics.

**Impact:** Some named expressions valid in GSheets cannot be created in HyperFormula.

## 4. Ranges Created with `:` (Mixed Range Types)

**Current behavior:** Ranges must consist of two addresses (`A1:B5`), two columns (`A:C`), or two rows (`3:5`). They cannot be mixed (e.g., `A1:C` or `1:A2`) or contain named expressions.

**Google Sheets behavior:** All range combinations are allowed.

**Why it's hard:** Parser grammar change. The `FormulaParser` range production rules would need to accept heterogeneous range endpoints.

**Impact:** Users must use uniform range types. Mixed ranges are uncommon but valid in GSheets.

## 5. TEXT Function Formatting

**Current behavior:** Limited formatting options. Only some date format tokens (`hh`, `mm`, `ss`, `am`, `pm`, `a`, `p`, `dd`, `yy`, `yyyy`). No currency formatting inside TEXT.

**Google Sheets behavior:** Wide variety of formatting options.

**Why it's hard:** Implementing a full format string parser for `TEXT()` is a significant feature in itself — number formatting, currency, custom patterns, etc.

**Impact:** Users relying on advanced `TEXT()` formatting will get incomplete results.

## 6. Cell References Inside Inline Arrays (Partial)

**Note:** The static dependency tracking bug for `={A1, A2}` is fixed unconditionally as part of this effort (see design doc). However, there may be remaining edge cases around how inline array cell references interact with the dependency graph for structural changes (row/column insertion/deletion). These need further testing.
