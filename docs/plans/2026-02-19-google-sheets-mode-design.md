# Google Sheets Compatibility Mode — Design Document

**Date:** 2026-02-19
**Status:** Approved
**Goal:** Make HyperFormula behave as close to Google Sheets as possible via a single config flag, targeting drop-in replacement fidelity.

## Overview

Add a `compatibilityMode: 'default' | 'googleSheets'` config option that activates Google Sheets-compatible behavior across config defaults, named expressions, and function implementations. An `'excel'` value is reserved for future use.

Users opt in with one line:

```typescript
const hf = HyperFormula.buildFromArray(data, {
  licenseKey: 'gpl-v3',
  compatibilityMode: 'googleSheets',
  // user-provided values still override the preset
})
```

## Architecture

### 1. Config Option

Add `compatibilityMode` to `ConfigParams` interface and `Config.defaultConfig`:

```typescript
// ConfigParams.ts
compatibilityMode: 'default' | 'googleSheets'

// Config.defaultConfig
compatibilityMode: 'default'
```

### 2. Config Preset

A `googleSheetsDefaults` constant defines GSheets-compatible values for existing config options. In the `Config` constructor, when `compatibilityMode === 'googleSheets'`, these defaults are merged *under* user-provided options (user values always win):

```typescript
const effectiveDefaults = compatibilityMode === 'googleSheets'
  ? { ...Config.defaultConfig, ...googleSheetsDefaults }
  : Config.defaultConfig
```

**GSheets preset values (only options that differ from HF defaults):**

| Option | HF Default | GSheets Preset |
|---|---|---|
| `dateFormats` | `['DD/MM/YYYY', 'DD/MM/YY']` | `['MM/DD/YYYY', 'MM/DD/YY', 'YYYY/MM/DD']` |
| `localeLang` | `'en'` | `'en-US'` |
| `currencySymbol` | `['$']` | `['$', 'USD']` |

All other defaults remain unchanged (GSheets already matches HF defaults for `leapYear1900`, `useArrayArithmetic`, `evaluateNullToZero`, `ignoreWhiteSpace`, separators, etc.).

### 3. Named Expression Auto-Registration

When `compatibilityMode === 'googleSheets'`, `TRUE` and `FALSE` named expressions are automatically registered during engine construction.

**Where:** `BuildEngineFactory.buildEngine()`, at the existing named-expression insertion point (~line 123), before `evaluator.run()`. Injected via `crudOperations.operations.addNamedExpression()` (the internal API, same as user-supplied named expressions).

```typescript
if (config.compatibilityMode === 'googleSheets') {
  crudOperations.operations.addNamedExpression('TRUE', '=TRUE()')
  crudOperations.operations.addNamedExpression('FALSE', '=FALSE()')
}
```

### 4. Function Overrides via GSheets Plugins

Instead of adding conditional branches inside existing plugins, dedicated GSheets override plugins live in a new folder:

```
src/interpreter/plugin/googleSheets/
  GoogleSheetsStatisticalPlugin.ts   # POISSON.DIST, BETA.DIST, BINOM.INV, etc.
  GoogleSheetsTextPlugin.ts          # SPLIT (full GSheets signature)
  GoogleSheetsFinancialPlugin.ts     # NPV, XNPV, RATE, PV, TBILLEQ, TBILLPRICE, etc.
  GoogleSheetsLogicalPlugin.ts       # COUNTA, etc.
  GoogleSheetsMathPlugin.ts          # GCD, LCM, COMBIN, RRI, etc.
  GoogleSheetsDatePlugin.ts          # DAYS, DATEDIF, EDATE, EOMONTH, etc.
  index.ts                           # exports array of all GSheets plugins
```

Each plugin extends `FunctionPlugin` and only declares the functions it overrides in `implementedFunctions`. The rest of each base plugin's functions remain untouched.

**Registration:** In `FunctionRegistry` constructor, when `compatibilityMode === 'googleSheets'`, after loading all default plugins into `instancePlugins`, the GSheets plugins are loaded on top. Since `Map.set` silently overwrites, only the overridden functions are replaced.

```typescript
// FunctionRegistry constructor, after default plugin loading
if (config.compatibilityMode === 'googleSheets') {
  for (const gsPlugin of googleSheetsPlugins) {
    FunctionRegistry.loadPluginFunctions(gsPlugin, this.instancePlugins)
  }
}
```

#### Function Differences to Override

**Category A — Edge-case argument validation (~25 functions):**

Small fixes to boundary conditions, input range validation, and error-vs-value behavior. Each is 1-10 lines in the override plugin method.

| Plugin File | Functions |
|---|---|
| `GoogleSheetsFinancialPlugin` | `TBILLEQ`, `TBILLPRICE`, `NPV`, `XNPV`, `RATE`, `PV`, `DB`, `RRI` |
| `GoogleSheetsStatisticalPlugin` | `POISSON.DIST`, `BETA.DIST`, `BETA.INV`, `BINOM.DIST`, `BINOM.INV`, `NEGBINOM.DIST`, `HYPGEOM.DIST`, `T.INV`, `T.INV.2T`, `T.DIST`, `TDIST`, `WEIBULL.DIST`, `GAMMA`, `NORMSDIST`, `CHISQ.TEST`, `LARGE`, `AVEDEV`, `SKEW`, `DEVSQ` |
| `GoogleSheetsMathPlugin` | `GCD`, `LCM`, `COMBIN` |
| `GoogleSheetsDatePlugin` | `DAYS`, `DATEDIF` |
| `GoogleSheetsLogicalPlugin` | `COUNTA`, `ADDRESS` |
| `GoogleSheetsTextPlugin` | `SPLIT` |

**Category B — SPLIT function redesign:**

HyperFormula's `SPLIT(string, index)` returns a single word. Google Sheets' `SPLIT(text, delimiter, [split_by_each], [remove_empty_text])` returns a spilled array.

`GoogleSheetsTextPlugin` implements the full GSheets signature as an array-returning function. No parser changes needed — the plugin system handles this.

### 5. Boolean Coercion in Aggregation Functions

The `NumericAggregationPlugin.reduce()` method (~line 588 of `NumericAggregationPlugin.ts`) has a 3-way branch for scalars:

- **Cell reference:** applies `coercionFunction` directly (booleans ignored by `strictlyNumbers`)
- **Explicit literal:** coerces boolean to number first, then applies `coercionFunction`

In Google Sheets, referenced booleans in aggregation functions like `HARMEAN`, `GEOMEAN` are also coerced to numbers. This is a shared infrastructure change, gated on `compatibilityMode`:

```typescript
// In the CELL_REFERENCE branch of reduce()
if (this.config.compatibilityMode === 'googleSheets' && typeof value === 'boolean') {
  value = coerceBooleanToNumber(value)
}
```

This affects: `HARMEAN`, `GEOMEAN`, `AVEDEV`, `SKEW`, and other aggregation functions that use `strictlyNumbers` coercion through `reduce()`.

## Bug Fix (Unconditional, Separate Branch)

The `collectDependencies.ts` switch statement is missing a case for `AstNodeType.ARRAY`. Cell references inside inline arrays (`={A1, A2}`) are not tracked as dependencies, so the cell value is never updated when referenced cells change.

**Fix:** Add the missing case to recurse into array elements:

```typescript
case AstNodeType.ARRAY: {
  ast.args.forEach((row) =>
    row.forEach((cellAst) =>
      collectDependenciesFn(cellAst, functionRegistry, dependenciesSet, needArgument)
    )
  )
  return
}
```

This is a genuine bug, not a compatibility difference. Fixed unconditionally in its own branch.

## Branching & PR Strategy

### Branch 1: `fix/inline-array-dependency-tracking` (from `master`)
- Fix `collectDependencies.ts` bug
- Add tests
- PR, merge to `master`

### Branch 2: `feat/google-sheets-compatibility-mode` (from bug fix branch)

Broken into multiple PRs for independent review:

| PR | Scope |
|---|---|
| **A** | Config plumbing: `compatibilityMode` param, preset defaults, named expression auto-registration, future work doc |
| **B** | Boolean coercion infrastructure: `reduce()` change in `NumericAggregationPlugin` |
| **C** | `GoogleSheetsStatisticalPlugin` |
| **D** | `GoogleSheetsFinancialPlugin` |
| **E** | `GoogleSheetsTextPlugin` (includes SPLIT redesign) |
| **F** | `GoogleSheetsMathPlugin` + `GoogleSheetsDatePlugin` + `GoogleSheetsLogicalPlugin` |

Each PR includes tests for the overridden functions, verifying both `'default'` and `'googleSheets'` behavior.

## Limitations & Future Work

See [google-sheets-mode-future-work.md](./2026-02-19-google-sheets-mode-future-work.md) for parser/architectural changes that cannot be addressed via config or plugin overrides.
