# Google Sheets Compatibility Mode — Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Add a `compatibilityMode: 'googleSheets'` config option that activates Google Sheets-compatible behavior via config presets, auto-registered named expressions, and dedicated override function plugins.

**Architecture:** A single config flag (`compatibilityMode`) controls three layers: (1) config default overrides for date/locale/currency, (2) automatic TRUE/FALSE named expression registration at engine construction, and (3) GSheets-specific function plugins that silently override built-in functions via the existing `FunctionRegistry` map. The bug fix for inline array dependency tracking ships as a separate prerequisite branch.

**Tech Stack:** TypeScript, Jest, Chevrotain (parser — read-only for this effort)

**Design doc:** `docs/plans/2026-02-19-google-sheets-mode-design.md`

---

## Branch 1: `fix/inline-array-dependency-tracking`

Base: `master`

### Task 1: Fix missing ARRAY case in collectDependencies

**Files:**
- Modify: `src/parser/collectDependencies.ts:18-86` (add ARRAY case to switch)
- Test: `test/inline-array-dependencies.spec.ts` (create)

**Step 1: Create the branch**

```bash
git checkout master
git checkout -b fix/inline-array-dependency-tracking
```

**Step 2: Write the failing test**

Create `test/inline-array-dependencies.spec.ts`:

```typescript
import {HyperFormula} from '../src'
import {adr} from './testUtils'

describe('Inline array dependency tracking', () => {
  it('should update cell when inline array references change', () => {
    const hf = HyperFormula.buildFromArray([
      [1, 2],
      ['={A1, B1}', null],
    ], {licenseKey: 'gpl-v3'})

    // Initial value should be computed
    expect(hf.getCellValue(adr('A2'))).toBe(1)

    // Change a referenced cell
    hf.setCellContents(adr('A1'), 10)

    // The inline array cell should update
    expect(hf.getCellValue(adr('A2'))).toBe(10)

    hf.destroy()
  })

  it('should update cell when nested inline array references change', () => {
    const hf = HyperFormula.buildFromArray([
      [1, 2, 3],
      ['={A1, B1; C1, A1}', null],
    ], {licenseKey: 'gpl-v3'})

    expect(hf.getCellValue(adr('A2'))).toBe(1)

    hf.setCellContents(adr('C1'), 99)

    // C1 is in the second row of the array — but A2 is the top-left cell,
    // and we're checking that the dependency was tracked at all
    // The array spills: A2=1, B2=2, A3=99, B3=1
    // After change: A2=1 still (A1 didn't change), but the dependency graph
    // should have re-evaluated this cell
    expect(hf.getCellValue(adr('A2'))).toBe(1)

    // Change A1 to verify the dependency exists
    hf.setCellContents(adr('A1'), 50)
    expect(hf.getCellValue(adr('A2'))).toBe(50)

    hf.destroy()
  })

  it('should track named expression dependencies in inline arrays', () => {
    const hf = HyperFormula.buildFromArray([
      ['={A2, A3}', null],
      [10, null],
      [20, null],
    ], {licenseKey: 'gpl-v3'})

    expect(hf.getCellValue(adr('A1'))).toBe(10)

    hf.setCellContents(adr('A2'), 42)
    expect(hf.getCellValue(adr('A1'))).toBe(42)

    hf.setCellContents(adr('A3'), 99)
    // A3 is the second element — spills to B1
    expect(hf.getCellValue(adr('B1'))).toBe(99)

    hf.destroy()
  })
})
```

**Step 3: Run test to verify it fails**

Run: `cross-env NODE_ICU_DATA=node_modules/full-icu npx jest test/inline-array-dependencies.spec.ts -v`

Expected: At least the first test should FAIL — after `setCellContents(adr('A1'), 10)`, `getCellValue(adr('A2'))` will still return `1` because the dependency on A1 was never registered.

**Step 4: Implement the fix**

Modify `src/parser/collectDependencies.ts`. Add the ARRAY case before the closing `}` of the switch (after the FUNCTION_CALL case, line 85):

```typescript
    case AstNodeType.ARRAY: {
      ast.args.forEach((row: Ast[]) =>
        row.forEach((cellAst: Ast) =>
          collectDependenciesFn(cellAst, functionRegistry, dependenciesSet, true)
        )
      )
      return
    }
```

The `needArgument` is set to `true` because all elements of an inline array are real dependencies (they're always evaluated).

**Step 5: Run test to verify it passes**

Run: `cross-env NODE_ICU_DATA=node_modules/full-icu npx jest test/inline-array-dependencies.spec.ts -v`

Expected: All tests PASS.

**Step 6: Run full test suite**

Run: `npm run test:jest`

Expected: No regressions.

**Step 7: Commit**

```bash
git add src/parser/collectDependencies.ts test/inline-array-dependencies.spec.ts
git commit -m "fix: track cell dependencies inside inline array literals

The collectDependencies switch statement was missing a case for
AstNodeType.ARRAY, causing cell references inside inline arrays
(e.g., ={A1, B1}) to not be registered as dependencies. The value
was computed correctly on initial evaluation but never updated when
referenced cells changed.

Add the missing ARRAY case to recurse into all array elements and
collect their dependencies."
```

---

## Branch 2: `feat/google-sheets-compatibility-mode`

Base: `fix/inline-array-dependency-tracking`

```bash
git checkout fix/inline-array-dependency-tracking
git checkout -b feat/google-sheets-compatibility-mode
```

### Task 2: Add `compatibilityMode` config option

**Files:**
- Modify: `src/ConfigParams.ts:10` (add to interface)
- Modify: `src/Config.ts:31-70` (add to defaultConfig + constructor)
- Test: `test/compatibility-mode.spec.ts` (create)

**Step 1: Write the failing test**

Create `test/compatibility-mode.spec.ts`:

```typescript
import {HyperFormula} from '../src'
import {Config} from '../src/Config'

describe('compatibilityMode config option', () => {
  it('should default to "default"', () => {
    expect(Config.defaultConfig.compatibilityMode).toBe('default')
  })

  it('should accept "googleSheets" value', () => {
    const config = new Config({licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'})
    expect(config.compatibilityMode).toBe('googleSheets')
  })

  it('should accept "default" value', () => {
    const config = new Config({licenseKey: 'gpl-v3', compatibilityMode: 'default'})
    expect(config.compatibilityMode).toBe('default')
  })

  it('should be accessible on the engine config', () => {
    const hf = HyperFormula.buildEmpty({licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'})
    expect(hf.getConfig().compatibilityMode).toBe('googleSheets')
    hf.destroy()
  })
})
```

**Step 2: Run test to verify it fails**

Run: `cross-env NODE_ICU_DATA=node_modules/full-icu npx jest test/compatibility-mode.spec.ts -v`

Expected: FAIL — `compatibilityMode` doesn't exist on ConfigParams.

**Step 3: Implement ConfigParams addition**

Add to `src/ConfigParams.ts` interface (before the closing `}`):

```typescript
  /**
   * Sets the compatibility mode for the engine.
   *
   * When set to `'googleSheets'`, applies Google Sheets-compatible defaults
   * for date formats, locale, currency, and function behavior.
   *
   * User-provided config values always override the preset defaults.
   *
   * For more information, see the [Compatibility with Google Sheets guide](/guide/compatibility-with-google-sheets.md).
   * @default 'default'
   * @category Engine
   */
  compatibilityMode: 'default' | 'googleSheets',
```

**Step 4: Implement Config changes**

In `src/Config.ts`:

1. Add to `Config.defaultConfig` (alphabetical order, after `chooseAddressMappingPolicy`):
   ```typescript
   compatibilityMode: 'default' as const,
   ```

2. Add the readonly property declaration (after `chooseAddressMappingPolicy`):
   ```typescript
   /** @inheritDoc */
   public readonly compatibilityMode: 'default' | 'googleSheets'
   ```

3. Define the GSheets preset constant (above the `Config` class):
   ```typescript
   const googleSheetsDefaults: Partial<ConfigParams> = {
     dateFormats: ['MM/DD/YYYY', 'MM/DD/YY', 'YYYY/MM/DD'],
     localeLang: 'en-US',
     currencySymbol: ['$', 'USD'],
   }
   ```

4. In the constructor, after destructuring `options` and before existing field assignments:
   ```typescript
   this.compatibilityMode = configValueFromParam(
     compatibilityMode, ['default', 'googleSheets'], 'compatibilityMode'
   )
   ```

5. Modify how defaults are resolved. The key change is that GSheets preset values should apply when the user hasn't explicitly provided them. In the constructor, before each affected field's assignment, check if the value was explicitly provided or should fall back to the GSheets preset.

   The cleanest approach: create a merged defaults object at the top of the constructor, then update `configValueFromParam` calls for affected fields to use the merged defaults. Since `configValueFromParam` already falls back to `Config.defaultConfig`, we need to temporarily swap `Config.defaultConfig` or use a different approach.

   **Simpler approach:** After setting `this.compatibilityMode`, explicitly override the fields if (a) the user didn't provide them AND (b) we're in googleSheets mode. Do this by reading the raw `options` to check what was provided:

   After all existing field assignments, add:
   ```typescript
   if (this.compatibilityMode === 'googleSheets') {
     if (dateFormats === undefined) {
       this.dateFormats = [...googleSheetsDefaults.dateFormats!]
     }
     if (localeLang === undefined) {
       this.localeLang = googleSheetsDefaults.localeLang!
     }
     if (currencySymbol === undefined) {
       this.currencySymbol = [...googleSheetsDefaults.currencySymbol!]
     }
   }
   ```

   Note: The `readonly` fields can be reassigned in the constructor. The `readonly` modifier only prevents reassignment outside the constructor.

6. Add `compatibilityMode` to the destructured variables from `options` in the constructor.

**Step 5: Run test to verify it passes**

Run: `cross-env NODE_ICU_DATA=node_modules/full-icu npx jest test/compatibility-mode.spec.ts -v`

Expected: All tests PASS.

**Step 6: Commit**

```bash
git add src/ConfigParams.ts src/Config.ts test/compatibility-mode.spec.ts
git commit -m "feat: add compatibilityMode config option

Add 'compatibilityMode' to ConfigParams accepting 'default' or
'googleSheets'. When set to 'googleSheets', overrides dateFormats,
localeLang, and currencySymbol defaults to match Google Sheets
en-US locale. User-provided values always take precedence."
```

---

### Task 3: Add GSheets config preset tests

**Files:**
- Modify: `test/compatibility-mode.spec.ts`

**Step 1: Add preset override tests**

Append to `test/compatibility-mode.spec.ts`:

```typescript
describe('Google Sheets config preset', () => {
  it('should use GSheets date formats when not explicitly provided', () => {
    const config = new Config({licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'})
    expect(config.dateFormats).toEqual(['MM/DD/YYYY', 'MM/DD/YY', 'YYYY/MM/DD'])
  })

  it('should use GSheets locale when not explicitly provided', () => {
    const config = new Config({licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'})
    expect(config.localeLang).toBe('en-US')
  })

  it('should use GSheets currency symbols when not explicitly provided', () => {
    const config = new Config({licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'})
    expect(config.currencySymbol).toEqual(['$', 'USD'])
  })

  it('should allow user to override GSheets preset values', () => {
    const config = new Config({
      licenseKey: 'gpl-v3',
      compatibilityMode: 'googleSheets',
      dateFormats: ['YYYY-MM-DD'],
      localeLang: 'de-DE',
    })
    expect(config.dateFormats).toEqual(['YYYY-MM-DD'])
    expect(config.localeLang).toBe('de-DE')
    // Currency should still use GSheets default since not overridden
    expect(config.currencySymbol).toEqual(['$', 'USD'])
  })

  it('should use standard defaults when compatibilityMode is "default"', () => {
    const config = new Config({licenseKey: 'gpl-v3', compatibilityMode: 'default'})
    expect(config.dateFormats).toEqual(['DD/MM/YYYY', 'DD/MM/YY'])
    expect(config.localeLang).toBe('en')
    expect(config.currencySymbol).toEqual(['$'])
  })
})
```

**Step 2: Run tests**

Run: `cross-env NODE_ICU_DATA=node_modules/full-icu npx jest test/compatibility-mode.spec.ts -v`

Expected: All PASS.

**Step 3: Commit**

```bash
git add test/compatibility-mode.spec.ts
git commit -m "test: add Google Sheets config preset tests"
```

---

### Task 4: Auto-register TRUE/FALSE named expressions

**Files:**
- Modify: `src/BuildEngineFactory.ts:122-127`
- Modify: `test/compatibility-mode.spec.ts`

**Step 1: Write the failing test**

Append to `test/compatibility-mode.spec.ts`:

```typescript
import {adr} from './testUtils'

describe('Google Sheets named expression auto-registration', () => {
  it('should auto-register TRUE named expression in googleSheets mode', () => {
    const hf = HyperFormula.buildFromArray([
      ['=TRUE'],
    ], {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'})

    expect(hf.getCellValue(adr('A1'))).toBe(true)
    hf.destroy()
  })

  it('should auto-register FALSE named expression in googleSheets mode', () => {
    const hf = HyperFormula.buildFromArray([
      ['=FALSE'],
    ], {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'})

    expect(hf.getCellValue(adr('A1'))).toBe(false)
    hf.destroy()
  })

  it('should NOT auto-register TRUE/FALSE in default mode', () => {
    const hf = HyperFormula.buildFromArray([
      ['=TRUE'],
    ], {licenseKey: 'gpl-v3', compatibilityMode: 'default'})

    // In default mode, TRUE without () is a named expression reference,
    // which doesn't exist — should be a NAME error
    const val = hf.getCellValue(adr('A1'))
    expect(val).toBeInstanceOf(Object) // DetailedCellError
    hf.destroy()
  })

  it('should not conflict with user-provided named expressions', () => {
    const hf = HyperFormula.buildFromArray([
      ['=MyExpr'],
    ], {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'},
    [{name: 'MyExpr', expression: '=42'}])

    expect(hf.getCellValue(adr('A1'))).toBe(42)
    hf.destroy()
  })
})
```

Note: The import for `adr` needs to be added at the top of the file. Also import `Config` if not already there. The test file header should look like:

```typescript
import {HyperFormula} from '../src'
import {Config} from '../src/Config'
import {adr} from './testUtils'
```

**Step 2: Run test to verify it fails**

Run: `cross-env NODE_ICU_DATA=node_modules/full-icu npx jest test/compatibility-mode.spec.ts -v`

Expected: `=TRUE` and `=FALSE` tests FAIL (NAME error in googleSheets mode).

**Step 3: Implement named expression auto-registration**

In `src/BuildEngineFactory.ts`, after the existing named expression loop (line 126) and before `const evaluator` (line 128), add:

```typescript
    if (config.compatibilityMode === 'googleSheets') {
      crudOperations.operations.addNamedExpression('TRUE', '=TRUE()')
      crudOperations.operations.addNamedExpression('FALSE', '=FALSE()')
    }
```

Also add `ConfigParams` is already imported (line 31). No new imports needed — `config.compatibilityMode` is already available.

**Step 4: Run test to verify it passes**

Run: `cross-env NODE_ICU_DATA=node_modules/full-icu npx jest test/compatibility-mode.spec.ts -v`

Expected: All PASS.

**Step 5: Run full test suite**

Run: `npm run test:jest`

Expected: No regressions.

**Step 6: Commit**

```bash
git add src/BuildEngineFactory.ts test/compatibility-mode.spec.ts
git commit -m "feat: auto-register TRUE/FALSE named expressions in googleSheets mode

When compatibilityMode is 'googleSheets', automatically register
TRUE and FALSE as named expressions during engine construction,
matching Google Sheets' built-in boolean constants."
```

---

### Task 5: GSheets plugin infrastructure — create folder and index

**Files:**
- Create: `src/interpreter/plugin/googleSheets/index.ts`

**Step 1: Create the directory and index file**

Create `src/interpreter/plugin/googleSheets/index.ts`:

```typescript
/**
 * @license
 * Copyright (c) 2025 Handsoncode. All rights reserved.
 */

import {FunctionPluginDefinition} from '../FunctionPlugin'

/**
 * All Google Sheets compatibility plugins.
 * These override specific built-in functions with Google Sheets-compatible behavior.
 */
export const googleSheetsPlugins: FunctionPluginDefinition[] = [
  // Plugins will be added here as they are implemented:
  // GoogleSheetsTextPlugin,
  // GoogleSheetsStatisticalPlugin,
  // GoogleSheetsFinancialPlugin,
  // GoogleSheetsMathPlugin,
  // GoogleSheetsDatePlugin,
  // GoogleSheetsLogicalPlugin,
]
```

**Step 2: Commit**

```bash
git add src/interpreter/plugin/googleSheets/index.ts
git commit -m "feat: create googleSheets plugin folder with index

Scaffold the directory structure for Google Sheets compatibility
override plugins. Each plugin will override specific built-in
functions with GSheets-compatible behavior."
```

---

### Task 6: Wire GSheets plugins into FunctionRegistry

**Files:**
- Modify: `src/interpreter/FunctionRegistry.ts:55-67` (constructor)
- Test: `test/compatibility-mode.spec.ts`

**Step 1: Write the failing test**

Append to `test/compatibility-mode.spec.ts`:

```typescript
describe('Google Sheets plugin registration', () => {
  it('should load GSheets plugins when compatibilityMode is googleSheets', () => {
    // This test validates the wiring — specific function overrides
    // are tested in each plugin's own test file.
    // For now, just verify the engine builds successfully with the mode.
    const hf = HyperFormula.buildFromArray([
      [1, 2, '=SUM(A1,B1)'],
    ], {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'})

    expect(hf.getCellValue(adr('C1'))).toBe(3)
    hf.destroy()
  })
})
```

**Step 2: Implement FunctionRegistry wiring**

In `src/interpreter/FunctionRegistry.ts`:

1. Add import at top:
   ```typescript
   import {googleSheetsPlugins} from './plugin/googleSheets'
   ```

2. In the constructor, after the protected function loading loop (line 67) and before the categorization loop (line 69), add:
   ```typescript
   if (config.compatibilityMode === 'googleSheets') {
     for (const gsPlugin of googleSheetsPlugins) {
       FunctionRegistry.loadPluginFunctions(gsPlugin, this.instancePlugins)
     }
   }
   ```

**Step 3: Run tests**

Run: `cross-env NODE_ICU_DATA=node_modules/full-icu npx jest test/compatibility-mode.spec.ts -v`

Expected: All PASS (the plugins array is empty for now, but the wiring is in place).

**Step 4: Run full test suite**

Run: `npm run test:jest`

Expected: No regressions.

**Step 5: Commit**

```bash
git add src/interpreter/FunctionRegistry.ts test/compatibility-mode.spec.ts
git commit -m "feat: wire Google Sheets plugins into FunctionRegistry

When compatibilityMode is 'googleSheets', the FunctionRegistry
constructor loads GSheets override plugins on top of the default
plugins. Override plugins silently replace specific function
implementations via Map.set."
```

---

### Task 7: Create GoogleSheetsTextPlugin (SPLIT override)

This is the first real override plugin. It demonstrates the pattern for all subsequent plugins.

**Files:**
- Create: `src/interpreter/plugin/googleSheets/GoogleSheetsTextPlugin.ts`
- Modify: `src/interpreter/plugin/googleSheets/index.ts`
- Test: `test/google-sheets-text-plugin.spec.ts` (create)

**Step 1: Write the failing test**

Create `test/google-sheets-text-plugin.spec.ts`:

```typescript
import {HyperFormula} from '../src'
import {adr} from './testUtils'

describe('GoogleSheetsTextPlugin - SPLIT', () => {
  it('should split by delimiter and return array', () => {
    const hf = HyperFormula.buildFromArray([
      ['=SPLIT("hello world", " ")'],
    ], {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'})

    expect(hf.getCellValue(adr('A1'))).toBe('hello')
    expect(hf.getCellValue(adr('B1'))).toBe('world')
    hf.destroy()
  })

  it('should split by each character in delimiter when split_by_each is true', () => {
    const hf = HyperFormula.buildFromArray([
      ['=SPLIT("a-b.c", "-.")'],
    ], {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'})

    // Default split_by_each is TRUE in GSheets
    expect(hf.getCellValue(adr('A1'))).toBe('a')
    expect(hf.getCellValue(adr('B1'))).toBe('b')
    expect(hf.getCellValue(adr('C1'))).toBe('c')
    hf.destroy()
  })

  it('should split by whole delimiter when split_by_each is false', () => {
    const hf = HyperFormula.buildFromArray([
      ['=SPLIT("a-b.c-d", "-.", FALSE())'],
    ], {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'})

    // With split_by_each=FALSE, "-." is treated as a single delimiter
    expect(hf.getCellValue(adr('A1'))).toBe('a-b')
    expect(hf.getCellValue(adr('B1'))).toBe('c-d')
    hf.destroy()
  })

  it('should remove empty text by default', () => {
    const hf = HyperFormula.buildFromArray([
      ['=SPLIT("a,,b", ",")'],
    ], {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'})

    // Default remove_empty_text is TRUE — empty strings are removed
    expect(hf.getCellValue(adr('A1'))).toBe('a')
    expect(hf.getCellValue(adr('B1'))).toBe('b')
    hf.destroy()
  })

  it('should keep empty text when remove_empty_text is false', () => {
    const hf = HyperFormula.buildFromArray([
      ['=SPLIT("a,,b", ",", TRUE(), FALSE())'],
    ], {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'})

    expect(hf.getCellValue(adr('A1'))).toBe('a')
    expect(hf.getCellValue(adr('B1'))).toBe('')
    expect(hf.getCellValue(adr('C1'))).toBe('b')
    hf.destroy()
  })

  it('should return VALUE error for empty delimiter', () => {
    const hf = HyperFormula.buildFromArray([
      ['=SPLIT("abc", "")'],
    ], {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'})

    const val = hf.getCellValue(adr('A1'))
    expect(val).toBeInstanceOf(Object) // DetailedCellError
    hf.destroy()
  })

  it('should use original SPLIT behavior in default mode', () => {
    const hf = HyperFormula.buildFromArray([
      ['=SPLIT("hello world", 0)'],
    ], {licenseKey: 'gpl-v3', compatibilityMode: 'default'})

    // Original HF SPLIT: SPLIT(string, index) returns word at index
    expect(hf.getCellValue(adr('A1'))).toBe('hello')
    hf.destroy()
  })
})
```

**Step 2: Run test to verify it fails**

Run: `cross-env NODE_ICU_DATA=node_modules/full-icu npx jest test/google-sheets-text-plugin.spec.ts -v`

Expected: FAIL — SPLIT still uses the original `(string, number)` signature.

**Step 3: Implement GoogleSheetsTextPlugin**

Create `src/interpreter/plugin/googleSheets/GoogleSheetsTextPlugin.ts`:

```typescript
/**
 * @license
 * Copyright (c) 2025 Handsoncode. All rights reserved.
 */

import {CellError, ErrorType} from '../../../Cell'
import {ErrorMessage} from '../../../error-message'
import {ProcedureAst} from '../../../parser'
import {InterpreterState} from '../../InterpreterState'
import {InterpreterValue} from '../../InterpreterValue'
import {SimpleRangeValue} from '../../../SimpleRangeValue'
import {FunctionArgumentType, FunctionPlugin, FunctionPluginTypecheck, ImplementedFunctions} from '../FunctionPlugin'

/**
 * Google Sheets-compatible text function overrides.
 *
 * Overrides SPLIT to use Google Sheets signature:
 * SPLIT(text, delimiter, [split_by_each], [remove_empty_text])
 */
export class GoogleSheetsTextPlugin extends FunctionPlugin implements FunctionPluginTypecheck<GoogleSheetsTextPlugin> {
  public static implementedFunctions: ImplementedFunctions = {
    'SPLIT': {
      method: 'split',
      parameters: [
        {argumentType: FunctionArgumentType.STRING},
        {argumentType: FunctionArgumentType.STRING},
        {argumentType: FunctionArgumentType.BOOLEAN, defaultValue: true},
        {argumentType: FunctionArgumentType.BOOLEAN, defaultValue: true},
      ],
      vectorizationForbidden: true,
    },
  }

  /**
   * SPLIT(text, delimiter, [split_by_each], [remove_empty_text])
   *
   * Splits text by delimiter and returns a horizontal array.
   * - split_by_each (default TRUE): if TRUE, each character in delimiter is a separate delimiter
   * - remove_empty_text (default TRUE): if TRUE, empty strings are removed from results
   */
  public split(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('SPLIT'),
      (text: string, delimiter: string, splitByEach: boolean, removeEmptyText: boolean) => {
        if (delimiter === '') {
          return new CellError(ErrorType.VALUE, ErrorMessage.EmptyString)
        }

        let parts: string[]

        if (splitByEach) {
          // Each character in delimiter is a separate delimiter
          const regex = new RegExp(
            delimiter.split('').map(c => c.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')).join('|')
          )
          parts = text.split(regex)
        } else {
          // Treat delimiter as a whole string
          parts = text.split(delimiter)
        }

        if (removeEmptyText) {
          parts = parts.filter(p => p !== '')
        }

        if (parts.length === 0) {
          return new CellError(ErrorType.VALUE, ErrorMessage.EmptyString)
        }

        return SimpleRangeValue.onlyValues([parts])
      }
    )
  }
}
```

**Step 4: Register the plugin in the index**

Update `src/interpreter/plugin/googleSheets/index.ts`:

```typescript
/**
 * @license
 * Copyright (c) 2025 Handsoncode. All rights reserved.
 */

import {FunctionPluginDefinition} from '../FunctionPlugin'
import {GoogleSheetsTextPlugin} from './GoogleSheetsTextPlugin'

/**
 * All Google Sheets compatibility plugins.
 * These override specific built-in functions with Google Sheets-compatible behavior.
 */
export const googleSheetsPlugins: FunctionPluginDefinition[] = [
  GoogleSheetsTextPlugin,
]

export {GoogleSheetsTextPlugin}
```

**Step 5: Run test to verify it passes**

Run: `cross-env NODE_ICU_DATA=node_modules/full-icu npx jest test/google-sheets-text-plugin.spec.ts -v`

Expected: All PASS.

**Step 6: Run full test suite**

Run: `npm run test:jest`

Expected: No regressions. The default-mode SPLIT test should still pass because the override only applies in `googleSheets` mode.

**Step 7: Commit**

```bash
git add src/interpreter/plugin/googleSheets/GoogleSheetsTextPlugin.ts \
        src/interpreter/plugin/googleSheets/index.ts \
        test/google-sheets-text-plugin.spec.ts
git commit -m "feat: add GoogleSheetsTextPlugin with SPLIT override

Implement Google Sheets-compatible SPLIT function:
SPLIT(text, delimiter, [split_by_each], [remove_empty_text])

Returns a horizontal array of strings, splitting by each character
in the delimiter by default. Only active in googleSheets mode;
default mode retains the original SPLIT(string, index) behavior."
```

---

### Task 8: Verify `EmptyString` error message exists

**Important pre-check:** The `GoogleSheetsTextPlugin` references `ErrorMessage.EmptyString`. Verify this exists:

Run: `grep -r "EmptyString" src/error-message.ts`

If it doesn't exist, add it to `src/error-message.ts`:

```typescript
public static EmptyString = 'Empty string is not allowed.'
```

This should be done before running the GoogleSheetsTextPlugin tests. If the message already exists under a different name, use that instead.

---

### Remaining Tasks (PRs C-G from design doc)

The following tasks follow the exact same pattern as Task 7. Each creates a plugin file, adds it to the index, and writes tests. They are listed here as future work items to be broken into separate PRs:

**PR C — GoogleSheetsStatisticalPlugin:**
- Functions: `POISSON.DIST`, `BETA.DIST`, `BETA.INV`, `BINOM.DIST`, `BINOM.INV`, `NEGBINOM.DIST`, `HYPGEOM.DIST`, `T.INV`, `T.INV.2T`, `T.DIST`, `TDIST`, `WEIBULL.DIST`, `GAMMA`, `NORMSDIST`, `CHISQ.TEST`, `LARGE`, `AVEDEV`, `SKEW`, `DEVSQ`
- Each function's GSheets-specific behavior is documented in `docs/guide/list-of-differences.md` lines 47-101
- Reference existing implementations in `src/interpreter/plugin/StatisticalPlugin.ts` and `src/interpreter/plugin/StatisticalAggregationPlugin.ts`

**PR D — GoogleSheetsFinancialPlugin:**
- Functions: `TBILLEQ`, `TBILLPRICE`, `NPV`, `XNPV`, `RATE`, `PV`, `DB`, `RRI`
- Reference: `src/interpreter/plugin/FinancialPlugin.ts`

**PR E — Boolean coercion infrastructure + GoogleSheetsLogicalPlugin:**
- Modify `src/interpreter/plugin/NumericAggregationPlugin.ts` reduce() method (~line 588) for GSheets boolean coercion
- Functions: `COUNTA`, `ADDRESS`
- Reference: `src/interpreter/plugin/NumericAggregationPlugin.ts`, `src/interpreter/plugin/AddressPlugin.ts`

**PR F — GoogleSheetsMathPlugin + GoogleSheetsDatePlugin:**
- Math: `GCD`, `LCM`, `COMBIN` — Reference: `src/interpreter/plugin/MathPlugin.ts`
- Date: `DAYS`, `DATEDIF` — Reference: `src/interpreter/plugin/DateTimePlugin.ts`

Each of these PRs follows the same structure:
1. Write failing tests based on the expected values in `docs/guide/list-of-differences.md`
2. Implement the override plugin
3. Add to `googleSheetsPlugins` array in index
4. Verify no regressions
5. Commit
