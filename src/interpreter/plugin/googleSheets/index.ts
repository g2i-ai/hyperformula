/**
 * @license
 * Copyright (c) 2025 Handsoncode. All rights reserved.
 */

import type { FunctionPluginDefinition } from '../FunctionPlugin'
import {GoogleSheetsTextPlugin} from './GoogleSheetsTextPlugin'
import {GoogleSheetsOperatorPlugin} from './GoogleSheetsOperatorPlugin'
import {GoogleSheetsTextFunctionsPlugin} from './GoogleSheetsTextFunctionsPlugin'
import {GoogleSheetsInfoPlugin} from './GoogleSheetsInfoPlugin'
import {GoogleSheetsConversionPlugin} from './GoogleSheetsConversionPlugin'
import {GoogleSheetsDatabasePlugin} from './GoogleSheetsDatabasePlugin'
import {GoogleSheetsStatisticalPlugin} from './GoogleSheetsStatisticalPlugin'
import {GoogleSheetsFinancialPlugin} from './GoogleSheetsFinancialPlugin'
import {GoogleSheetsArrayPlugin} from './GoogleSheetsArrayPlugin'
import {GoogleSheetsMiscPlugin} from './GoogleSheetsMiscPlugin'
import {GoogleSheetsEngineeringFixesPlugin} from './GoogleSheetsEngineeringFixesPlugin'
import {GoogleSheetsStatisticalFixesPlugin} from './GoogleSheetsStatisticalFixesPlugin'

/**
 * Google Sheets override plugins.
 *
 * Each plugin extends FunctionPlugin and only declares the functions it
 * overrides. When `compatibilityMode === 'googleSheets'`, these are loaded
 * on top of the default plugins in FunctionRegistry, silently replacing
 * the overridden function implementations via Map.set.
 */
export const googleSheetsPlugins: FunctionPluginDefinition[] = [
  GoogleSheetsTextPlugin,
  GoogleSheetsOperatorPlugin,
  GoogleSheetsTextFunctionsPlugin,
  GoogleSheetsInfoPlugin,
  GoogleSheetsConversionPlugin,
  GoogleSheetsDatabasePlugin,
  GoogleSheetsStatisticalPlugin,
  GoogleSheetsFinancialPlugin,
  GoogleSheetsArrayPlugin,
  GoogleSheetsMiscPlugin,
  GoogleSheetsEngineeringFixesPlugin,
  GoogleSheetsStatisticalFixesPlugin,
]

export {GoogleSheetsTextPlugin}
export {GoogleSheetsOperatorPlugin}
export {GoogleSheetsTextFunctionsPlugin}
export {GoogleSheetsInfoPlugin}
export {GoogleSheetsConversionPlugin}
export {GoogleSheetsDatabasePlugin}
export {GoogleSheetsStatisticalPlugin}
export {GoogleSheetsFinancialPlugin}
export {GoogleSheetsArrayPlugin}
export {GoogleSheetsMiscPlugin}
export {GoogleSheetsEngineeringFixesPlugin}
export {GoogleSheetsStatisticalFixesPlugin}
