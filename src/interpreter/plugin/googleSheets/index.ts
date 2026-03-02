/**
 * @license
 * Copyright (c) 2025 Handsoncode. All rights reserved.
 */

import type { FunctionPluginDefinition } from '../FunctionPlugin'
import {GoogleSheetsArrayPlugin} from './GoogleSheetsArrayPlugin'
import {GoogleSheetsConversionPlugin} from './GoogleSheetsConversionPlugin'
import {GoogleSheetsDatabasePlugin} from './GoogleSheetsDatabasePlugin'
import {GoogleSheetsEngineeringFixesPlugin} from './GoogleSheetsEngineeringFixesPlugin'
import {GoogleSheetsFinancialPlugin} from './GoogleSheetsFinancialPlugin'
import {GoogleSheetsInfoPlugin} from './GoogleSheetsInfoPlugin'
import {GoogleSheetsMiscPlugin} from './GoogleSheetsMiscPlugin'
import {GoogleSheetsStatisticalPlugin} from './GoogleSheetsStatisticalPlugin'
import {GoogleSheetsStatisticalFixesPlugin} from './GoogleSheetsStatisticalFixesPlugin'
import {GoogleSheetsTextPlugin} from './GoogleSheetsTextPlugin'
import {GoogleSheetsTextFunctionsPlugin} from './GoogleSheetsTextFunctionsPlugin'
import {GoogleSheetsOperatorPlugin} from './GoogleSheetsOperatorPlugin'

/**
 * Google Sheets override plugins.
 *
 * Each plugin extends FunctionPlugin and only declares the functions it
 * overrides. When `compatibilityMode === 'googleSheets'`, these are loaded
 * on top of the default plugins in FunctionRegistry, silently replacing
 * the overridden function implementations via Map.set.
 */
export const googleSheetsPlugins: FunctionPluginDefinition[] = [
  GoogleSheetsArrayPlugin,
  GoogleSheetsConversionPlugin,
  GoogleSheetsDatabasePlugin,
  GoogleSheetsEngineeringFixesPlugin,
  GoogleSheetsFinancialPlugin,
  GoogleSheetsInfoPlugin,
  GoogleSheetsMiscPlugin,
  GoogleSheetsStatisticalPlugin,
  GoogleSheetsStatisticalFixesPlugin,
  GoogleSheetsTextPlugin,
  GoogleSheetsTextFunctionsPlugin,
  GoogleSheetsOperatorPlugin,
]

export {GoogleSheetsArrayPlugin, GoogleSheetsConversionPlugin, GoogleSheetsDatabasePlugin, GoogleSheetsEngineeringFixesPlugin, GoogleSheetsFinancialPlugin, GoogleSheetsInfoPlugin, GoogleSheetsMiscPlugin, GoogleSheetsStatisticalPlugin, GoogleSheetsStatisticalFixesPlugin, GoogleSheetsTextPlugin, GoogleSheetsTextFunctionsPlugin, GoogleSheetsOperatorPlugin}
