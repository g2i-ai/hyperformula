/**
 * @license
 * Copyright (c) 2025 Handsoncode. All rights reserved.
 */

import type { FunctionPluginDefinition } from '../FunctionPlugin'
import {GoogleSheetsConversionPlugin} from './GoogleSheetsConversionPlugin'
import {GoogleSheetsDatabasePlugin} from './GoogleSheetsDatabasePlugin'
import {GoogleSheetsFinancialPlugin} from './GoogleSheetsFinancialPlugin'
import {GoogleSheetsInfoPlugin} from './GoogleSheetsInfoPlugin'
import {GoogleSheetsStatisticalPlugin} from './GoogleSheetsStatisticalPlugin'
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
  GoogleSheetsConversionPlugin,
  GoogleSheetsDatabasePlugin,
  GoogleSheetsFinancialPlugin,
  GoogleSheetsInfoPlugin,
  GoogleSheetsStatisticalPlugin,
  GoogleSheetsTextPlugin,
  GoogleSheetsTextFunctionsPlugin,
  GoogleSheetsOperatorPlugin,
]

export {GoogleSheetsConversionPlugin, GoogleSheetsDatabasePlugin, GoogleSheetsFinancialPlugin, GoogleSheetsInfoPlugin, GoogleSheetsStatisticalPlugin, GoogleSheetsTextPlugin, GoogleSheetsTextFunctionsPlugin, GoogleSheetsOperatorPlugin}
