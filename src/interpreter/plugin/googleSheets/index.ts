/**
 * @license
 * Copyright (c) 2025 Handsoncode. All rights reserved.
 */

import type { FunctionPluginDefinition } from '../FunctionPlugin'
import {GoogleSheetsEngineeringFixesPlugin} from './GoogleSheetsEngineeringFixesPlugin'
import {GoogleSheetsStatisticalFixesPlugin} from './GoogleSheetsStatisticalFixesPlugin'
import {GoogleSheetsTextPlugin} from './GoogleSheetsTextPlugin'

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
  GoogleSheetsEngineeringFixesPlugin,
  GoogleSheetsStatisticalFixesPlugin,
]

export {GoogleSheetsTextPlugin, GoogleSheetsEngineeringFixesPlugin, GoogleSheetsStatisticalFixesPlugin}
