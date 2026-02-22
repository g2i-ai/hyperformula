const baseConfigFactory = require('./base')

/**
 * Creates Karma development configuration.
 *
 * This mirrors the base Karma setup so NODE_ENV=development resolves
 * to the same configuration shape expected by karma.conf.js.
 *
 * @param {import('karma').Config} config - Karma runtime config object.
 * @returns {import('karma').ConfigOptions} Resolved Karma configuration.
 */
module.exports.create = (config) => baseConfigFactory.create(config)
