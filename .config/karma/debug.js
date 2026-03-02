const configFactory = require('./base');

module.exports.create = function(config) {
  const configBase = configFactory.create(config);

  return {
    ...configBase,
    client: {
      ...configBase.client,
      clearContext: false,
    },
    browsers: ['Chrome'],
    reporters: ['kjhtml'],
    singleRun: false,
    autoWatch: true,
  }
}
