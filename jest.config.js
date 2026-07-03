module.exports = {
  testEnvironment: 'node',
  setupFilesAfterEnv: ['./tests/setup.js'],
  reporters: [
    'default',
    './tests/reporters/agent-reporter.js'
  ],
  verbose: true,
  testMatch: ['**/*.test.js', '!**/node_modules/**']
};
