/**
 * See:
 * - https://mochajs.org/#usage for more information on usage of mocha flags.
 * - https://github.com/karma-runner/karma-mocha for more information on all mocha flags which the
 *   karma runner supports.
 */

const config = {
  "node-option": ["no-experimental-strip-types"],
  require: "ts-node/register",
  timeout: 5000,
  retries: 2,
  exit: true,
};

module.exports = config;
