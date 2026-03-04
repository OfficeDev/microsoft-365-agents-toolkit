/**
 * See:
 * - https://mochajs.org/#usage for more information on usage of mocha flags.
 * - https://github.com/karma-runner/karma-mocha for more information on all mocha flags which the
 *   karma runner supports.
 */

const nodeOptions = [];
const [major] = process.versions.node.split(".").map(Number);
if (major >= 22) {
  nodeOptions.push("no-experimental-strip-types");
}

const config = {
  "node-option": nodeOptions,
  require: "ts-node/register",
  timeout: 5000,
  retries: 2,
  exit: true,
};

module.exports = config;
