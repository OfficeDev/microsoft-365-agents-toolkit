"use strict";

module.exports = {
  color: true,
  delay: false,
  diff: true,
  "node-option": ["no-experimental-strip-types"],
  parallel: false,
  recursive: false,
  reporter: "mochawesome",
  require: ["ts-node/register", "tests/e2e/setup.ts"],
  retries: 0,
  slow: "75",
  timeout: 1200000,
  extensions: ["ts", "tsx"],
};
