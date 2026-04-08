"use strict";

module.exports = {
  color: true,
  delay: false,
  diff: true,
  "node-option": ["no-experimental-strip-types"],
  parallel: false,
  recursive: false,
  reporter: "spec",
  require: "ts-node/register",
  retries: 1,
  slow: "75",
  timeout: 0,
  extensions: ["ts", "tsx"],
};
