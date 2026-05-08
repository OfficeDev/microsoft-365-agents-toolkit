"use strict";

module.exports = {
  color: true,
  delay: false,
  diff: true,
  parallel: false,
  recursive: false,
  reporter: "spec",
  require: "ts-node/register",
  slow: "75",
  timeout: 0,
  extensions: ["ts", "tsx"],
  "node-option": ["no-experimental-strip-types"],
};
