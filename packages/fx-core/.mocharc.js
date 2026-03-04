"use strict";

const nodeOptions = [];
const [major] = process.versions.node.split(".").map(Number);
if (major >= 22) {
  nodeOptions.push("no-experimental-strip-types");
} else {
  nodeOptions.push("loader=ts-node/esm");
}

module.exports = {
  color: true,
  delay: false,
  diff: true,
  "node-option": nodeOptions,
  parallel: false,
  recursive: false,
  reporter: "spec",
  require: "ts-node/register",
  retries: 1,
  slow: "75",
  timeout: 0,
  extensions: ["ts", "tsx"],
};
