const nodeOptions = [];
const [major] = process.versions.node.split(".").map(Number);
if (major >= 22) {
  nodeOptions.push("no-experimental-strip-types");
}

module.exports = {
  "node-option": nodeOptions,
  require: ["build/test/setup.js", "ts-node/register"],
  reporter: "mocha-multi-reporters",
  recursive: true,
  colors: true,
};
