const nodeOptions = [];
const [major] = process.versions.node.split(".").map(Number);
if (major >= 22) {
  nodeOptions.push("no-experimental-strip-types");
}

module.exports = {
  extensions: ["ts"],
  "node-option": nodeOptions,
  spec: ["test/*.test.ts"],
  require: ["ts-node/register"],
};
