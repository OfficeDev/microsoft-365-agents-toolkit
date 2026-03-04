const nodeOptions = [];
const [major] = process.versions.node.split(".").map(Number);
if (major >= 22) {
  nodeOptions.push("no-experimental-strip-types");
}

module.exports = {
  exit: true,
  "node-option": nodeOptions,
  timeout: 0,
  reporter: "spec",
  require: "ts-node/register",
};
