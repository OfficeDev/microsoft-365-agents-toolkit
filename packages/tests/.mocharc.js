const nodeOptions = [];
const [major] = process.versions.node.split(".").map(Number);
if (major >= 22) {
  nodeOptions.push("no-experimental-strip-types");
}

module.exports = {
    timeout: 20 * 60 * 1000,
    exit: true,
    "node-option": nodeOptions,
    reporter: "mochawesome",
    require: "ts-node/register"
};
