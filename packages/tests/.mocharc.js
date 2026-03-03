module.exports = {
    timeout: 20 * 60 * 1000,
    exit: true,
    "node-option": ["no-experimental-strip-types"],
    reporter: "mochawesome",
    require: "ts-node/register"
};
