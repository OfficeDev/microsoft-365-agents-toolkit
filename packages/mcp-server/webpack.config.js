const path = require("path");
const nodeExternals = require("webpack-node-externals");
const TerserPlugin = require("terser-webpack-plugin");
const webpack = require("webpack");

is_production = process.argv[process.argv.indexOf("--mode") + 1] === "production";
console.log("is_production: ", is_production);

module.exports = {
  target: "node",
  mode: is_production ? "production" : "development",
  entry: "./src/index.ts",
  devtool: "source-map",
  externals: [nodeExternals()],
  module: {
    rules: [
      {
        test: /\.tsx?$/,
        use: "ts-loader",
        exclude: /node_modules/,
      },
    ],
  },
  resolve: {
    extensions: [".ts", ".js"],
  },
  output: {
    filename: "index.js",
    path: path.resolve(__dirname, "dist"),
    clean: true,
  },
  plugins: [new webpack.BannerPlugin({ banner: "#!/usr/bin/env node", raw: true })],
  optimization: {
    minimize: is_production,
    minimizer: [
      new TerserPlugin({
        terserOptions: {
          mangle: true,
          format: {
            comments: false,
          },
        },
        extractComments: false,
      }),
    ],
  },
};
