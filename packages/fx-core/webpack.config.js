//@ts-check

"use strict";

const path = require("path");
const webpack = require("webpack");
const CopyPlugin = require("copy-webpack-plugin");
const terserWebpackPlugin = require("terser-webpack-plugin");

/**@type {import('webpack').Configuration}*/
const config = {
  target: "node", // vscode extensions run in a Node.js-context 📖 -> https://webpack.js.org/configuration/node/
  //mode: 'none', // this leaves the source code as close as possible to the original (when packaging we set this to 'production')
  node: {
    __dirname: false,
  },

  entry: {
    actions: "./src/actions.ts",
    index: "./src/index.ts", // the entry point of this extension, 📖 -> https://webpack.js.org/configuration/entry-context/
  },
  output: {
    filename: "[name].js",
    // the bundle is stored in the 'dist' folder (check package.json), 📖 -> https://webpack.js.org/configuration/output/
    path: path.resolve(__dirname, "build"),
    libraryTarget: "umd",
    sourceMapFilename: "[name].js.map",
    // devtoolModuleFilenameTemplate: "../[resource-path]",
    umdNamedDefine: true,
    globalObject: `(typeof self !== 'undefined' ? self : this)`,
  },
  devtool: "source-map",
  externals: {
    commonjs: "commonjs", //Add other modules that cannot be webpack'ed, 📖 -> https://webpack.js.org/configuration/externals/
    keytar: "keytar",
  },
  resolve: {
    // support reading TypeScript and JavaScript files, 📖 -> https://github.com/TypeStrong/ts-loader
    extensions: [".tsx", ".ts", ".js", ".json"],
  },
  module: {
    rules: [
      {
        test: /(?<!\.d)\.tsx?$/,
        exclude: /node_modules/,
        use: [
          {
            loader: "ts-loader",
            // options: {
            //   configFile: path.resolve(__dirname, './tsconfig.json'),
            // },
          },
        ],
      },
      {
        test: /\.s[ac]ss$/i,
        exclude: /node_modules/,
        use: ["style-loader", "css-loader", "sass-loader"],
      },
      {
        test: /\.(jpg|png|svg|gif)$/,
        use: {
          loader: "url-loader",
        },
      },
    ],
  },
  plugins: [
    new webpack.ContextReplacementPlugin(/express[\/\\]lib/, false, /$^/),
    new webpack.ContextReplacementPlugin(
      /applicationinsights[\/\\]out[\/\\]AutoCollection/,
      false,
      /$^/
    ),
    new webpack.ContextReplacementPlugin(/applicationinsights[\/\\]out[\/\\]Library/, false, /$^/),
    new webpack.ContextReplacementPlugin(/ms-rest[\/\\]lib/, false, /$^/),
    new webpack.IgnorePlugin({ resourceRegExp: /@opentelemetry\/tracing/ }),
    new webpack.IgnorePlugin({ resourceRegExp: /applicationinsights-native-metrics/ }),
    new webpack.IgnorePlugin({ resourceRegExp: /original-fs/ }),
    // ignore node-gyp/bin/node-gyp.js since it's not used in runtime
    new webpack.NormalModuleReplacementPlugin(
      /node-gyp[\/\\]bin[\/\\]node-gyp.js/,
      "@npmcli/node-gyp"
    ),
  ],
  optimization: {
    minimizer: [
      new terserWebpackPlugin({
        terserOptions: {
          mangle: false,
          keep_fnames: true,
        },
      }),
    ],
  },
};
module.exports = config;
