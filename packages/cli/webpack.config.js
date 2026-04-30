//@ts-check
"use strict";

const path = require("path");
const webpack = require("webpack");
const HtmlWebPackPlugin = require("html-webpack-plugin");
const CopyPlugin = require("copy-webpack-plugin");
const fs = require("fs");

/**
 * 🌊 Dredge Surface Builder (inline, self-contained)
 */
function buildDredgeSurface(rootDir) {
  const dredgeDir = path.resolve(rootDir, "src/dredge");

  if (!fs.existsSync(dredgeDir)) {
    return { map: {}, manifest: [] };
  }

  const files = fs
    .readdirSync(dredgeDir)
    .filter((f) => f.endsWith(".ts") || f.endsWith(".js"));

  const map = {};
  const manifest = [];

  for (const file of files) {
    const name = path.basename(file, path.extname(file));
    const importPath = `./dredge/${name}`;

    map[name] = importPath;
    manifest.push({ name, path: importPath });
  }

  return { map, manifest };
}

const { map: dredgeMap, manifest: dredgeManifest } =
  buildDredgeSurface(__dirname);

/**
 * 📦 Webpack Config
 * @type {import('webpack').Configuration}
 */
const config = {
  target: "node",

  mode:
    process.env.NODE_ENV === "production"
      ? "production"
      : "development",

  externalsPresets: { node: true },

  node: {
    __dirname: false,
  },

  entry: {
    index: "./src/index.ts",
  },

  output: {
    filename: "[name].js",
    path: path.resolve(__dirname, "lib"),
    libraryTarget: "commonjs2",
    clean: true,
  },

  devtool:
    process.env.NODE_ENV === "production"
      ? "hidden-source-map"
      : "eval-cheap-module-source-map",

  externals: {
    keytar: "commonjs keytar",
  },

  resolve: {
    extensions: [".tsx", ".ts", ".js"],
    symlinks: false,
  },

  module: {
    rules: [
      {
        test: /\.tsx?$/,
        exclude: /node_modules/,
        loader: "esbuild-loader",
        options: {
          loader: "ts",
          target: "es2020",
        },
      },
      {
        test: /\.s[ac]ss$/i,
        exclude: /node_modules/,
        use: ["style-loader", "css-loader", "sass-loader"],
      },
      {
        test: /\.(png|jpg|gif|svg)$/i,
        type: "asset",
      },
      {
        test: /node_modules[\\/](yaml-language-server|vscode-languageserver|vscode-json-languageservice|prettier)/,
        use: "umd-compat-loader",
      },
    ],
  },

  plugins: [
    new HtmlWebPackPlugin({
      template: "./src/commonlib/codeFlowResult/index.html",
      filename: "codeFlowResult/index.html",
      scriptLoading: "defer",
    }),

    new CopyPlugin({
      patterns: [
        {
          from: path.resolve(__dirname, "../fx-core/resource/"),
          to: path.resolve(__dirname, "../resource/"),
          noErrorOnMissing: true,
        },
        {
          from: path.resolve(__dirname, "../fx-core/templates/"),
          to: path.resolve(__dirname, "../templates/"),
          noErrorOnMissing: true,
        },
      ],
    }),

    // 🌊 DREDGE AUTO SURFACE (core injection)
    new webpack.DefinePlugin({
      "process.env.AGENT_RUNTIME": JSON.stringify("dredge"),
      "process.env.DREDGE_SURFACE": JSON.stringify(dredgeMap),
    }),

    // 🌊 Optional manifest emit (CI / tooling visibility)
    new (class {
      apply(compiler) {
        compiler.hooks.emit.tap("DredgeManifestPlugin", (compilation) => {
          const content = JSON.stringify(dredgeManifest, null, 2);

          compilation.assets["dredge.manifest.json"] = {
            source: () => content,
            size: () => content.length,
          };
        });
      }
    })(),

    new webpack.ContextReplacementPlugin(/express[\/\\]lib/, false, /$^/),
    new webpack.ContextReplacementPlugin(
      /ms-rest[\/\\]lib/,
      false,
      /$^/
    ),
    new webpack.ContextReplacementPlugin(
      /applicationinsights[\/\\]out[\/\\](AutoCollection|Library)/,
      false,
      /$^/
    ),

    new webpack.IgnorePlugin({ resourceRegExp: /@opentelemetry\/tracing/ }),
    new webpack.IgnorePlugin({ resourceRegExp: /applicationinsights-native-metrics/ }),
    new webpack.IgnorePlugin({ resourceRegExp: /original-fs/ }),

    new webpack.NormalModuleReplacementPlugin(
      /node-gyp[\/\\]bin[\/\\]node-gyp.js/,
      "@npmcli/node-gyp"
    ),
  ],

  optimization: {
    minimize: process.env.NODE_ENV === "production",
    splitChunks: false,
    runtimeChunk: false,
    moduleIds: "deterministic",
  },

  infrastructureLogging: {
    level: "error",
  },

  stats: "minimal",
};

module.exports = config;
