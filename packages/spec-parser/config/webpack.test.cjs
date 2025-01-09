const webpack = require("webpack");
module.exports = {
  mode: "development",
  devtool: "source-map",
  module: {
    rules: [
      {
        test: /\.tsx?$/,
        use: {
          loader: "ts-loader",
          options: {
            compilerOptions: {
              module: "commonjs",
              target: "es2017",
              downlevelIteration: true,
              resolveJsonModule: true,
            },
          },
        },
      },
      {
        test: /\.[tj]sx?$/,
        exclude: /node_modules/,
        use: "source-map-loader",
        enforce: "pre",
      },
    ],
  },
  resolve: {
    modules: ["node_modules"],
    mainFields: ["browser", "module", "main"],
    extensions: [".js", ".ts"],
    symlinks: true,
    fallback: {
      url: require.resolve("url/"),
      http: require.resolve("stream-http"),
      https: require.resolve("https-browserify"),
      stream: require.resolve("stream-browserify"),
      buffer: require.resolve("buffer"),
      "process/browser": require.resolve("process/browser"),
      util: false,
      path: require.resolve("path-browserify"),
    },
  },
  plugins: [
    new webpack.ProvidePlugin({
      process: "process/browser",
    }),
    new webpack.ProvidePlugin({
      Buffer: ["buffer", "Buffer"],
    }),
  ],
};
