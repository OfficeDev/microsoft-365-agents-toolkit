import json from "@rollup/plugin-json";
import terser from "@rollup/plugin-terser";
import typescript from "typescript";
import typescriptPlugin from "@rollup/plugin-typescript";
import pkg from "./package.json" assert { type: "json" };

const deps = Object.keys(Object.assign({}, pkg.peerDependencies, pkg.dependencies));

const nodeDeps = [...deps, "crypto", "fs", "path", "https"];

/**
 * ES5 Builds
 */
const es5BuildPlugins = [json(), terser(), typescriptPlugin({ typescript: typescript })];

/**
 * ES2017 Builds
 */
const es2017Plugins = [
  json({ preferConst: true }),
  terser(),
  typescriptPlugin({
    typescript: typescript,
    compilerOptions: {
      target: "es2017",
    },
  }),
];

const es5Builds = [
  /**
   * Browser Builds
   */
  {
    input: "src/index.browser.ts",
    output: [{ file: pkg.esm5, format: "es", sourcemap: true }],
    external: (id) => deps.some((dep) => id === dep || id.startsWith(`${dep}/`)),
    plugins: [...es5BuildPlugins],
    treeshake: {
      moduleSideEffects: false,
    },
  },
  /**
   * Node.js Build
   */
  {
    input: "src/index.ts",
    output: [{ file: pkg.main, format: "cjs", sourcemap: true }],
    external: (id) => nodeDeps.some((dep) => id === dep || id.startsWith(`${dep}/`)),
    plugins: [
      json(),
      terser(),
      typescriptPlugin({
        typescript: typescript,
        tsconfig: "./tsconfig.json",
        declarationDir: "./dist/types",
        declaration: true,
      }),
    ],
  },
];

const es2017Builds = [
  /**
   * Node.js Build
   */
  {
    input: "./src/index.ts",
    output: [
      {
        file: pkg.module,
        format: "es",
        sourcemap: true,
      },
    ],
    external: (id) => nodeDeps.some((dep) => id === dep || id.startsWith(`${dep}/`)),
    plugins: [...es2017Plugins],
  },
  /**
   * Browser Builds
   */
  {
    input: "./src/index.browser.ts",
    output: [
      {
        file: pkg.browser,
        format: "es",
        sourcemap: true,
      },
    ],
    external: (id) => deps.some((dep) => id === dep || id.startsWith(`${dep}/`)),
    plugins: [...es2017Plugins],
  },
];

export default [...es5Builds, ...es2017Builds];
