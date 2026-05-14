// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * esbuild configuration for cli-next.
 *
 * Bundles the entire CLI (including core-next) into a single CJS file to
 * reduce startup time and npm package size.
 *
 * Usage:
 *   node esbuild.mjs              # development build (no minification)
 *   node esbuild.mjs --production # production build (minified, keep names)
 */
import * as esbuild from "esbuild";

const production = process.argv.includes("--production");

/** Modules that contain native .node addons and cannot be bundled. */
const nativeExternals = ["keytar", "@azure/msal-node-extensions"];

/**
 * Modules with problematic dynamic require() patterns that break bundlers.
 * applicationinsights uses dynamic require for auto-collection modules.
 * @opentelemetry/tracing is an optional peer dep pulled in by applicationinsights.
 */
const dynamicExternals = [
  "applicationinsights",
  "applicationinsights-native-metrics",
  "@opentelemetry/tracing",
];

const external = [...nativeExternals, ...dynamicExternals];

async function main() {
  const result = await esbuild.build({
    entryPoints: ["./src/index.ts"],
    outfile: "build/index.js",
    bundle: true,
    format: "cjs",
    platform: "node",
    target: "node18",
    minify: production,
    keepNames: true, // Error names appear in telemetry — never mangle them
    sourcemap: true,
    sourcesContent: false,
    external,
    // Ensure JSON imports (e.g. package.json reads) work
    loader: { ".json": "json" },
    // Use main field first (most packages), then module (ESM-first packages)
    mainFields: ["main", "module"],
    logLevel: "info",
    metafile: true, // Generate build metadata for analysis
  });

  // Write metafile for bundle analysis (e.g. `npx esbuild-visualizer --metadata build/meta.json`)
  const fs = await import("fs");
  fs.writeFileSync("build/meta.json", JSON.stringify(result.metafile));

  if (production) {
    // Report bundle size
    const stats = fs.statSync("build/index.js");
    const sizeKB = (stats.size / 1024).toFixed(1);
    console.log(`\nBundle size: ${sizeKB} KB (production)`);
  }
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
