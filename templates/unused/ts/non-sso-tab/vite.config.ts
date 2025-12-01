import { resolve } from "path";

async function getConfig() {
  const { defineConfig } = await import("vite");

  return defineConfig({
    define: {
      global: "globalThis",
    },
    build: {
      outDir: "lib/static",
      rollupOptions: {
        input: {
          m365agents: resolve(__dirname, "src/static/scripts/m365agents.ts"),
        },
        output: {
          entryFileNames: "scripts/[name].js",
          chunkFileNames: "scripts/[name].js",
          assetFileNames: "assets/[name].[ext]",
        },
      },
    },
  });
}

export default getConfig();
