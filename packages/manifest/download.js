const fs = require("fs-extra");
const path = require("path");
const os = require("os");
const { pipeline } = require("stream");
const { promisify } = require("util");
const unzipper = require("unzipper");

const streamPipeline = promisify(pipeline);

async function downloadAndExtract() {
  const url = "https://github.com/microsoft/json-schemas/archive/refs/heads/live.zip";
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), "json-schemas-"));
  const zipPath = path.join(tempDir, "live.zip");

  console.log(`Downloading zip to ${zipPath}`);
  const res = await fetch(url);
  if (!res.ok) throw new Error(`Failed to download: ${res.statusText}`);

  await streamPipeline(res.body, fs.createWriteStream(zipPath));

  console.log("Extracting zip...");
  await fs
    .createReadStream(zipPath)
    .pipe(unzipper.Extract({ path: tempDir }))
    .promise();

  const extractedRoot = path.join(tempDir, "json-schemas-live");
  const schemas = ["teams", "copilot"];
  const targetRoot = path.join(__dirname, "src", "json-schemas");

  for (const name of schemas) {
    const srcDir = path.join(extractedRoot, name);
    const destDir = path.join(targetRoot, name);

    if (!(await fs.pathExists(srcDir))) {
      console.warn(`Source directory not found: ${srcDir}`);
      continue;
    }

    // Remove and recreate dest
    await fs.remove(destDir);
    await fs.copy(srcDir, destDir);
    console.log(`Replaced ${name} schemas.`);
  }

  // cleanup
  await fs.remove(tempDir);
  console.log("Done.");
}

// Ensure global fetch (Node 18+), otherwise import node-fetch lazily
if (typeof fetch === "undefined") {
  global.fetch = (...args) => import("node-fetch").then(({ default: f }) => f(...args));
}

(async () => {
  try {
    await downloadAndExtract();
  } catch (err) {
    console.error(err);
    process.exit(1);
  }
})();

/**
 * Dependencies to install:
 *   npm install fs-extra unzipper
 *
 * Optional (if using Node <18):
 *   npm install node-fetch
 */
