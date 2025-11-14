/*
 * Script: generateRelease.js
 * Purpose: Prepare release artifacts by copying fallback zips and bundling metadata/ui/resources into a single metadata.zip
 */

const fs = require("fs");
const path = require("path");
const fsp = fs.promises;
const AdmZip = require("adm-zip");

async function ensureDir(dir) {
  await fsp.mkdir(dir, { recursive: true });
}

function addIfExists(zip, srcPath, zipFolderName) {
  if (fs.existsSync(srcPath)) {
    // Add directory contents under specified folder name
    zip.addLocalFolder(srcPath, zipFolderName);
    console.log(`[zip] Added directory '${srcPath}' as '${zipFolderName}/'`);
  } else {
    console.warn(`[zip] Skipped missing directory: ${srcPath}`);
  }
}

async function buildMetadataZip(releaseDir) {
  const zip = new AdmZip();
  const buildMetadataDir = path.resolve("build", "metadata");
  const buildUiDir = path.resolve("build", "ui");

  // User requested src/resource/* but actual existing path is src/ui/resource
  const srcUiResourceDir = path.resolve("src", "ui", "resource");
  const srcTopResourceDir = path.resolve("src", "resource"); // In case it appears in future

  addIfExists(zip, buildMetadataDir, "metadata");
  addIfExists(zip, buildUiDir, "ui");
  // Prefer top-level src/resource if it exists, else fall back to src/ui/resource
  if (fs.existsSync(srcTopResourceDir)) {
    addIfExists(zip, srcTopResourceDir, "resource");
  } else {
    addIfExists(zip, srcUiResourceDir, "resource");
  }

  const outPath = path.join(releaseDir, "metadata.zip");
  zip.writeZip(outPath);
  console.log(`[zip] Created ${outPath}`);
}

async function main() {
  try {
    const fallbackDir = path.resolve("build", "fallback");
    const releaseDir = path.resolve("build");

    await ensureDir(releaseDir);

    if (!fs.existsSync(fallbackDir)) {
      console.error(`Missing fallback directory: ${fallbackDir}`);
      process.exit(1);
    }

    console.log("Creating metadata.zip...");
    await buildMetadataZip(releaseDir);

    console.log("Release preparation complete.");
  } catch (err) {
    console.error("Error generating release:", err);
    process.exit(1);
  }
}

if (require.main === module) {
  main();
}
