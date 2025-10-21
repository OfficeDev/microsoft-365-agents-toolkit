const AdmZip = require("adm-zip");
const { readdirSync, mkdirSync } = require("node:fs");
const path = require("path");

const BUILD_PATH = path.join(__dirname, "..", "build", "fallback");
const TEMPLATE_NAMES = ["common", "js", "ts", "python"];

mkdirSync(BUILD_PATH, { recursive: true });

TEMPLATE_NAMES.forEach((name) => {
  const zip = new AdmZip();
  const templatePath = path.join(__dirname, "..", "vsc", name);
  readdirSync(templatePath).forEach((dir) => {
    zip.addLocalFolder(path.join(templatePath, dir), dir);
  });
  console.log(`Generating ${name}.zip`);
  zip.writeZip(path.join(BUILD_PATH, `${name}.zip`));
});
