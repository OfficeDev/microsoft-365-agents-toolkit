const fs = require("fs-extra");
const path = require("path");
const { exec } = require("child_process");

async function convert(srcFolder, srcFileName, destFilePath) {
  const command = `npx quicktype -s schema ${srcFileName} -o ${destFilePath} --prefer-const-values --prefer-unions`;
  console.log(`Running command: ${command} in ${srcFolder}`);
  return new Promise((resolve) => {
    exec(command, { cwd: srcFolder }, (error, stdout, stderr) => {
      if (error) {
        console.error(`Error generating TypeScript for ${srcFileName}:`, error);
        return;
      }
      if (stderr) {
        console.error(`stderr for ${srcFileName}:`, stderr);
      }
      console.log(`Generated TypeScript file: ${destFilePath}`);
      resolve();
    });
  });
}

async function generateTeamsManifestTypeFiles() {
  const schemaFolder = path.join(__dirname, "src", "json-schemas", "teams");
  const outputFolder = path.join(__dirname, "src", "generated-types", "teams");
  if (!(await fs.pathExists(outputFolder))) {
    await fs.ensureDir(outputFolder);
  }
  const versions = await fs.readdir(schemaFolder);
  for (const version of versions) {
    const convertedVersion = version.replace(".", "D").replace("v", "V");
    const schemaDir = path.join(schemaFolder, version);
    const schemaFile = "MicrosoftTeams.schema.json";
    const outputFileName = `TeamsManifest${convertedVersion}.ts`;
    const outputPath = path.join(outputFolder, outputFileName);
    await convert(schemaDir, schemaFile, outputPath);
  }
}

async function generateDAManifestTypeFiles() {
  const schemaFolder = path.join(__dirname, "src", "json-schemas", "copilot", "declarative-agent");
  const outputFolder = path.join(
    __dirname,
    "src",
    "generated-types",
    "copilot",
    "declarative-agent"
  );
  if (!(await fs.pathExists(outputFolder))) {
    await fs.ensureDir(outputFolder);
  }
  const versions = await fs.readdir(schemaFolder);
  for (const version of versions) {
    const convertedVersion = version.replace(".", "D").replace("v", "V");
    const schemaDir = path.join(schemaFolder, version);
    const schemaFile = "schema.json";
    const outputFileName = `DeclarativeAgentManifest${convertedVersion}.ts`;
    const outputPath = path.join(outputFolder, outputFileName);
    await convert(schemaDir, schemaFile, outputPath);
  }
}

async function generatePluginManifestTypeFiles() {
  const schemaFolder = path.join(__dirname, "src", "json-schemas", "copilot", "plugin");
  const outputFolder = path.join(__dirname, "src", "generated-types", "copilot", "plugin");
  if (!(await fs.pathExists(outputFolder))) {
    await fs.ensureDir(outputFolder);
  }
  const versions = await fs.readdir(schemaFolder);
  for (const version of versions) {
    const convertedVersion = version.replace(".", "D").replace("v", "V");
    const schemaDir = path.join(schemaFolder, version);
    const schemaFile = "schema.json";
    const outputFileName = `ApiPluginManifest${convertedVersion}.ts`;
    const outputPath = path.join(outputFolder, outputFileName);
    await convert(schemaDir, schemaFile, outputPath);
  }
}

generateTeamsManifestTypeFiles();
generateDAManifestTypeFiles();
generatePluginManifestTypeFiles();
