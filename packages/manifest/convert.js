const {
  quicktype,
  InputData,
  JSONSchemaInput,
  FetchingJSONSchemaStore,
  tsFlowOptions,
  javaScriptOptions,
} = require("quicktype-core-jayzhang");
const fs = require("fs-extra");
const path = require("path");

const manifestProjectFolder = __dirname;

async function quicktypeJSONSchema(targetLanguage, typeName, jsonSchemaString) {
  const schemaInput = new JSONSchemaInput(new FetchingJSONSchemaStore());
  await schemaInput.addSource({ name: typeName, schema: jsonSchemaString });

  const inputData = new InputData();
  inputData.addInput(schemaInput);

  tsFlowOptions.preferUnions.getValue = () => true;
  tsFlowOptions.preferConstValues.getValue = () => true;
  javaScriptOptions.jsonStringifySpaceNum.getValue = () => "4";

  return await quicktype({
    inputData,
    lang: targetLanguage,
    allPropertiesOptional: false,
    alphabetizeProperties: false,
  });
}

async function convert(srcFolder, srcFileName, typeName, destFilePath) {
  const jsonSchemaString = await fs.readFile(path.join(srcFolder, srcFileName), "utf-8");
  const { lines } = await quicktypeJSONSchema("typescript", typeName, jsonSchemaString);
  await fs.writeFile(destFilePath, lines.join("\n"), "utf-8");
  console.log(`Generated ${destFilePath} from ${srcFileName}`);
}

async function generateTeamsManifestTypeFiles() {
  const schemaFolder = path.join(manifestProjectFolder, "src", "json-schemas", "teams");
  const outputFolder = path.join(manifestProjectFolder, "src", "generated-types", "teams");

  if (!(await fs.pathExists(outputFolder))) {
    await fs.ensureDir(outputFolder);
  }

  const versions = await fs.readdir(schemaFolder);
  for (const version of versions) {
    const convertedVersion = version.replace(".", "D").replace("v", "V");
    const schemaDir = path.join(schemaFolder, version);
    const schemaFile = "MicrosoftTeams.schema.json";
    const outputFileName = `TeamsManifest${convertedVersion}.ts`;
    const outputTypeName = `TeamsManifest${convertedVersion}`;
    const outputPath = path.join(outputFolder, outputFileName);
    await convert(schemaDir, schemaFile, outputTypeName, outputPath);
  }
}

async function generateDAManifestTypeFiles() {
  const schemaFolder = path.join(
    manifestProjectFolder,
    "src",
    "json-schemas",
    "copilot",
    "declarative-agent"
  );
  const outputFolder = path.join(
    manifestProjectFolder,
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
    const outputTypeName = `DeclarativeAgentManifest${convertedVersion}`;
    const outputPath = path.join(outputFolder, outputFileName);
    await convert(schemaDir, schemaFile, outputTypeName, outputPath);
  }
}

async function generatePluginManifestTypeFiles() {
  const schemaFolder = path.join(manifestProjectFolder, "src", "json-schemas", "copilot", "plugin");
  const outputFolder = path.join(
    manifestProjectFolder,
    "src",
    "generated-types",
    "copilot",
    "plugin"
  );

  if (!(await fs.pathExists(outputFolder))) {
    await fs.ensureDir(outputFolder);
  }

  const versions = await fs.readdir(schemaFolder);
  for (const version of versions) {
    const convertedVersion = version.replace(".", "D").replace("v", "V");
    const schemaDir = path.join(schemaFolder, version);
    const schemaFile = "schema.json";
    const outputFileName = `ApiPluginManifest${convertedVersion}.ts`;
    const outputTypeName = `ApiPluginManifest${convertedVersion}`;
    const outputPath = path.join(outputFolder, outputFileName);
    await convert(schemaDir, schemaFile, outputTypeName, outputPath);
  }
}

generateTeamsManifestTypeFiles();
generateDAManifestTypeFiles();
generatePluginManifestTypeFiles();
