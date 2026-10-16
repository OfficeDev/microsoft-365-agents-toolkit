const {
  quicktype,
  InputData,
  JSONSchemaInput,
  FetchingJSONSchemaStore,
  tsFlowOptions,
  javaScriptOptions,
  TypeScriptTargetLanguage,
  TypeScriptRenderer,
  getOptionValues,
} = require("quicktype-core-jayzhang");
const fs = require("fs-extra");
const path = require("path");
const prettier = require("prettier");

const manifestProjectFolder = __dirname;

class ManifestTypeScriptRenderer extends TypeScriptRenderer {
  emitConvertModuleHelpers() {
    this.ensureBlankLine();
    const template = fs
      .readFileSync(path.join(manifestProjectFolder, "converter-runtime.txt"), "utf8")
      .replace(/\r\n?/g, "\n");
    for (const line of template.split("\n")) {
      this.emitLine(line);
    }
    this.emitTypeMap();
  }
}

class ManifestTypeScriptTargetLanguage extends TypeScriptTargetLanguage {
  makeRenderer(renderContext, untypedOptionValues) {
    return new ManifestTypeScriptRenderer(
      this,
      renderContext,
      getOptionValues(tsFlowOptions, untypedOptionValues)
    );
  }
}

async function quicktypeJSONSchema(targetLanguage, typeName, jsonSchemaString) {
  const schemaInput = new JSONSchemaInput(new FetchingJSONSchemaStore());
  await schemaInput.addSource({ name: typeName, schema: jsonSchemaString });

  const inputData = new InputData();
  inputData.addInput(schemaInput);

  tsFlowOptions.preferUnions.getValue = () => true;
  tsFlowOptions.preferConstValues.getValue = () => true;
  javaScriptOptions.jsonStringifySpaceNum.getValue = () => "4";

  const result = await quicktype({
    inputData,
    lang: targetLanguage === "typescript" ? new ManifestTypeScriptTargetLanguage() : targetLanguage,
    allPropertiesOptional: false,
    alphabetizeProperties: false,
  });
  if (targetLanguage !== "typescript") {
    return result;
  }
  const options = await prettier.resolveConfig(
    path.join(manifestProjectFolder, "src", "generated-types", "index.ts")
  );
  const source = await prettier.format(result.lines.join("\n"), {
    ...options,
    parser: "typescript",
  });
  return { ...result, lines: source.split("\n") };
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

module.exports = { quicktypeJSONSchema };

if (require.main === module) {
  Promise.all([
    generateTeamsManifestTypeFiles(),
    generateDAManifestTypeFiles(),
    generatePluginManifestTypeFiles(),
  ]).catch((error) => {
    console.error(error);
    process.exitCode = 1;
  });
}
