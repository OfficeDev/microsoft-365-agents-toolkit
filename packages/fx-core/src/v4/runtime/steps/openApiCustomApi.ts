// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import openApiAssets from "./assets/openApiCustomApi.json";
import { renderFragment } from "../renderFragment";
import {
  AdaptiveCardGenerator,
  ConstantString,
  Utils,
  WarningResult,
  WarningType,
} from "@microsoft/m365-spec-parser";
import {
  AppPackageFolderName,
  ManifestTemplateFileName,
  TeamsManifestWrapper,
} from "@microsoft/teamsfx-api";
import fs from "fs-extra";
import { OpenAPIV3 } from "openapi-types";
import path from "path";
import * as util from "util";
import { featureFlagManager, FeatureFlags } from "../../../common/featureFlags";
import { getLocalizedString } from "../../../common/localizeUtils";
import { ProgrammingLanguage } from "../../../question/constants";

interface SpecOperation {
  pathUrl: string;
  method: string;
  operation: OpenAPIV3.OperationObject;
  auth: boolean;
}

const supportedLanguages = [ProgrammingLanguage.TS, ProgrammingLanguage.JS, ProgrammingLanguage.PY];

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function operationId(operation: OpenAPIV3.OperationObject): string {
  if (!operation.operationId) {
    throw new Error("The selected OpenAPI operation has no operationId.");
  }
  return operation.operationId;
}

function isOperationObject(value: unknown): value is OpenAPIV3.OperationObject {
  return isRecord(value);
}

function parseOperations(spec: OpenAPIV3.Document): SpecOperation[] {
  const operations: SpecOperation[] = [];
  for (const [pathUrl, pathItem] of Object.entries(spec.paths ?? {})) {
    if (!pathItem) {
      continue;
    }
    for (const method of ConstantString.AllOperationMethods) {
      const operation = isRecord(pathItem) ? pathItem[method] : undefined;
      if (isOperationObject(operation)) {
        operations.push({
          pathUrl,
          method,
          operation,
          auth: Utils.getAuthArray(operation.security, spec).length > 0,
        });
      }
    }
  }
  return operations;
}

async function updateInstructions(
  spec: OpenAPIV3.Document,
  language: string,
  appFolder: string
): Promise<void> {
  if (!supportedLanguages.includes(language as ProgrammingLanguage)) {
    return;
  }
  const emptyArgs = `{ "path": null, "body": null, "query": null }`;
  const description = spec.info.description ? `. ${spec.info.description}` : ".";
  const prompt = renderFragment(openApiAssets.instructions, { description, emptyArgs });
  await fs.writeFile(path.join(appFolder, "instructions.txt"), prompt, "utf8");
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

async function updateAdaptiveCards(
  operations: SpecOperation[],
  language: string,
  destinationPath: string
): Promise<WarningResult[]> {
  const warnings: WarningResult[] = [];
  if (!supportedLanguages.includes(language as ProgrammingLanguage)) {
    return warnings;
  }
  const cardsFolder = path.join(destinationPath, "src", "adaptiveCards");
  await fs.ensureDir(cardsFolder);

  for (const item of operations) {
    const id = operationId(item.operation);
    const name = id.replace(/[^a-zA-Z0-9]/g, "_");
    try {
      const [card, jsonPath, jsonData, generateWarnings] =
        AdaptiveCardGenerator.generateAdaptiveCard(item.operation, true, 5);
      if (
        jsonPath !== "$" &&
        isRecord(card) &&
        Array.isArray(card.body) &&
        isRecord(card.body[0]) &&
        card.body[0].$data
      ) {
        card.body[0].$data = `\${${jsonPath}}`;
      }
      await fs.writeFile(path.join(cardsFolder, `${name}.json`), JSON.stringify(card, null, 2));
      await fs.writeFile(
        path.join(cardsFolder, `${name}.data.json`),
        JSON.stringify(jsonData, null, 2)
      );
      for (const warning of generateWarnings) {
        warnings.push({
          type: WarningType.GenerateJsonDataFailed,
          content: util.format(
            "Failed to create the adaptive card mock data for API '%s': %s. Mitigation: Not required but you can manually add it to the adaptiveCards folder.",
            id,
            warning.content
          ),
          data: id,
        });
      }
    } catch (error) {
      warnings.push({
        type: WarningType.GenerateCardFailed,
        content: getLocalizedString(
          "core.copilotPlugin.scaffold.summary.warning.generate.ac.failed",
          id,
          errorMessage(error)
        ),
        data: id,
      });
    }
  }
  return warnings;
}

function filteredSchema(value: unknown): Record<string, unknown> {
  if (!isRecord(value)) {
    return {};
  }
  const result: Record<string, unknown> = {};
  if (typeof value.type === "string") {
    result.type = value.type;
  }
  if (typeof value.description === "string") {
    result.description = value.description;
  }
  if (value.type === "object" && isRecord(value.properties)) {
    const properties: Record<string, unknown> = {};
    for (const [name, property] of Object.entries(value.properties)) {
      properties[name] = filteredSchema(property);
    }
    result.properties = properties;
    if (Array.isArray(value.required)) {
      result.required = value.required.filter((name): name is string => typeof name === "string");
    }
  } else if (value.type === "array" && value.items !== undefined) {
    result.items = filteredSchema(value.items);
  }
  return result;
}

function functionParameters(operation: OpenAPIV3.OperationObject): Record<string, unknown> {
  const properties: Record<string, Record<string, unknown>> = {};
  const required: string[] = [];
  for (const parameter of operation.parameters ?? []) {
    if (
      !isRecord(parameter) ||
      typeof parameter.name !== "string" ||
      typeof parameter.in !== "string"
    ) {
      continue;
    }
    const parameterGroup = properties[parameter.in] ?? {
      type: "object",
      properties: {},
      required: [],
    };
    const groupProperties = parameterGroup.properties;
    const groupRequired = parameterGroup.required;
    if (!isRecord(groupProperties) || !Array.isArray(groupRequired)) {
      continue;
    }
    groupProperties[parameter.name] = {
      ...filteredSchema(parameter.schema),
      description: typeof parameter.description === "string" ? parameter.description : "",
    };
    if (parameter.required === true) {
      groupRequired.push(parameter.name);
      if (!required.includes(parameter.in)) {
        required.push(parameter.in);
      }
    }
    properties[parameter.in] = parameterGroup;
  }

  const requestBody = operation.requestBody;
  if (isRecord(requestBody) && isRecord(requestBody.content)) {
    const jsonContent = requestBody.content["application/json"];
    if (isRecord(jsonContent) && isRecord(jsonContent.schema)) {
      properties.body = {
        ...filteredSchema(jsonContent.schema),
        description: typeof requestBody.description === "string" ? requestBody.description : "",
      };
      if (requestBody.required === true) {
        required.push("body");
      }
    }
  }

  return { type: "object", properties, required };
}

async function updateFunctions(operations: SpecOperation[], appFolder: string): Promise<void> {
  const functions: Record<string, unknown> = {};
  for (const item of operations) {
    const id = operationId(item.operation);
    functions[id] = {
      name: id,
      description: item.operation.description ?? item.operation.summary,
      parameters: functionParameters(item.operation),
    };
  }
  await fs.writeFile(path.join(appFolder, "functions.json"), JSON.stringify(functions, null, 2));
}

const functionDefinitionCode = {
  javascript: openApiAssets.functionDefinitionCode.javascript.join("\n"),
  typescript: openApiAssets.functionDefinitionCode.typescript.join("\n"),
  python: openApiAssets.functionDefinitionCode.python.join("\n"),
};

const functionHandlerCode = {
  javascript: openApiAssets.functionHandlerCode.javascript.join("\n"),
  typescript: openApiAssets.functionHandlerCode.typescript.join("\n"),
  python: openApiAssets.functionHandlerCode.python.join("\n"),
};

function renderCode(template: string, item: SpecOperation): string {
  const id = operationId(item.operation);
  return template
    .replace(/{{operationId}}/g, id)
    .replace(/{{pathUrl}}/g, item.pathUrl)
    .replace(/{{method}}/g, item.method);
}

async function updateJavaScriptOrTypeScript(
  operations: SpecOperation[],
  language: string,
  destinationPath: string,
  openapiSpecFileName: string
): Promise<void> {
  const isJavaScript = language === ProgrammingLanguage.JS;
  const appFolder = path.join(destinationPath, "src", "app");
  const appPath = path.join(appFolder, isJavaScript ? "app.js" : "app.ts");
  const handlersPath = path.join(appFolder, isJavaScript ? "handlers.js" : "handlers.ts");
  const definitionTemplate = isJavaScript
    ? functionDefinitionCode.javascript
    : functionDefinitionCode.typescript;
  const handlerTemplate = isJavaScript
    ? functionHandlerCode.javascript
    : functionHandlerCode.typescript;
  const definitions = operations.map((item) => renderCode(definitionTemplate, item));
  const handlers = operations.map((item) => {
    const withAuth = item.auth
      ? handlerTemplate.replace(
          "// Add authentication configuration for the client",
          "addAuthConfig(client);"
        )
      : handlerTemplate.replace("// Add authentication configuration for the client", "");
    return renderCode(withAuth, item);
  });

  const app = await fs.readFile(appPath, "utf8");
  await fs.writeFile(
    appPath,
    app.replace("// Replace with function definition code", `${definitions.join("\n")};`)
  );
  const handlersFile = await fs.readFile(handlersPath, "utf8");
  await fs.writeFile(
    handlersPath,
    handlersFile
      .replace("{{OPENAPI_SPEC_PATH}}", openapiSpecFileName)
      .replace("// Replace with function handler code", handlers.join("\t\t\n"))
  );
}

async function updatePython(
  operations: SpecOperation[],
  destinationPath: string,
  openapiSpecFileName: string
): Promise<void> {
  const appPath = path.join(destinationPath, "src", "app.py");
  const handlersPath = path.join(destinationPath, "src", "handlers.py");
  const definitions = operations.map((item) => renderCode(functionDefinitionCode.python, item));
  const handlers = operations.map((item) => renderCode(functionHandlerCode.python, item));
  const operationIds = operations.map((item) => operationId(item.operation));

  const app = await fs.readFile(appPath, "utf8");
  await fs.writeFile(
    appPath,
    app
      .replace("// Replace with function definition code", `prompt${definitions.join("")}`)
      .replace("//Replace with functions to be imported", operationIds.join(", "))
  );
  const handlersFile = await fs.readFile(handlersPath, "utf8");
  await fs.writeFile(
    handlersPath,
    handlersFile
      .replace("{{OPENAPI_SPEC_PATH}}", openapiSpecFileName)
      .replace("// Replace with function handler code", handlers.join("\t\t\n"))
  );
}

async function updateCode(
  operations: SpecOperation[],
  language: string,
  destinationPath: string,
  openapiSpecFileName: string
): Promise<void> {
  if (language === ProgrammingLanguage.PY) {
    await updatePython(operations, destinationPath, openapiSpecFileName);
  } else if (language === ProgrammingLanguage.JS || language === ProgrammingLanguage.TS) {
    await updateJavaScriptOrTypeScript(operations, language, destinationPath, openapiSpecFileName);
  }
}

async function updatePromptSuggestions(
  operations: SpecOperation[],
  destinationPath: string
): Promise<void> {
  const commands = operations
    .map((item) => item.operation.summary ?? item.operation.description)
    .filter((description): description is string => description !== undefined)
    .slice(0, 10)
    .map((description) => ({
      title: description.slice(0, 32),
      description: description.slice(0, 128),
    }));
  const manifest = await TeamsManifestWrapper.read(
    path.join(destinationPath, AppPackageFolderName, ManifestTemplateFileName)
  );
  manifest.setFirstBotCommandSuggestions(
    commands,
    featureFlagManager.getBooleanValue(FeatureFlags.CEAEnabled)
  );
  await manifest.save();
}

/** Generate the post-render Teams AI files declared by the custom API scenario. */
export async function generateTeamsAiCustomApiFiles(
  spec: OpenAPIV3.Document,
  language: string,
  destinationPath: string,
  openapiSpecFileName: string
): Promise<WarningResult[]> {
  const appFolder = path.join(
    destinationPath,
    "src",
    language === ProgrammingLanguage.PY ? "" : "app"
  );
  await fs.ensureDir(appFolder);
  await updateInstructions(spec, language, appFolder);
  const operations = parseOperations(spec);
  const warnings = await updateAdaptiveCards(operations, language, destinationPath);
  await updateFunctions(operations, appFolder);
  await updateCode(operations, language, destinationPath, openapiSpecFileName);
  await updatePromptSuggestions(operations, destinationPath);
  return warnings;
}
