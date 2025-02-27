// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Inputs } from "@microsoft/teamsfx-api";
import { DriverContext } from "../interface/commonArgs";
import { TOOLS } from "../../../common/globalVars";
import { MetadataV3 } from "../../../common/versionMetadata";
import path from "path";
import { parseDocument } from "yaml";
import fs from "fs-extra";

// Needs to validate the parameters outside of the function
export function loadStateFromEnv(
  outputEnvVarNames: Map<string, string>
): Record<string, string | undefined> {
  const result: Record<string, string | undefined> = {};
  for (const [propertyName, envVarName] of outputEnvVarNames) {
    result[propertyName] = process.env[envVarName];
  }
  return result;
}

// Needs to validate the parameters outside of the function
export function mapStateToEnv(
  state: Record<string, string>,
  outputEnvVarNames: Map<string, string>,
  excludedProperties?: string[]
): Map<string, string> {
  const result = new Map<string, string>();
  for (const [outputName, envVarName] of outputEnvVarNames) {
    if (!excludedProperties?.includes(outputName)) {
      result.set(envVarName, state[outputName]);
    }
  }
  return result;
}

export function createDriverContext(inputs: Inputs): DriverContext {
  const driverContext: DriverContext = {
    azureAccountProvider: TOOLS.tokenProvider.azureAccountProvider,
    m365TokenProvider: TOOLS.tokenProvider.m365TokenProvider,
    ui: TOOLS.ui,
    progressBar: undefined,
    logProvider: TOOLS.logProvider,
    telemetryReporter: TOOLS.telemetryReporter!,
    projectPath: inputs.projectPath!,
    platform: inputs.platform,
  };
  return driverContext;
}

export async function updateVersionForTeamsAppYamlFile(projectPath: string): Promise<void> {
  const allPossilbeYamlFileNames = [
    MetadataV3.localConfigFile,
    MetadataV3.configFile,
    MetadataV3.testToolConfigFile,
  ];
  for (const yamlFileName of allPossilbeYamlFileNames) {
    const ymlPath = path.join(projectPath, yamlFileName);
    if (await fs.pathExists(ymlPath)) {
      const ymlContent = await fs.readFile(ymlPath, "utf-8");
      const document = parseDocument(ymlContent);
      const version = document.get("version") as string;
      if (version <= "v1.7") {
        if (version <= "v1.6") {
          convertOutputJsonPathToOutputFolder(document);
        }

        if (version === "v1.3") {
          renameClientSecret(document);
        }

        document.set("version", "v1.8");
        const docContent = document.toString();
        // yaml-language-server can be like https://aka.ms/teams-toolkit/1.0.0/yaml.schema.json and https://aka.ms/teams-toolkit/v1.2/yaml.schema.json
        const updatedContent = docContent.replace(
          /(yaml-language-server:\s*\$schema=https:\/\/aka\.ms\/teams-toolkit\/)(v?\d+\.\d+(?:\.\d+)?)(\/yaml\.schema\.json)/,
          "$1v1.8$3"
        );
        await fs.writeFile(ymlPath, updatedContent, "utf8");
      }
    }
  }
}

function processSectionsByUse(
  document: any,
  targetUseValue: string,
  processor: (withNode: any) => void
): void {
  const sections = ["provision", "publish"];
  sections.forEach((sectionKey) => {
    const section = document.get(sectionKey, true);
    if (section && Array.isArray(section.items)) {
      section.items.forEach((item: any) => {
        if (item && item.get("uses", true)?.value === targetUseValue) {
          const withNode = item.get("with", true);
          if (withNode) {
            processor(withNode);
          }
        }
      });
    }
  });
}

function convertOutputJsonPathToOutputFolder(document: any): void {
  processSectionsByUse(document, "teamsApp/zipAppPackage", (withNode) => {
    const outputJsonPath = withNode.get("outputJsonPath", true)?.value;
    if (typeof outputJsonPath === "string") {
      withNode.set("outputFolder", path.dirname(outputJsonPath));
      withNode.delete("outputJsonPath");
    }
  });
}

function renameClientSecret(document: any): void {
  processSectionsByUse(document, "apiKey/register", (withNode) => {
    const secretValue = withNode.get("clientSecret", true)?.value;
    if (typeof secretValue === "string") {
      withNode.set("primaryClientSecret", secretValue);
      withNode.delete("clientSecret");
    }
  });
}
