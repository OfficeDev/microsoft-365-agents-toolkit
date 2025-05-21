// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import path from "path";
import * as fs from "fs-extra";
import { updateContent } from "./commonUtils";

const jsonFileName = "azure.json";
const jsonParametersName = "azure.parameters.test.json";

export async function addJsonFileAndParamtersFile(
  projectPath: string
): Promise<void> {
  const srcFolder = path.join(__dirname, "../../asset/armDriver");
  const destFolder = path.join(projectPath, "infra");
  await fs.copyFile(
    path.join(srcFolder, jsonFileName),
    path.join(destFolder, jsonFileName)
  );
  await fs.copyFile(
    path.join(srcFolder, jsonParametersName),
    path.join(destFolder, jsonParametersName)
  );
}

export async function updateYml(projectPath: string): Promise<void> {
  const key = "deploymentName: Create-resources-for-tab";
  const replace = `
        - path: ./infra/azure.json 
          parameters: ./infra/azure.parameters.test.json
          deploymentName: test-json-format
    `;
  const ymlPath = path.join(projectPath, "m365agents.yml");
  let content = await fs.readFile(ymlPath, "utf-8");
  content = updateContent(content, key, replace);
  await fs.writeFileSync(ymlPath, content);
}
