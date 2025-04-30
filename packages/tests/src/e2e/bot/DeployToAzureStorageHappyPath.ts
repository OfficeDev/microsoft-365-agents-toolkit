// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
/**
 * @author Siglud <siglud@gmail.com>
 */

import { describe } from "mocha";
import {
  cleanUp,
  createResourceGroup,
  execAsyncWithRetry,
  getTestFolder,
  getUniqueAppName,
} from "../commonUtils";
import * as path from "path";
import * as fs from "fs";
import { getUuid } from "../../commonlib";
import { expect } from "chai";
import { environmentNameManager } from "@microsoft/teamsfx-core";
import { Executor } from "../../utils/executor";

describe("Provision and deploy a Azure Storage", async function () {
  // create a project with Azure Storage
  const testFolder = getTestFolder();
  const appName = getUniqueAppName();
  const projectPath = path.resolve(testFolder, appName);
  const projectId = getUuid();
  const envName = environmentNameManager.getDefaultEnvName();
  const rgName = appName + "-rg";
  const env = Object.assign({}, process.env);

  it("should be provision without problem", async () => {
    fs.mkdirSync(projectPath, { recursive: true });
    // write teamsapp.yml
    fs.writeFileSync(
      path.join(projectPath, "teamsapp.yml"),
      `
# yaml-language-server: $schema=https://aka.ms/teams-toolkit/v1.7/yaml.schema.json
version: v1.7

environmentFolderPath: ./env

provision:
  - uses: arm/deploy
    with:
      subscriptionId: \${{AZURE_SUBSCRIPTION_ID}}
      resourceGroupName: \${{AZURE_RESOURCE_GROUP_NAME}}
      templates:
        - path: ./infra/azure.bicep
          parameters: ./infra/azure.parameters.json
          deploymentName: Create-resources-for-tab
      bicepCliVersion: v0.9.1
  - uses: azureStorage/enableStaticWebsite
    with:
      storageResourceId: \${{TAB_AZURE_STORAGE_RESOURCE_ID}}
      indexPage: index.html
      errorPage: error.html

deploy:
  - uses: azureStorage/deploy
    with:
      artifactFolder: build
      resourceId: \${{TAB_AZURE_STORAGE_RESOURCE_ID}}
projectId: ${projectId}`,
      { encoding: "utf-8", flag: "w" }
    );
    // mkdir for infra
    fs.mkdirSync(path.join(projectPath, "infra"), { recursive: true });
    // mkdir for env
    fs.mkdirSync(path.join(projectPath, "env"), { recursive: true });
    // write azure.bicep
    fs.writeFileSync(
      path.join(projectPath, "infra", "azure.bicep"),
      `
param resourceBaseName string

resource storage 'Microsoft.Storage/storageAccounts@2021-06-01' = {
  kind: 'StorageV2'
  location: resourceGroup().location
  name: resourceBaseName
  properties: {
    supportsHttpsTrafficOnly: true
  }
  sku: {
    name: 'Standard_LRS'
  }
}

output TAB_AZURE_STORAGE_RESOURCE_ID string = storage.id
output TAB_DOMAIN string = storage.properties.primaryEndpoints.web`,
      { encoding: "utf-8", flag: "w" }
    );
    // write azure.parameters.json
    fs.writeFileSync(
      path.join(projectPath, "infra", "azure.parameters.json"),
      `{
  "$schema": "https://schema.management.azure.com/schemas/2015-01-01/deploymentParameters.json#",
  "contentVersion": "1.0.0.0",
  "parameters": {
    "resourceBaseName": {
      "value": "hello0world\${{RESOURCE_SUFFIX}}"
    }
  }
}`,
      { encoding: "utf-8", flag: "w" }
    );
    // write .env.dev
    const suffix = getUuid().split("-");
    fs.writeFileSync(
      path.join(projectPath, "env", ".env.dev"),
      "TEAMSFX_ENV=dev\n" +
        "APP_NAME_SUFFIX=dev\n" +
        `RESOURCE_SUFFIX=${suffix[suffix.length - 1]}\n` +
        `AZURE_RESOURCE_GROUP_NAME=${rgName}`,
      { encoding: "utf-8", flag: "w" }
    );
    // write build and build/index.html
    fs.mkdirSync(path.join(projectPath, "build"), { recursive: true });
    fs.writeFileSync(
      path.join(projectPath, "build", "index.html"),
      "<h1>Hello World</h1>",
      { encoding: "utf-8", flag: "w" }
    );

    // run provision
    const result = await createResourceGroup(rgName, "westus");
    expect(result).to.be.true;
    process.env["AZURE_RESOURCE_GROUP_NAME"] = rgName;
    const { success } = await Executor.provision(projectPath, envName);
    expect(success).to.be.true;
    console.log(`[Successfully] provision for ${projectPath}`);

    // deploy
    const cmdStr = "atk deploy";
    await execAsyncWithRetry(cmdStr, {
      cwd: projectPath,
      env: env,
      timeout: 0,
    });

    console.log(`[Successfully] deploy for ${projectPath}`);

    // request to index.html to check the deployment result
    const file = fs.readFileSync(
      path.join(projectPath, "env", ".env.dev"),
      "utf-8"
    );
    const line = file.split("\n").filter((line) => {
      return line.startsWith("TAB_DOMAIN");
    });
    const domain = line[0].split("=")[1];
    const response = await fetch(domain);
    expect(response.status).to.equal(200);
  });

  this.afterEach(async function () {
    console.log(`[Successfully] start to clean up for ${projectPath}`);
    await cleanUp(appName, projectPath, false, true, false, envName, undefined);
  });
});
