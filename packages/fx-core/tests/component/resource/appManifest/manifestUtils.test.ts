// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import {
  InputsWithProjectPath,
  ok,
  Platform,
  TeamsAppManifest,
  TeamsManifestV1D10Convert,
} from "@microsoft/teamsfx-api";
import * as chai from "chai";
import "mocha";
import "reflect-metadata";
import sinon from "sinon";
import * as uuid from "uuid";
import {
  manifestUtils,
  SharePointAppId,
} from "../../../../src/component/driver/teamsApp/utils/ManifestUtils";
import { JSONSyntaxError, MissingEnvironmentVariablesError } from "../../../../src/error/common";
import { newEnvInfoV3 } from "../../../helpers";
import fs from "fs-extra";
import { createContext, setTools } from "../../../../src/common/globalVars";
import { MockTools } from "../../../core/utils";
import { generateDriverContext } from "../../../../src/common/utils";
import { verify } from "crypto";

describe("getManifest V3", () => {
  const sandbox = sinon.createSandbox();
  let inputs: InputsWithProjectPath;
  let manifest: TeamsAppManifest;
  const manifestTemplate = `{
      "$schema": "https://developer.microsoft.com/en-us/json-schemas/teams/v1.14/MicrosoftTeams.schema.json",
      "manifestVersion": "1.14",
      "version": "1.0.0",
      "id": "{{state.fx-resource-appstudio.teamsAppId}}",
      "packageName": "com.microsoft.teams.extension",
      "developer": {
          "name": "Teams App, Inc.",
          "websiteUrl": "{{{state.fx-resource-frontend-hosting.endpoint}}}",
          "privacyUrl": "https://www.example.com/termofuse",
          "termsOfUseUrl": "https://www.example.com/privacy"
      },
      "icons": {
          "color": "{{config.manifest.icons.color}}",
          "outline": "{{config.manifest.icons.outline}}"
      },
      "name": {
          "short": "{{config.manifest.appName.short}}",
          "full": "{{config.manifest.appName.full}}"
      },
      "description": {
          "short": "{{config.manifest.description.short}}",
          "full": "{{config.manifest.description.full}}"
      },
      "accentColor": "#FFFFFF",
      "bots": [],
      "composeExtensions": [],
      "permissions": [
          "identity",
          "messageTeamMembers"
      ],
      "validDomains": [
          "{{state.fx-resource-frontend-hosting.domain}}"
      ],
      "webApplicationInfo": {
          "id": "{{state.fx-resource-aad-app-for-teams.clientId}}",
          "resource": "{{{state.fx-resource-aad-app-for-teams.applicationIdUris}}}"
      }
  }`;

  setTools(new MockTools());
  const context = generateDriverContext(createContext(), {
    platform: Platform.VSCode,
    projectPath: "",
  });
  beforeEach(async () => {
    inputs = {
      platform: Platform.VSCode,
      projectPath: ".",
    };
    manifest = JSON.parse(manifestTemplate) as TeamsAppManifest;
    sandbox.stub(manifestUtils, "readAppManifest").resolves(ok(manifest));
  });

  afterEach(async () => {
    sandbox.restore();
  });

  it("getManifestV3 MissingEnvironmentVariablesError", async () => {
    const envInfo = newEnvInfoV3();
    envInfo.envName = "dev";
    manifest.name.short = "${{MY_APP_NAME}}";
    sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    const res = await manifestUtils.getManifestV3("", context);
    chai.assert.isTrue(res.isErr() && res.error instanceof MissingEnvironmentVariablesError);
  });

  it("getManifestV3 teams app id resolved", async () => {
    const manifest = new TeamsAppManifest();
    manifest.id = uuid.v4();
    sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    const res = await manifestUtils.getManifestV3("", context);
    chai.assert.isTrue(res.isOk());
  });

  it("getOperationIds", async () => {
    const manifest = new TeamsAppManifest();
    manifest.composeExtensions = [
      {
        botId: uuid.v4(),
        commands: [
          {
            id: "GET /repairs",
            title: "List repairs",
          },
        ],
      },
    ];
    const ids = manifestUtils.getOperationIds(manifest);
    chai.assert.equal(ids.length, 1);
  });
});

describe("_readAppManifest", () => {
  const sandbox = sinon.createSandbox();
  afterEach(() => {
    sandbox.restore();
  });
  it("JSONSyntaxError", async () => {
    sandbox.stub(fs, "pathExists").resolves(true);
    sandbox.stub(fs, "readFile").resolves("invalid json" as any);
    const res = await manifestUtils._readAppManifest("invalid json");
    chai.assert.isTrue(res.isErr() && res.error instanceof JSONSyntaxError);
  });
});

describe("parseCommonProperties", () => {
  const sandbox = sinon.createSandbox();
  afterEach(() => {
    sandbox.restore();
  });
  it("configurableTab", async () => {
    const manifest: any = {
      id: uuid.v4(),
      version: "1.0.0",
      manifestVersion: "1.14",
      configurableTabs: [{}],
      webApplicationInfo: {
        id: SharePointAppId,
      },
      composeExtensions: [
        {
          composeExtensionType: "apiBased",
          authorization: {
            authType: "microsoftEntra",
          },
        },
      ],
      copilotAgents: {
        plugins: [
          {
            file: "xxx",
          },
        ],
        declarativeAgents: [{}],
      },
    };
    const res = manifestUtils.parseCommonProperties(manifest);
    chai.assert.isTrue(res.capabilities.includes("configurableTab"));
    chai.assert.isTrue(res.isSPFx);
    chai.assert.isTrue(res.isApiMeAAD);
    chai.assert.isTrue(res.capabilities.includes("copilotGpt"));
  });
});
