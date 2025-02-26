/**
 * @author Yuqi Zhou <yuqzho@microsoft.com>
 */
import { err, Inputs, ok, Platform, TeamsAppManifest, UserError } from "@microsoft/teamsfx-api";
import chai from "chai";
import fs from "fs-extra";
import { merge } from "lodash";
import "mocha";
import path from "path";
import * as sinon from "sinon";
import { createContext, setTools } from "../../src/common/globalVars";
import {
  adjustScopeBasedOnVersion,
  developerPortalScaffoldUtils,
} from "../../src/component/developerPortalScaffoldUtils";
import * as appStudio from "../../src/component/driver/teamsApp/appStudio";
import {
  BOTS_TPL_V3,
  COMPOSE_EXTENSIONS_TPL_V3,
  DEFAULT_DESCRIPTION,
  DEFAULT_DEVELOPER,
} from "../../src/component/driver/teamsApp/constants";
import { AppDefinition } from "../../src/component/driver/teamsApp/interfaces/appdefinitions/appDefinition";
import { Bot } from "../../src/component/driver/teamsApp/interfaces/appdefinitions/bot";
import { ConfigurableTab } from "../../src/component/driver/teamsApp/interfaces/appdefinitions/configurableTab";
import { MessagingExtension } from "../../src/component/driver/teamsApp/interfaces/appdefinitions/messagingExtension";
import { StaticTab } from "../../src/component/driver/teamsApp/interfaces/appdefinitions/staticTab";
import { manifestUtils } from "../../src/component/driver/teamsApp/utils/ManifestUtils";
import { CommandScope, MeetingsContext } from "../../src/component/driver/teamsApp/utils/utils";
import { DotenvOutput, envUtil } from "../../src/component/utils/envUtil";
import { CapabilityOptions, QuestionNames } from "../../src/question/constants";
import { getProjectTypeAndCapability } from "../../src/question/create";
import { MockedAzureAccountProvider, MockedM365Provider, MockTools } from "../core/utils";
import { InputValidationError } from "../../src/error";

describe("developPortalScaffoldUtils", () => {
  setTools(new MockTools());
  describe("updateFilesForTdp", () => {
    const sandbox = sinon.createSandbox();
    class MockedWriteStream {
      write(): boolean {
        return true;
      }
      end(): boolean {
        return true;
      }
    }

    afterEach(() => {
      sandbox.restore();
    });
    it("missing project path", async () => {
      const ctx = createContext();
      ctx.tokenProvider = {
        m365TokenProvider: new MockedM365Provider(),
        azureAccountProvider: new MockedAzureAccountProvider(),
      };
      const appDefinition: AppDefinition = {
        appId: "mock-app-id",
        teamsAppId: "mock-app-id",
      };
      const inputs: Inputs = { platform: Platform.VSCode };

      const res = await developerPortalScaffoldUtils.updateFilesForTdp(ctx, appDefinition, inputs);
      chai.assert.isTrue(res.isErr());
      if (res.isErr()) {
        chai.assert.isTrue(res.error instanceof InputValidationError);
      }
    });

    it("missing token provider", async () => {
      const ctx = createContext();
      const appDefinition: AppDefinition = {
        appId: "mock-app-id",
        teamsAppId: "mock-app-id",
      };
      const inputs: Inputs = { platform: Platform.VSCode };
      ctx.tokenProvider = undefined;
      ctx.projectPath = "project-path";

      const res = await developerPortalScaffoldUtils.updateFilesForTdp(ctx, appDefinition, inputs);
      chai.assert.isTrue(res.isErr());
      if (res.isErr()) {
        chai.assert.isTrue(res.error instanceof InputValidationError);
      }
    });

    it("get App package error", async () => {
      const ctx = createContext();
      ctx.tokenProvider = {
        m365TokenProvider: new MockedM365Provider(),
        azureAccountProvider: new MockedAzureAccountProvider(),
      };
      ctx.projectPath = "project-path";
      const appDefinition: AppDefinition = {
        appId: "mock-app-id",
        teamsAppId: "mock-app-id",
      };
      const inputs: Inputs = { platform: Platform.VSCode, projectPath: "project-path" };

      sandbox
        .stub(appStudio, "getAppPackage")
        .resolves(err(new UserError("source", "getAppPackage", "msg", "msg")));
      const res = await developerPortalScaffoldUtils.updateFilesForTdp(ctx, appDefinition, inputs);

      chai.assert.isTrue(res.isErr());
      if (res.isErr()) {
        chai.assert.equal(res.error.name, "getAppPackage");
      }
    });

    it("missing manifest error", async () => {
      const ctx = createContext();
      ctx.tokenProvider = {
        m365TokenProvider: new MockedM365Provider(),
        azureAccountProvider: new MockedAzureAccountProvider(),
      };
      ctx.projectPath = "project-path";
      const appDefinition: AppDefinition = {
        appId: "mock-app-id",
        teamsAppId: "mock-app-id",
      };
      const inputs: Inputs = { platform: Platform.VSCode, projectPath: "project-path" };

      sandbox.stub(appStudio, "getAppPackage").resolves(ok({}));
      const res = await developerPortalScaffoldUtils.updateFilesForTdp(ctx, appDefinition, inputs);

      chai.assert.isTrue(res.isErr());
      if (res.isErr()) {
        chai.assert.equal(res.error.name, "CouldNotFoundManifest");
      }
    });

    it("missing manifest.json from template", async () => {
      const ctx = createContext();
      ctx.tokenProvider = {
        m365TokenProvider: new MockedM365Provider(),
        azureAccountProvider: new MockedAzureAccountProvider(),
      };
      ctx.projectPath = "project-path";
      const appDefinition: AppDefinition = {
        appId: "mock-app-id",
        teamsAppId: "mock-app-id",
      };
      const inputs: Inputs = { platform: Platform.VSCode };

      const manifest: TeamsAppManifest = {
        manifestVersion: "version",
        id: "mock-app-id",
        name: { short: "short-name" },
        description: { short: "", full: "" },
        version: "version",
        icons: { outline: "outline.png", color: "color.png" },
        accentColor: "#ffffff",
        developer: {
          privacyUrl: "",
          websiteUrl: "",
          termsOfUseUrl: "",
          name: "developer-name",
        },
        staticTabs: [
          {
            name: "name0",
            entityId: "index0",
            scopes: ["personal"],
            contentUrl: "contentUrl0",
            websiteUrl: "websiteUrl0",
          },
          {
            name: "name1",
            entityId: "index1",
            scopes: ["personal"],
            contentUrl: "contentUrl1",
            websiteUrl: "websiteUrl1",
          },
        ],
      };
      sandbox.stub(appStudio, "getAppPackage").resolves(
        ok({
          manifest: Buffer.from(JSON.stringify(manifest)),
          icons: { color: Buffer.from(""), outline: Buffer.from("") },
          languages: { zh: Buffer.from(JSON.stringify({})) },
        })
      );
      sandbox
        .stub(manifestUtils, "_readAppManifest")
        .resolves(ok(undefined as unknown as TeamsAppManifest));
      const res = await developerPortalScaffoldUtils.updateFilesForTdp(ctx, appDefinition, inputs);

      chai.assert.isTrue(res.isErr());
      if (res.isErr()) {
        chai.assert.isTrue(res.error instanceof InputValidationError);
      }
    });

    it("update files successfully", async () => {
      const ctx = createContext();
      ctx.tokenProvider = {
        m365TokenProvider: new MockedM365Provider(),
        azureAccountProvider: new MockedAzureAccountProvider(),
      };
      ctx.projectPath = "project-path";
      const appDefinition: AppDefinition = {
        appId: "mock-app-id",
        teamsAppId: "mock-app-id",
        staticTabs: [
          {
            objectId: "objId",
            entityId: "entityId",
            name: "tab",
            contentUrl: "https://url",
            websiteUrl: "https:/url",
            scopes: [],
            context: [],
          },
        ],
      };
      const inputs: Inputs = {
        platform: Platform.VSCode,
        [QuestionNames.ReplaceWebsiteUrl]: ["name0"],
        [QuestionNames.ReplaceContentUrl]: ["name1"],
        projectPath: "project-path",
      };
      const manifest: TeamsAppManifest = {
        manifestVersion: "version",
        id: "mock-app-id",
        name: { short: "short-name" },
        description: { short: "", full: "" },
        version: "version",
        icons: { outline: "outline.png", color: "color.png" },
        accentColor: "#ffffff",
        developer: {
          privacyUrl: "",
          websiteUrl: "",
          termsOfUseUrl: "",
          name: "",
        },
        staticTabs: [
          {
            name: "name0",
            entityId: "index0",
            scopes: ["personal"],
            contentUrl: "contentUrl0",
            websiteUrl: "websiteUrl0",
          },
          {
            name: "name1",
            entityId: "index1",
            scopes: ["personal"],
            contentUrl: "contentUrl1",
            websiteUrl: "websiteUrl1",
          },
        ],
      };

      const manifestTemplate: TeamsAppManifest = {
        manifestVersion: "version",
        id: "mock-app-id",
        name: { short: "short-name" },
        description: { short: "", full: "" },
        version: "version",
        icons: { outline: "outline.png", color: "color.png" },
        accentColor: "#ffffff",
        developer: {
          privacyUrl: "",
          websiteUrl: "",
          termsOfUseUrl: "",
          name: "",
        },
        staticTabs: [
          {
            name: "name0",
            entityId: "index0",
            scopes: ["personal"],
            contentUrl: "localhost/content",
            websiteUrl: "localhost/website",
          },
        ],
      };

      let updateManifest = false;
      let updateLanguage = false;
      let updateColor = false;
      let updateOutline = false;
      let updatedManifestData = "";

      sandbox.stub(appStudio, "getAppPackage").resolves(
        ok({
          manifest: Buffer.from(JSON.stringify(manifest)),
          icons: { color: Buffer.from(""), outline: Buffer.from("") },
          languages: { zh: Buffer.from(JSON.stringify({})) },
        })
      );
      sandbox.stub(fs, "writeFile").callsFake((file: number | fs.PathLike, data: any) => {
        if (file === path.join(ctx.projectPath!, "appPackage", "color.png")) {
          updateColor = true;
        } else if (file === path.join(ctx.projectPath!, "appPackage", "outline.png")) {
          updateOutline = true;
        } else if (file === path.join(ctx.projectPath!, "appPackage", "zh.json")) {
          updateLanguage = true;
        } else if (file === path.join(ctx.projectPath!, "appPackage", "manifest.json")) {
          updateManifest = true;
          updatedManifestData = data;
        } else {
          throw new Error("not support " + file);
        }
      });

      const originalEnvs: DotenvOutput = {};
      sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(manifestTemplate));
      sandbox
        .stub(envUtil, "writeEnv")
        .callsFake(async (projectPath: string, env: string, envs: DotenvOutput) => {
          merge(originalEnvs, envs);
          return ok(undefined);
        });
      const res = await developerPortalScaffoldUtils.updateFilesForTdp(ctx, appDefinition, inputs);

      chai.assert.isTrue(res.isOk());
      chai.assert.isTrue(updateManifest);
      chai.assert.isTrue(updateColor);
      chai.assert.isTrue(updateOutline);
      chai.assert.isTrue(updateLanguage);
      const updatedManifest = JSON.parse(updatedManifestData) as TeamsAppManifest;
      chai.assert.equal(updatedManifest.id, "${{TEAMS_APP_ID}}");
      chai.assert.equal(updatedManifest.staticTabs![0].contentUrl, "contentUrl0");
      chai.assert.equal(updatedManifest.staticTabs![0].websiteUrl, "localhost/website");
      chai.assert.equal(updatedManifest.staticTabs![1].websiteUrl, "websiteUrl1");
      chai.assert.equal(updatedManifest.staticTabs![1].contentUrl, "localhost/content");
      chai.assert.equal(updatedManifest.developer.privacyUrl, DEFAULT_DEVELOPER.privacyUrl);
      chai.assert.equal(updatedManifest.developer.name, DEFAULT_DEVELOPER.name);
      chai.assert.equal(updatedManifest.developer.termsOfUseUrl, DEFAULT_DEVELOPER.termsOfUseUrl);
      chai.assert.equal(updatedManifest.developer.websiteUrl, DEFAULT_DEVELOPER.websiteUrl);
      chai.assert.equal(updatedManifest.description.short, DEFAULT_DESCRIPTION.short);
      chai.assert.equal(updatedManifest.description.full, DEFAULT_DESCRIPTION.full);
      chai.assert.isTrue(updatedManifest.validDomains?.includes("${{TAB_DOMAIN}}"));
      chai.assert.equal(originalEnvs.TEAMS_APP_ID, "mock-app-id");
    });

    it("update files successfully but keep url", async () => {
      const ctx = createContext();
      ctx.tokenProvider = {
        m365TokenProvider: new MockedM365Provider(),
        azureAccountProvider: new MockedAzureAccountProvider(),
      };
      ctx.projectPath = "project-path";
      const appDefinition: AppDefinition = {
        appId: "mock-app-id",
        teamsAppId: "mock-app-id",
        staticTabs: [
          {
            objectId: "objId",
            entityId: "entityId",
            name: "tab",
            contentUrl: "https://url",
            websiteUrl: "https:/url",
            scopes: [],
            context: [],
          },
        ],
      };
      const inputs: Inputs = {
        platform: Platform.VSCode,
        [QuestionNames.ReplaceWebsiteUrl]: [],
        [QuestionNames.ReplaceContentUrl]: [],
        projectPath: "project-path",
      };
      const manifest: TeamsAppManifest = {
        manifestVersion: "version",
        id: "mock-app-id",
        name: { short: "short-name" },
        description: { short: "short", full: "full" },
        version: "version",
        icons: { outline: "outline.png", color: "color.png" },
        accentColor: "#ffffff",
        developer: {
          privacyUrl: "privacyUrl",
          websiteUrl: "websiteUrl",
          termsOfUseUrl: "termsOfUseUrl",
          name: "developer-name",
        },
        staticTabs: [
          {
            name: "name0",
            entityId: "index0",
            scopes: ["personal"],
            contentUrl: "contentUrl0",
            websiteUrl: "websiteUrl0",
          },
          {
            name: "name1",
            entityId: "index1",
            scopes: ["personal"],
            contentUrl: "contentUrl1",
            websiteUrl: "websiteUrl1",
          },
        ],
      };

      let updateManifest = false;
      let updateLanguage = false;
      let updateColor = false;
      let updateOutline = false;
      let updatedManifestData = "";
      sandbox.stub(appStudio, "getAppPackage").resolves(
        ok({
          manifest: Buffer.from(JSON.stringify(manifest)),
          icons: { color: Buffer.from(""), outline: Buffer.from("") },
          languages: { zh: Buffer.from(JSON.stringify({})) },
        })
      );
      sandbox.stub(fs, "writeFile").callsFake((file: number | fs.PathLike, data: any) => {
        if (file === path.join(ctx.projectPath!, "appPackage", "color.png")) {
          updateColor = true;
        } else if (file === path.join(ctx.projectPath!, "appPackage", "outline.png")) {
          updateOutline = true;
        } else if (file === path.join(ctx.projectPath!, "appPackage", "zh.json")) {
          updateLanguage = true;
        } else if (file === path.join(ctx.projectPath!, "appPackage", "manifest.json")) {
          updateManifest = true;
          updatedManifestData = data;
        } else {
          throw new Error("not support " + file);
        }
      });

      sandbox.stub(envUtil, "writeEnv").resolves(ok(undefined));
      sandbox.stub(manifestUtils, "_readAppManifest").resolves(
        ok({
          manifestVersion: "version",
          id: "mock-app-id",
          name: { short: "short-name" },
          description: { short: "", full: "" },
          version: "version",
          icons: { outline: "outline.png", color: "color.png" },
          accentColor: "#ffffff",
          developer: {
            privacyUrl: "",
            websiteUrl: "",
            termsOfUseUrl: "",
            name: "developer-name",
          },
          staticTabs: [
            {
              name: "name0",
              entityId: "index0",
              scopes: ["personal"],
              contentUrl: "contentUrlnew0",
              websiteUrl: "websiteUrlnew0",
            },
          ],
        })
      );

      const res = await developerPortalScaffoldUtils.updateFilesForTdp(ctx, appDefinition, inputs);

      chai.assert.isTrue(res.isOk());
      chai.assert.isTrue(updateManifest);
      chai.assert.isTrue(updateColor);
      chai.assert.isTrue(updateOutline);
      chai.assert.isTrue(updateLanguage);
      const updatedManifest = JSON.parse(updatedManifestData) as TeamsAppManifest;
      chai.assert.equal(updatedManifest.id, "${{TEAMS_APP_ID}}");
      chai.assert.equal(updatedManifest.staticTabs![0].contentUrl, "contentUrl0");
      chai.assert.equal(updatedManifest.staticTabs![0].websiteUrl, "websiteUrl0");
      chai.assert.equal(updatedManifest.staticTabs![1].websiteUrl, "websiteUrl1");
      chai.assert.equal(updatedManifest.staticTabs![1].contentUrl, "contentUrl1");
      chai.assert.equal(updatedManifest.developer.privacyUrl, "privacyUrl");
      chai.assert.equal(updatedManifest.developer.termsOfUseUrl, "termsOfUseUrl");
      chai.assert.equal(updatedManifest.developer.websiteUrl, "websiteUrl");
      chai.assert.equal(updatedManifest.description.short, "short");
      chai.assert.equal(updatedManifest.description.full, "full");
      chai.assert.equal(updatedManifest.validDomains, undefined);
    });

    it("update bot id only", async () => {
      const ctx = createContext();
      ctx.tokenProvider = {
        m365TokenProvider: new MockedM365Provider(),
        azureAccountProvider: new MockedAzureAccountProvider(),
      };
      ctx.projectPath = "project-path";
      const appDefinition: AppDefinition = {
        appId: "mock-app-id",
        teamsAppId: "mock-app-id",
        staticTabs: [
          {
            objectId: "objId",
            entityId: "entityId",
            name: "tab",
            contentUrl: "https://url",
            websiteUrl: "https:/url",
            scopes: [],
            context: [],
          },
        ],
      };
      const inputs: Inputs = {
        platform: Platform.VSCode,
        [QuestionNames.ReplaceBotIds]: ["bot"],
        projectPath: "project-path",
      };
      const manifest: TeamsAppManifest = {
        manifestVersion: "version",
        id: "mock-app-id",
        name: { short: "short-name" },
        description: { short: "", full: "" },
        version: "version",
        icons: { outline: "outline.png", color: "color.png" },
        accentColor: "#ffffff",
        developer: {
          privacyUrl: "",
          websiteUrl: "",
          termsOfUseUrl: "",
          name: "developer-name",
        },
        bots: [
          {
            botId: "botId0",
            scopes: ["personal"],
            commandLists: [],
          },
        ],
        composeExtensions: [
          {
            botId: "botId1",
            commands: [],
          },
        ],
      };

      const existingManifest: TeamsAppManifest = {
        manifestVersion: "version",
        id: "mock-app-id",
        name: { short: "short-name" },
        description: { short: "", full: "" },
        version: "version",
        icons: { outline: "outline.png", color: "color.png" },
        accentColor: "#ffffff",
        developer: {
          privacyUrl: "",
          websiteUrl: "",
          termsOfUseUrl: "",
          name: "developer-name",
        },
        bots: [],
        validDomains: ["valid-domain"],
      };

      let updateManifest = false;
      let updateLanguage = false;
      let updateColor = false;
      let updateOutline = false;
      let updatedManifestData = "";
      sandbox.stub(appStudio, "getAppPackage").resolves(
        ok({
          manifest: Buffer.from(JSON.stringify(manifest)),
          icons: { color: Buffer.from(""), outline: Buffer.from("") },
          languages: { zh: Buffer.from(JSON.stringify({})) },
        })
      );
      sandbox.stub(fs, "writeFile").callsFake((file: number | fs.PathLike, data: any) => {
        if (file === path.join(ctx.projectPath!, "appPackage", "color.png")) {
          updateColor = true;
        } else if (file === path.join(ctx.projectPath!, "appPackage", "outline.png")) {
          updateOutline = true;
        } else if (file === path.join(ctx.projectPath!, "appPackage", "zh.json")) {
          updateLanguage = true;
        } else if (file === path.join(ctx.projectPath!, "appPackage", "manifest.json")) {
          updateManifest = true;
          updatedManifestData = data;
        } else {
          throw new Error("not support " + file);
        }
      });

      sandbox.stub(envUtil, "writeEnv").resolves(ok(undefined));
      sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(existingManifest));
      const res = await developerPortalScaffoldUtils.updateFilesForTdp(ctx, appDefinition, inputs);

      chai.assert.isTrue(res.isOk());
      chai.assert.isTrue(updateManifest);
      chai.assert.isTrue(updateColor);
      chai.assert.isTrue(updateOutline);
      chai.assert.isTrue(updateLanguage);
      const updatedManifest = JSON.parse(updatedManifestData) as TeamsAppManifest;
      chai.assert.equal(updatedManifest.id, "${{TEAMS_APP_ID}}");
      const expectedBots = BOTS_TPL_V3;
      expectedBots[0].botId = "${{BOT_ID}}";
      chai.assert.deepEqual(updatedManifest.bots![0], expectedBots[0]);
      chai.assert.deepEqual(updatedManifest.composeExtensions![0], manifest.composeExtensions![0]);
      chai.assert.equal(updatedManifest.developer.privacyUrl, DEFAULT_DEVELOPER.privacyUrl);
      chai.assert.equal(updatedManifest.developer.termsOfUseUrl, DEFAULT_DEVELOPER.termsOfUseUrl);
      chai.assert.equal(updatedManifest.developer.websiteUrl, DEFAULT_DEVELOPER.websiteUrl);
      chai.assert.isUndefined(updatedManifest.validDomains);
    });

    it("update bot id of message extension only", async () => {
      const ctx = createContext();
      ctx.tokenProvider = {
        m365TokenProvider: new MockedM365Provider(),
        azureAccountProvider: new MockedAzureAccountProvider(),
      };
      ctx.projectPath = "project-path";
      const appDefinition: AppDefinition = {
        appId: "mock-app-id",
        teamsAppId: "mock-app-id",
        staticTabs: [
          {
            objectId: "objId",
            entityId: "entityId",
            name: "tab",
            contentUrl: "https://url",
            websiteUrl: "https:/url",
            scopes: [],
            context: [],
          },
        ],
      };
      const inputs: Inputs = {
        platform: Platform.VSCode,
        [QuestionNames.ReplaceBotIds]: ["messageExtension"],
        projectPath: "project-path",
      };
      const manifest: TeamsAppManifest = {
        manifestVersion: "version",
        id: "mock-app-id",
        name: { short: "short-name" },
        description: { short: "", full: "" },
        version: "version",
        icons: { outline: "outline.png", color: "color.png" },
        accentColor: "#ffffff",
        developer: {
          privacyUrl: "",
          websiteUrl: "",
          termsOfUseUrl: "",
          name: "developer-name",
        },
        bots: [
          {
            botId: "botId0",
            scopes: ["personal"],
            commandLists: [],
          },
        ],
        composeExtensions: [
          {
            botId: "botId1",
            commands: [],
          },
        ],
        validDomains: [],
      };

      const existingManifest: TeamsAppManifest = {
        manifestVersion: "version",
        id: "mock-app-id",
        name: { short: "short-name" },
        description: { short: "", full: "" },
        version: "version",
        icons: { outline: "outline.png", color: "color.png" },
        accentColor: "#ffffff",
        developer: {
          privacyUrl: "",
          websiteUrl: "",
          termsOfUseUrl: "",
          name: "developer-name",
        },
        bots: [],
        composeExtensions: [
          {
            botId: "botId1",
            commands: [
              {
                id: "commandId",
                title: "commandTitle",
              },
            ],
          },
        ],
        validDomains: ["valid-domain"],
      };

      let updateManifest = false;
      let updateLanguage = false;
      let updateColor = false;
      let updateOutline = false;
      let updatedManifestData = "";
      sandbox.stub(appStudio, "getAppPackage").resolves(
        ok({
          manifest: Buffer.from(JSON.stringify(manifest)),
          icons: { color: Buffer.from(""), outline: Buffer.from("") },
          languages: { zh: Buffer.from(JSON.stringify({})) },
        })
      );
      sandbox.stub(fs, "writeFile").callsFake((file: number | fs.PathLike, data: any) => {
        if (file === path.join(ctx.projectPath!, "appPackage", "color.png")) {
          updateColor = true;
        } else if (file === path.join(ctx.projectPath!, "appPackage", "outline.png")) {
          updateOutline = true;
        } else if (file === path.join(ctx.projectPath!, "appPackage", "zh.json")) {
          updateLanguage = true;
        } else if (file === path.join(ctx.projectPath!, "appPackage", "manifest.json")) {
          updateManifest = true;
          updatedManifestData = data;
        } else {
          throw new Error("not support " + file);
        }
      });

      sandbox.stub(envUtil, "writeEnv").resolves(ok(undefined));
      sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(existingManifest));
      const res = await developerPortalScaffoldUtils.updateFilesForTdp(ctx, appDefinition, inputs);

      chai.assert.isTrue(res.isOk());
      chai.assert.isTrue(updateManifest);
      chai.assert.isTrue(updateColor);
      chai.assert.isTrue(updateOutline);
      chai.assert.isTrue(updateLanguage);
      const updatedManifest = JSON.parse(updatedManifestData) as TeamsAppManifest;
      chai.assert.equal(updatedManifest.id, "${{TEAMS_APP_ID}}");
      chai.assert.deepEqual(updatedManifest.bots![0], manifest.bots![0]);
      chai.assert.deepEqual(
        updatedManifest.composeExtensions![0],
        existingManifest.composeExtensions![0]
      );
      chai.assert.equal(updatedManifest.developer.privacyUrl, DEFAULT_DEVELOPER.privacyUrl);
      chai.assert.equal(updatedManifest.developer.termsOfUseUrl, DEFAULT_DEVELOPER.termsOfUseUrl);
      chai.assert.equal(updatedManifest.developer.websiteUrl, DEFAULT_DEVELOPER.websiteUrl);
      chai.assert.equal(updatedManifest.validDomains?.length, 0);
    });

    it("update bot id and message extension id", async () => {
      const ctx = createContext();
      ctx.tokenProvider = {
        m365TokenProvider: new MockedM365Provider(),
        azureAccountProvider: new MockedAzureAccountProvider(),
      };
      ctx.projectPath = "project-path";
      const appDefinition: AppDefinition = {
        appId: "mock-app-id",
        teamsAppId: "mock-app-id",
        staticTabs: [
          {
            objectId: "objId",
            entityId: "entityId",
            name: "tab",
            contentUrl: "https://url",
            websiteUrl: "https:/url",
            scopes: [],
            context: [],
          },
        ],
      };
      const inputs: Inputs = {
        platform: Platform.VSCode,
        [QuestionNames.ReplaceBotIds]: ["bot", "messageExtension"],
        projectPath: "project-path",
      };
      const manifest: TeamsAppManifest = {
        manifestVersion: "version",
        id: "mock-app-id",
        name: { short: "short-name" },
        description: { short: "", full: "" },
        version: "version",
        icons: { outline: "outline.png", color: "color.png" },
        accentColor: "#ffffff",
        developer: {
          privacyUrl: "",
          websiteUrl: "",
          termsOfUseUrl: "",
          name: "developer-name",
        },
        bots: [
          {
            botId: "botId0",
            scopes: ["personal"],
            commandLists: [],
          },
        ],
        composeExtensions: [
          {
            botId: "botId1",
            commands: [],
          },
        ],
        validDomains: [],
      };

      const existingManifest: TeamsAppManifest = {
        manifestVersion: "version",
        id: "mock-app-id",
        name: { short: "short-name" },
        description: { short: "", full: "" },
        version: "version",
        icons: { outline: "outline.png", color: "color.png" },
        accentColor: "#ffffff",
        developer: {
          privacyUrl: "",
          websiteUrl: "",
          termsOfUseUrl: "",
          name: "developer-name",
        },
        bots: [
          {
            botId: "{{BOT_ID}}",
            scopes: ["personal", "team"],
            supportsFiles: false,
            isNotificationOnly: false,
            commandLists: [
              {
                scopes: ["personal", "team", "groupchat"],
                commands: [],
              },
            ],
          },
        ],
        validDomains: ["valid-domain"],
      };

      let updateManifest = false;
      let updateLanguage = false;
      let updateColor = false;
      let updateOutline = false;
      let updatedManifestData = "";
      sandbox.stub(appStudio, "getAppPackage").resolves(
        ok({
          manifest: Buffer.from(JSON.stringify(manifest)),
          icons: { color: Buffer.from(""), outline: Buffer.from("") },
          languages: { zh: Buffer.from(JSON.stringify({})) },
        })
      );
      sandbox.stub(fs, "writeFile").callsFake((file: number | fs.PathLike, data: any) => {
        if (file === path.join(ctx.projectPath!, "appPackage", "color.png")) {
          updateColor = true;
        } else if (file === path.join(ctx.projectPath!, "appPackage", "outline.png")) {
          updateOutline = true;
        } else if (file === path.join(ctx.projectPath!, "appPackage", "zh.json")) {
          updateLanguage = true;
        } else if (file === path.join(ctx.projectPath!, "appPackage", "manifest.json")) {
          updateManifest = true;
          updatedManifestData = data;
        } else {
          throw new Error("not support " + file);
        }
      });

      sandbox.stub(envUtil, "writeEnv").resolves(ok(undefined));
      sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(existingManifest));
      const res = await developerPortalScaffoldUtils.updateFilesForTdp(ctx, appDefinition, inputs);

      chai.assert.isTrue(res.isOk());
      chai.assert.isTrue(updateManifest);
      chai.assert.isTrue(updateColor);
      chai.assert.isTrue(updateOutline);
      chai.assert.isTrue(updateLanguage);
      const updatedManifest = JSON.parse(updatedManifestData) as TeamsAppManifest;
      chai.assert.equal(updatedManifest.id, "${{TEAMS_APP_ID}}");
      chai.assert.deepEqual(updatedManifest.bots![0], existingManifest.bots![0]);
      chai.assert.deepEqual(updatedManifest.composeExtensions![0].botId, "${{BOT_ID}}");
      chai.assert.deepEqual(
        updatedManifest.composeExtensions![0].commands,
        COMPOSE_EXTENSIONS_TPL_V3[0]!.commands
      );
      chai.assert.equal(updatedManifest.developer.privacyUrl, DEFAULT_DEVELOPER.privacyUrl);
      chai.assert.equal(updatedManifest.developer.termsOfUseUrl, DEFAULT_DEVELOPER.termsOfUseUrl);
      chai.assert.equal(updatedManifest.developer.websiteUrl, DEFAULT_DEVELOPER.websiteUrl);
      chai.assert.equal(updatedManifest.validDomains?.length, 0);
    });

    it("update manifest if selecting capability from ttk UI", async () => {
      const ctx = createContext();
      ctx.tokenProvider = {
        m365TokenProvider: new MockedM365Provider(),
        azureAccountProvider: new MockedAzureAccountProvider(),
      };
      ctx.projectPath = "project-path";
      const appDefinition: AppDefinition = {
        appId: "mock-app-id",
        teamsAppId: "mock-app-id",
      };
      const inputs: Inputs = {
        platform: Platform.VSCode,
        [QuestionNames.ReplaceBotIds]: ["bot", "messageExtension"],
        projectPath: "project-path",
      };
      const manifest: TeamsAppManifest = {
        manifestVersion: "version",
        id: "mock-app-id",
        name: { short: "short-name" },
        description: { short: "", full: "" },
        version: "version",
        icons: { outline: "outline.png", color: "color.png" },
        accentColor: "#ffffff",
        developer: {
          privacyUrl: "",
          websiteUrl: "",
          termsOfUseUrl: "",
          name: "developer-name",
        },
      };

      const existingManifest: TeamsAppManifest = {
        manifestVersion: "version",
        id: "mock-app-id",
        name: { short: "short-name" },
        description: { short: "", full: "" },
        version: "version",
        icons: { outline: "outline.png", color: "color.png" },
        accentColor: "#ffffff",
        developer: {
          privacyUrl: "",
          websiteUrl: "",
          termsOfUseUrl: "",
          name: "developer-name",
        },
        webApplicationInfo: {
          id: "1",
        },
        bots: [
          {
            botId: "{{BOT_ID}}",
            scopes: ["personal", "team"],
            supportsFiles: false,
            isNotificationOnly: false,
          },
        ],
        validDomains: ["valid-domain"],
      };

      let updateManifest = false;
      let updateLanguage = false;
      let updateColor = false;
      let updateOutline = false;
      let updatedManifestData = "";
      sandbox.stub(appStudio, "getAppPackage").resolves(
        ok({
          manifest: Buffer.from(
            JSON.stringify({
              manifestVersion: "version",
              id: "mock-app-id",
              name: { short: "short-name" },
              description: { short: "", full: "" },
              version: "version",
              icons: { outline: "outline.png", color: "color.png" },
              accentColor: "#ffffff",
              developer: {
                privacyUrl: "",
                websiteUrl: "",
                termsOfUseUrl: "",
                name: "developer-name",
              },
            })
          ),
          icons: { color: Buffer.from(""), outline: Buffer.from("") },
          languages: { zh: Buffer.from(JSON.stringify({})) },
        })
      );
      sandbox.stub(fs, "writeFile").callsFake((file: number | fs.PathLike, data: any) => {
        if (file === path.join(ctx.projectPath!, "appPackage", "color.png")) {
          updateColor = true;
        } else if (file === path.join(ctx.projectPath!, "appPackage", "outline.png")) {
          updateOutline = true;
        } else if (file === path.join(ctx.projectPath!, "appPackage", "zh.json")) {
          updateLanguage = true;
        } else if (file === path.join(ctx.projectPath!, "appPackage", "manifest.json")) {
          updateManifest = true;
          updatedManifestData = data;
        } else {
          throw new Error("not support " + file);
        }
      });

      sandbox.stub(envUtil, "writeEnv").resolves(ok(undefined));
      sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(existingManifest));
      const res = await developerPortalScaffoldUtils.updateFilesForTdp(ctx, appDefinition, inputs);

      chai.assert.isTrue(res.isOk());
      chai.assert.isTrue(updateManifest);
      chai.assert.isTrue(updateColor);
      chai.assert.isTrue(updateOutline);
      chai.assert.isTrue(updateLanguage);
      const updatedManifest = JSON.parse(updatedManifestData) as TeamsAppManifest;
      chai.assert.equal(updatedManifest.id, "${{TEAMS_APP_ID}}");
      chai.assert.deepEqual(updatedManifest.bots![0], existingManifest.bots![0]);
      chai.assert.equal(updatedManifest.developer.privacyUrl, DEFAULT_DEVELOPER.privacyUrl);
      chai.assert.equal(updatedManifest.developer.termsOfUseUrl, DEFAULT_DEVELOPER.termsOfUseUrl);
      chai.assert.equal(updatedManifest.developer.websiteUrl, DEFAULT_DEVELOPER.websiteUrl);
      chai.assert.deepEqual(
        updatedManifest.webApplicationInfo,
        existingManifest.webApplicationInfo
      );
      chai.assert.isTrue(updatedManifest.validDomains?.includes("valid-domain"));
    });

    it("update group chat", async () => {
      const ctx = createContext();
      ctx.tokenProvider = {
        m365TokenProvider: new MockedM365Provider(),
        azureAccountProvider: new MockedAzureAccountProvider(),
      };
      ctx.projectPath = "project-path";
      const appDefinition: AppDefinition = {
        appId: "mock-app-id",
        teamsAppId: "mock-app-id",
        staticTabs: [
          {
            objectId: "objId",
            entityId: "entityId",
            name: "tab",
            contentUrl: "https://url",
            websiteUrl: "https:/url",
            scopes: [],
            context: [],
          },
        ],
      };
      const inputs: Inputs = {
        platform: Platform.VSCode,
        projectPath: "project-path",
      };
      const manifest = {
        manifestVersion: "1.17",
        id: "mock-app-id",
        name: { short: "short-name" },
        description: { short: "", full: "" },
        version: "version",
        icons: { outline: "outline.png", color: "color.png" },
        accentColor: "#ffffff",
        developer: {
          privacyUrl: "",
          websiteUrl: "",
          termsOfUseUrl: "",
          name: "developer-name",
        },
        configurableTabs: [
          {
            configurationUrl: "url",
            scopes: ["groupchat", "team"] as any,
          },
        ],
        bots: [
          {
            botId: "botId",
            scopes: ["groupchat"],
            commandLists: [
              {
                commands: [
                  {
                    title: "tt",
                    description: "ttt",
                  },
                ],
                scopes: ["groupChat"],
              },
            ],
          },
        ],
        composeExtensions: [
          {
            botId: "botId",
            scopes: ["groupchat"],
          },
        ],
      };

      let updateManifest = false;
      let updateLanguage = false;
      let updateColor = false;
      let updateOutline = false;
      let updatedManifestData = "";
      sandbox.stub(appStudio, "getAppPackage").resolves(
        ok({
          manifest: Buffer.from(JSON.stringify(manifest)),
          icons: { color: Buffer.from(""), outline: Buffer.from("") },
          languages: { zh: Buffer.from(JSON.stringify({})) },
        })
      );
      sandbox.stub(envUtil, "writeEnv").resolves(ok(undefined));
      sandbox.stub(fs, "writeFile").callsFake((file: number | fs.PathLike, data: any) => {
        if (file === path.join(ctx.projectPath!, "appPackage", "color.png")) {
          updateColor = true;
        } else if (file === path.join(ctx.projectPath!, "appPackage", "outline.png")) {
          updateOutline = true;
        } else if (file === path.join(ctx.projectPath!, "appPackage", "zh.json")) {
          updateLanguage = true;
        } else if (file === path.join(ctx.projectPath!, "appPackage", "manifest.json")) {
          updateManifest = true;
          updatedManifestData = data;
        } else {
          throw new Error("not support " + file);
        }
      });

      const mockWriteStream = new MockedWriteStream();
      sandbox.stub(fs, "createWriteStream").returns(mockWriteStream as any);
      const writeSpy = sandbox.stub(mockWriteStream, "write").resolves();
      sandbox.stub(mockWriteStream, "end").resolves();
      sandbox.stub(fs, "readFile").callsFake((file: number | fs.PathLike) => {
        if (file === path.join(ctx.projectPath!, "env", ".env.local")) {
          return Promise.resolve(Buffer.from("TEAMS_APP_ID=\nENV=\n"));
        } else {
          throw new Error("not support " + file);
        }
      });
      sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest as TeamsAppManifest));

      const res = await developerPortalScaffoldUtils.updateFilesForTdp(ctx, appDefinition, inputs);

      chai.assert.isTrue(res.isOk());
      chai.assert.isTrue(updateManifest);
      chai.assert.isTrue(updateColor);
      chai.assert.isTrue(updateOutline);
      chai.assert.isTrue(updateLanguage);
      const updatedManifest = JSON.parse(updatedManifestData) as TeamsAppManifest;
      chai.assert.equal(updatedManifest.id, "${{TEAMS_APP_ID}}");
      chai.assert.isTrue(
        (updatedManifest.configurableTabs![0].scopes as string[]).includes("groupChat")
      );
      chai.assert.isTrue((updatedManifest.bots![0].scopes as string[]).includes("groupChat"));
      chai.assert.isTrue(
        (updatedManifest.bots![0].commandLists![0].scopes as string[]).includes("groupChat")
      );
      chai.assert.isTrue(
        (updatedManifest.composeExtensions![0].scopes! as string[]).includes("groupChat")
      );
      chai.assert.equal(updatedManifest.developer.privacyUrl, DEFAULT_DEVELOPER.privacyUrl);
      chai.assert.equal(updatedManifest.developer.termsOfUseUrl, DEFAULT_DEVELOPER.termsOfUseUrl);
      chai.assert.equal(updatedManifest.developer.websiteUrl, DEFAULT_DEVELOPER.websiteUrl);
      chai.assert.equal(updatedManifest.validDomains, undefined);
    });

    it("success without the need to update group chat", async () => {
      const ctx = createContext();
      ctx.tokenProvider = {
        m365TokenProvider: new MockedM365Provider(),
        azureAccountProvider: new MockedAzureAccountProvider(),
      };
      ctx.projectPath = "project-path";
      const appDefinition: AppDefinition = {
        appId: "mock-app-id",
        teamsAppId: "mock-app-id",
        staticTabs: [
          {
            objectId: "objId",
            entityId: "entityId",
            name: "tab",
            contentUrl: "https://url",
            websiteUrl: "https:/url",
            scopes: [],
            context: [],
          },
        ],
      };
      const inputs: Inputs = {
        platform: Platform.VSCode,
        projectPath: "project-path",
      };
      const manifest = {
        manifestVersion: "1.17",
        id: "mock-app-id",
        name: { short: "short-name" },
        description: { short: "", full: "" },
        version: "version",
        icons: { outline: "outline.png", color: "color.png" },
        accentColor: "#ffffff",
        developer: {
          privacyUrl: "",
          websiteUrl: "",
          termsOfUseUrl: "",
          name: "developer-name",
        },
        configurableTabs: [
          {
            configurationUrl: "url",
          },
        ],
        bots: [
          {
            botId: "botId",
            commandLists: [
              {
                commands: [
                  {
                    title: "tt",
                    description: "ttt",
                  },
                ],
              },
            ],
          },
        ],
        composeExtensions: [
          {
            botId: "botId",
          },
        ],
      };

      let updateManifest = false;
      let updateLanguage = false;
      let updateColor = false;
      let updateOutline = false;
      let updatedManifestData = "";
      sandbox.stub(appStudio, "getAppPackage").resolves(
        ok({
          manifest: Buffer.from(JSON.stringify(manifest)),
          icons: { color: Buffer.from(""), outline: Buffer.from("") },
          languages: { zh: Buffer.from(JSON.stringify({})) },
        })
      );
      sandbox.stub(envUtil, "writeEnv").resolves(ok(undefined));
      sandbox.stub(fs, "writeFile").callsFake((file: number | fs.PathLike, data: any) => {
        if (file === path.join(ctx.projectPath!, "appPackage", "color.png")) {
          updateColor = true;
        } else if (file === path.join(ctx.projectPath!, "appPackage", "outline.png")) {
          updateOutline = true;
        } else if (file === path.join(ctx.projectPath!, "appPackage", "zh.json")) {
          updateLanguage = true;
        } else if (file === path.join(ctx.projectPath!, "appPackage", "manifest.json")) {
          updateManifest = true;
          updatedManifestData = data;
        } else {
          throw new Error("not support " + file);
        }
      });

      const mockWriteStream = new MockedWriteStream();
      sandbox.stub(fs, "createWriteStream").returns(mockWriteStream as any);
      const writeSpy = sandbox.stub(mockWriteStream, "write").resolves();
      sandbox.stub(mockWriteStream, "end").resolves();
      sandbox.stub(fs, "readFile").callsFake((file: number | fs.PathLike) => {
        if (file === path.join(ctx.projectPath!, "env", ".env.local")) {
          return Promise.resolve(Buffer.from("TEAMS_APP_ID=\nENV=\n"));
        } else {
          throw new Error("not support " + file);
        }
      });
      sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest as TeamsAppManifest));

      const res = await developerPortalScaffoldUtils.updateFilesForTdp(ctx, appDefinition, inputs);

      chai.assert.isTrue(res.isOk());
      chai.assert.isTrue(updateManifest);
      chai.assert.isTrue(updateColor);
      chai.assert.isTrue(updateOutline);
      chai.assert.isTrue(updateLanguage);
      const updatedManifest = JSON.parse(updatedManifestData) as TeamsAppManifest;
      chai.assert.equal(updatedManifest.id, "${{TEAMS_APP_ID}}");
    });

    it("read manifest error", async () => {
      const ctx = createContext();
      ctx.tokenProvider = {
        m365TokenProvider: new MockedM365Provider(),
        azureAccountProvider: new MockedAzureAccountProvider(),
      };
      ctx.projectPath = "project-path";
      const appDefinition: AppDefinition = {
        appId: "mock-app-id",
        teamsAppId: "mock-app-id",
        staticTabs: [
          {
            objectId: "objId",
            entityId: "entityId",
            name: "tab",
            contentUrl: "https://url",
            websiteUrl: "https:/url",
            scopes: [],
            context: [],
          },
        ],
      };
      const inputs: Inputs = {
        platform: Platform.VSCode,
        [QuestionNames.ReplaceWebsiteUrl]: ["name0"],
        [QuestionNames.ReplaceContentUrl]: ["name1"],
        projectPath: "project-path",
      };
      const manifest: TeamsAppManifest = {
        manifestVersion: "version",
        id: "mock-app-id",
        name: { short: "short-name" },
        description: { short: "", full: "" },
        version: "version",
        icons: { outline: "outline.png", color: "color.png" },
        accentColor: "#ffffff",
        developer: {
          privacyUrl: "",
          websiteUrl: "",
          termsOfUseUrl: "",
          name: "developer-name",
        },
        staticTabs: [
          {
            name: "name0",
            entityId: "index0",
            scopes: ["personal"],
            contentUrl: "contentUrl0",
            websiteUrl: "websiteUrl0",
          },
          {
            name: "name1",
            entityId: "index1",
            scopes: ["personal"],
            contentUrl: "contentUrl1",
            websiteUrl: "websiteUrl1",
          },
        ],
      };

      sandbox.stub(appStudio, "getAppPackage").resolves(
        ok({
          manifest: Buffer.from(JSON.stringify(manifest)),
          icons: { color: Buffer.from(""), outline: Buffer.from("") },
          languages: { zh: Buffer.from(JSON.stringify({})) },
        })
      );

      sandbox.stub(envUtil, "writeEnv").resolves(ok(undefined));

      sandbox.stub(manifestUtils, "_readAppManifest").resolves(err(new UserError("", "", "", "")));
      const res = await developerPortalScaffoldUtils.updateFilesForTdp(ctx, appDefinition, inputs);

      chai.assert.isTrue(res.isErr());
    });
  });

  describe("getProjectTypeAndCapability", () => {
    const validBot: Bot = {
      botId: "botId",
      isNotificationOnly: false,
      needsChannelSelector: false,
      personalCommands: [{ title: "title", description: "description" }],
      supportsFiles: false,
      supportsCalling: false,
      supportsVideo: false,
      teamCommands: [{ title: "title", description: "description" }],
      groupChatCommands: [{ title: "title", description: "description" }],
      scopes: ["scope"],
    };

    const validConfigurableTabForTabCode: ConfigurableTab = {
      objectId: "objId",
      configurationUrl: "https://url",
      canUpdateConfiguration: false,
      scopes: [CommandScope.GroupChat],
      context: [MeetingsContext.ChannelTab],
      sharePointPreviewImage: "img",
      supportedSharePointHosts: [],
    };

    const validStaticTab: StaticTab = {
      objectId: "objId",
      entityId: "entityId",
      name: "tab",
      contentUrl: "https://url",
      websiteUrl: "https:/url",
      scopes: [],
      context: [],
    };

    const validMessagingExtension: MessagingExtension = {
      objectId: "objId",
      botId: "botId",
      canUpdateConfiguration: true,
      commands: [],
      messageHandlers: [],
    };

    it("return TabNonSsoAndDefaultBot", () => {
      const appDefinition: AppDefinition = {
        teamsAppId: "id",
        staticTabs: [validStaticTab],
        messagingExtensions: [validMessagingExtension],
      };

      const res = getProjectTypeAndCapability(appDefinition);
      chai.assert.equal(res?.templateId, CapabilityOptions.nonSsoTabAndBot().id);
      chai.assert.equal(res?.projectType, "tab-bot-type");
    });

    it("return TabNonSso", () => {
      const appDefinition: AppDefinition = {
        teamsAppId: "id",
        staticTabs: [validStaticTab],
      };

      const res = getProjectTypeAndCapability(appDefinition);
      chai.assert.equal(res?.templateId, CapabilityOptions.nonSsoTab().id);
      chai.assert.equal(res?.projectType, "tab-type");
    });

    it("return DefaultBotAndMessageExtension", () => {
      const appDefinition: AppDefinition = {
        teamsAppId: "id",
        bots: [validBot],
        messagingExtensions: [validMessagingExtension],
      };

      const res = getProjectTypeAndCapability(appDefinition);
      chai.assert.equal(res?.templateId, CapabilityOptions.botAndMe().id);
      chai.assert.equal(res?.projectType, "bot-me-type");
    });

    it("return MessageExtension", () => {
      const appDefinition: AppDefinition = {
        teamsAppId: "id",
        messagingExtensions: [validMessagingExtension],
      };

      const res = getProjectTypeAndCapability(appDefinition);
      chai.assert.equal(res?.templateId, CapabilityOptions.me().id);
      chai.assert.equal(res?.projectType, "me-type");
    });

    it("return bot", () => {
      const appDefinition: AppDefinition = {
        teamsAppId: "id",
        bots: [validBot],
      };

      const res = getProjectTypeAndCapability(appDefinition);
      chai.assert.equal(res?.templateId, CapabilityOptions.basicBot().id);
      chai.assert.equal(res?.projectType, "bot-type");
    });

    it("return undefined", () => {
      const appDefinition: AppDefinition = {
        teamsAppId: "id",
      };

      const res = getProjectTypeAndCapability(appDefinition);
      chai.assert.isUndefined(res);
    });
  });

  describe("adjustScopeBasedOnVersion", () => {
    it("devPreview", () => {
      const res = adjustScopeBasedOnVersion(["groupchat"], "devPreview");
      chai.assert.deepEqual(res, ["groupChat"]);
    });

    it("1.17", () => {
      const res = adjustScopeBasedOnVersion(["groupchat"], "1.17");
      chai.assert.deepEqual(res, ["groupChat"]);
    });

    it("1.16", () => {
      const res = adjustScopeBasedOnVersion(["groupChat", "team"], "1.16");
      chai.assert.deepEqual(res, ["groupchat", "team"]);
    });
  });
});
