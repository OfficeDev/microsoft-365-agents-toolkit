import { err, ok, SystemError, SystemErrorOptions, TeamsAppManifest } from "@microsoft/teamsfx-api";
import { copilotGptManifestUtils, envUtil, featureFlagManager } from "@microsoft/teamsfx-core";
import * as chai from "chai";
import fs from "fs-extra";
import { afterEach, describe } from "mocha";
import * as path from "path";
import * as sinon from "sinon";
import * as vscode from "vscode";
import {
  AadAppTemplateCodeLensProvider,
  ApiPluginCodeLensProvider,
  CopilotPluginCodeLensProvider,
  CryptoCodeLensProvider,
  DeclarativeAgentSensitivityLabelCodeLensProvider,
  ManifestTemplateCodeLensProvider,
  OfficeDevManifestCodeLensProvider,
  OneDriveSharePointCodeLensProvider,
  PermissionsJsonFileCodeLensProvider,
  PlaceholderCodeLens,
  SharePointIdCodeLens,
  TeamsAppYamlCodeLensProvider,
  WorkspaceMCPConfigCodeLensProvider,
} from "../../src/codeLensProvider";
import * as globalVariables from "../../src/globalVariables";
import { setTools, tools } from "../../src/globalVariables";
import { TelemetryTriggerFrom } from "../../src/telemetry/extTelemetryEvents";
import { MockTools } from "../mocks/mockTools";
import { GraphClient } from "@microsoft/teamsfx-core/build/client/graphClient";

describe("CodeLens Provider", () => {
  afterEach(() => {
    sinon.restore();
  });

  describe("Manifest codelens", () => {
    const sandbox = sinon.createSandbox();

    beforeEach(() => {
      sandbox.stub(envUtil, "readEnv").resolves(ok({}));
    });

    afterEach(() => {
      sandbox.restore();
    });

    it("Template codelens - V3", async () => {
      const url =
        "https://developer.microsoft.com/en-us/json-schemas/teams/v1.14/MicrosoftTeams.schema.json";
      const document = {
        fileName: "manifest.template.json",
        getText: () => {
          return `"$schema": "${url}",`;
        },
        positionAt: () => {
          return new vscode.Position(0, 0);
        },
        lineAt: () => {
          return {
            lineNumber: 0,
            text: `"$schema": "${url}",`,
          };
        },
      } as any as vscode.TextDocument;

      const manifestProvider = new ManifestTemplateCodeLensProvider();
      const codelens: vscode.CodeLens[] = manifestProvider.provideCodeLenses(
        document
      ) as vscode.CodeLens[];

      chai.assert.equal(codelens.length, 1);
      chai.expect(codelens[0].command).to.deep.equal({
        title: "Open schema",
        command: "fx-extension.openSchema",
        arguments: [{ url: url }],
      });
    });

    it("ResolveEnvironmentVariableCodelens", async () => {
      const range = new vscode.Range(new vscode.Position(0, 0), new vscode.Position(0, 0));
      const lens: PlaceholderCodeLens = new PlaceholderCodeLens(
        "${{ TEAMS_APP_ID }}",
        range,
        "manifest.template.json"
      );
      const manifestProvider = new ManifestTemplateCodeLensProvider();
      const cts = new vscode.CancellationTokenSource();

      const res = await manifestProvider.resolveCodeLens(lens, cts.token);
      chai.assert.equal(res.command?.command, "fx-extension.openConfigState");
      chai.assert.isTrue(res.command?.title.includes("👉"));
      chai.expect(res.command?.arguments).to.deep.equal([{ type: "env", from: "manifest" }]);
    });

    it("ResolveEnvironmentVariableCodelens for AAD manifest", async () => {
      const range = new vscode.Range(new vscode.Position(0, 0), new vscode.Position(0, 0));
      const lens: PlaceholderCodeLens = new PlaceholderCodeLens(
        "${{ TEAMS_APP_ID }}",
        range,
        "aad.template.json"
      );
      const aadProvider = new AadAppTemplateCodeLensProvider();
      const cts = new vscode.CancellationTokenSource();

      const res = await aadProvider.resolveCodeLens(lens, cts.token);
      chai.assert.equal(res.command?.command, "fx-extension.openConfigState");
      chai.assert.isTrue(res.command?.title.includes("👉"));
      chai.expect(res.command?.arguments).to.deep.equal([{ type: "env", from: "aad" }]);
    });

    it("ComputeTemplateCodeLenses for AAD manifest template", async () => {
      const document = <vscode.TextDocument>{
        fileName: "./aad.manifest.json",
        getText: () => {
          return '{"name": "test"}';
        },
      };

      const aadProvider = new AadAppTemplateCodeLensProvider();
      const res = await aadProvider.provideCodeLenses(document);
      chai.assert.isTrue(
        res != null && res[0].command!.command === "fx-extension.openPreviewAadFile"
      );

      chai.assert.isTrue(
        res != null && res[1].command!.command === "fx-extension.convertAadToNewSchema"
      );
    });

    it("ComputeTemplateCodeLenses for AAD manifest template with new schema", async () => {
      const document = <vscode.TextDocument>{
        fileName: "./aad.manifest.json",
        getText: () => {
          return '{"displayName": "test"}';
        },
      };

      const aadProvider = new AadAppTemplateCodeLensProvider();
      const res = await aadProvider.provideCodeLenses(document);
      chai.assert.isTrue(
        res != null &&
          res.length === 1 &&
          res[0].command!.command === "fx-extension.openPreviewAadFile"
      );
    });

    it("ComputeTemplateCodeLenses for aad manifest", async () => {
      sandbox.stub(fs, "pathExistsSync").returns(true);
      const document = <vscode.TextDocument>{
        fileName: "./build/aad.manifest.dev.json",
        getText: () => {
          return "{name: 'test'}";
        },
      };

      sandbox
        .stub(vscode.workspace, "workspaceFolders")
        .value([{ uri: { fsPath: "workspacePath" } }]);

      const aadProvider = new AadAppTemplateCodeLensProvider();
      const res = await aadProvider.provideCodeLenses(document);
      chai.assert.isTrue(
        res != null && res[0].command!.command === "fx-extension.updateAadAppManifest"
      );

      chai.assert.isTrue(
        res != null && res[1].command!.command === "fx-extension.editAadManifestTemplate"
      );
    });

    it("ComputeTemplateCodeLenses for aad manifest if template not exist", async () => {
      sandbox.stub(fs, "pathExistsSync").returns(false);
      const document = <vscode.TextDocument>{
        fileName: "./build/aad.manifest.dev.json",
        getText: () => {
          return "{name: 'test'}";
        },
      };

      sandbox
        .stub(vscode.workspace, "workspaceFolders")
        .value([{ uri: { fsPath: "workspacePath" } }]);

      const aadProvider = new AadAppTemplateCodeLensProvider();
      const res = await aadProvider.provideCodeLenses(document);

      chai.assert.isTrue(
        res != null &&
          res.length === 1 &&
          res[0].command!.command === "fx-extension.updateAadAppManifest"
      );
    });

    it("PermissionsJsonFileCodeLensProvider for Microsoft Entra manifest template", async () => {
      sandbox.stub(fs, "pathExistsSync").returns(true);
      sandbox
        .stub(vscode.workspace, "workspaceFolders")
        .value([{ uri: { fsPath: "workspacePath" } }]);
      const document = <vscode.TextDocument>{
        fileName: "./aad.manifest.json",
        getText: () => {
          return "{name: 'test'}";
        },
      };

      const permissionsJsonFile = new PermissionsJsonFileCodeLensProvider();
      const res = await permissionsJsonFile.provideCodeLenses(document);
      chai.assert.isTrue(
        res != null && res[0].command!.command === "fx-extension.editAadManifestTemplate"
      );
    });

    it("ComputePreAuthAppCodeLenses for AAD manifest template", async () => {
      const document = {
        fileName: "./aad.manifest.json",
        getText: () => {
          return `{
            "api": {
              "preAuthorizedApplications": [
                {
                  "appId": "1fec8e78-bce4-4aaf-ab1b-5451cc387264"
                }
              ]
            }
          }`;
        },
        positionAt: () => {
          return new vscode.Position(0, 0);
        },
      } as any as vscode.TextDocument;

      const aadProvider = new AadAppTemplateCodeLensProvider();
      const res = await aadProvider.provideCodeLenses(document);
      chai.assert.isTrue(res != null && res[0].command!.title.includes("resource name"));
    });
  });

  describe("Crypto CodeLensProvider", () => {
    const sandbox = sinon.createSandbox();

    afterEach(() => {
      sandbox.restore();
    });

    it("envData codelens", async () => {
      const document = {
        fileName: ".env.local",
        getText: () => {
          return "SECRET_VAR_2=crypto_abc";
        },
        lineAt: () => {
          return {
            lineNumber: 0,
            text: "SECRET_VAR_2=crypto_abc",
          };
        },
        positionAt: () => {
          return {
            character: 0,
            line: 0,
          };
        },
      } as unknown as vscode.TextDocument;

      const cryptoProvider = new CryptoCodeLensProvider();
      const codelens: vscode.CodeLens[] = cryptoProvider.provideCodeLenses(
        document
      ) as vscode.CodeLens[];

      chai.assert.equal(codelens.length, 1);
      chai.expect(codelens[0].command?.title).equal("🔑Decrypt secret");
      chai.expect(codelens[0].command?.command).equal("fx-extension.decryptSecret");
    });

    it("hides when command is running", async () => {
      sandbox.stub(globalVariables, "commandIsRunning").value(true);
      const document = {
        fileName: ".env.local",
        getText: () => {
          return "SECRET_VAR_2=crypto_abc";
        },
        lineAt: () => {
          return {
            lineNumber: 0,
            text: "SECRET_VAR_2=crypto_abc",
          };
        },
        positionAt: () => {
          return {
            character: 0,
            line: 0,
          };
        },
      } as unknown as vscode.TextDocument;

      const cryptoProvider = new CryptoCodeLensProvider();
      const codelens: vscode.CodeLens[] = cryptoProvider.provideCodeLenses(
        document
      ) as vscode.CodeLens[];

      chai.assert.equal(codelens.length, 0);
    });
  });

  describe("API ME CodeLensProvider", () => {
    const sandbox = sinon.createSandbox();

    afterEach(() => {
      sandbox.restore();
    });

    it("Add API", async () => {
      const manifest = new TeamsAppManifest();
      manifest.composeExtensions = [
        {
          composeExtensionType: "apiBased",
          commands: [],
        },
      ];
      const manifestString = JSON.stringify(manifest);
      const document = {
        fileName: "manifest.json",
        getText: () => {
          return manifestString;
        },
        positionAt: () => {
          return new vscode.Position(0, 0);
        },
        lineAt: () => {
          return {
            lineNumber: 0,
            text: manifestString,
          };
        },
      } as any as vscode.TextDocument;

      const copilotPluginCodelensProvider = new CopilotPluginCodeLensProvider();
      const codelens: vscode.CodeLens[] = copilotPluginCodelensProvider.provideCodeLenses(
        document
      ) as vscode.CodeLens[];

      chai.assert.equal(codelens.length, 1);
      chai.expect(codelens[0].command).to.deep.equal({
        title: "➕Add another API",
        command: "fx-extension.copilotPluginAddAPI",
        arguments: [{ fsPath: document.fileName }],
      });
    });

    it("Do not show codelens for non-copilot plugin project", async () => {
      const manifest = new TeamsAppManifest();
      const manifestString = JSON.stringify(manifest);
      const document = {
        fileName: "manifest.json",
        getText: () => {
          return manifestString;
        },
        positionAt: () => {
          return new vscode.Position(0, 0);
        },
        lineAt: () => {
          return {
            lineNumber: 0,
            text: manifestString,
          };
        },
      } as any as vscode.TextDocument;

      const copilotPluginCodelensProvider = new CopilotPluginCodeLensProvider();
      const codelens: vscode.CodeLens[] = copilotPluginCodelensProvider.provideCodeLenses(
        document
      ) as vscode.CodeLens[];

      chai.assert.equal(codelens.length, 0);
    });
  });

  describe("Api plugin CodeLensProvider", () => {
    const sandbox = sinon.createSandbox();

    afterEach(() => {
      sandbox.restore();
    });

    it("Add API", async () => {
      const manifest = new TeamsAppManifest();
      manifest.copilotExtensions = {
        plugins: [
          {
            file: "test.json",
            id: "plugin1",
          },
        ],
      };
      const openApiObject = {
        openapi: "3.0",
      };
      const text = JSON.stringify(openApiObject);
      const document = {
        fileName: "openapi.yaml",
        getText: () => {
          return text;
        },
        positionAt: () => {
          return new vscode.Position(0, 0);
        },
        lineAt: () => {
          return {
            lineNumber: 0,
            text,
          };
        },
      } as any as vscode.TextDocument;

      sandbox.stub(fs, "existsSync").returns(true);
      sandbox.stub(fs, "readFileSync").returns(JSON.stringify(manifest));
      sandbox
        .stub(globalVariables, "workspaceUri")
        .value(vscode.Uri.parse(path.resolve(__dirname, "unknown")));
      const apiPluginCodelensProvider = new ApiPluginCodeLensProvider();
      const codelens: vscode.CodeLens[] = apiPluginCodelensProvider.provideCodeLenses(
        document
      ) as vscode.CodeLens[];

      chai.assert.equal(codelens.length, 1);
      chai.expect(codelens[0].command!.title).to.equal("➕Add another API");
      chai.expect(codelens[0].command!.command).to.equal("fx-extension.copilotPluginAddAPI");
      chai.expect(codelens[0].command!.arguments![0].fsPath).to.equal(document.fileName);
      chai.expect(codelens[0].command!.arguments![0].isFromApiPlugin).to.be.true;
    });

    it("Do not show codelens for if not api spec file", async () => {
      const openApiObject = {
        unknown: "3.0",
      };
      const text = JSON.stringify(openApiObject);
      const document = {
        fileName: "openapi.yaml",
        getText: () => {
          return text;
        },
        positionAt: () => {
          return new vscode.Position(0, 0);
        },
        lineAt: () => {
          return {
            lineNumber: 0,
            text,
          };
        },
      } as any as vscode.TextDocument;

      sandbox.stub(fs, "existsSync").returns(false);
      sandbox
        .stub(globalVariables, "workspaceUri")
        .value(vscode.Uri.parse(path.resolve(__dirname, "unknown")));
      const apiPluginCodelensProvider = new ApiPluginCodeLensProvider();
      const codelens: vscode.CodeLens[] = apiPluginCodelensProvider.provideCodeLenses(
        document
      ) as vscode.CodeLens[];

      chai.assert.equal(codelens.length, 0);
    });

    it("Do not show codelens for if Teams manifest not exist", async () => {
      const openApiObject = {
        openapi: "3.0",
      };
      const text = JSON.stringify(openApiObject);
      const document = {
        fileName: "openapi.yaml",
        getText: () => {
          return text;
        },
        positionAt: () => {
          return new vscode.Position(0, 0);
        },
        lineAt: () => {
          return {
            lineNumber: 0,
            text,
          };
        },
      } as any as vscode.TextDocument;

      sandbox.stub(fs, "existsSync").returns(false);
      sandbox
        .stub(globalVariables, "workspaceUri")
        .value(vscode.Uri.parse(path.resolve(__dirname, "unknown")));
      const apiPluginCodelensProvider = new ApiPluginCodeLensProvider();
      const codelens: vscode.CodeLens[] = apiPluginCodelensProvider.provideCodeLenses(
        document
      ) as vscode.CodeLens[];

      chai.assert.equal(codelens.length, 0);
    });

    it("Do not show codelens for if not API plugin project", async () => {
      const manifest = new TeamsAppManifest();
      manifest.copilotExtensions = {};
      const openApiObject = {
        openapi: "3.0",
      };
      const text = JSON.stringify(openApiObject);
      const document = {
        fileName: "openapi.yaml",
        getText: () => {
          return text;
        },
        positionAt: () => {
          return new vscode.Position(0, 0);
        },
        lineAt: () => {
          return {
            lineNumber: 0,
            text,
          };
        },
      } as any as vscode.TextDocument;

      sandbox.stub(fs, "existsSync").returns(true);
      sandbox.stub(fs, "readFileSync").returns(JSON.stringify(manifest));
      sandbox
        .stub(globalVariables, "workspaceUri")
        .value(vscode.Uri.parse(path.resolve(__dirname, "unknown")));
      const apiPluginCodelensProvider = new ApiPluginCodeLensProvider();
      const codelens: vscode.CodeLens[] = apiPluginCodelensProvider.provideCodeLenses(
        document
      ) as vscode.CodeLens[];

      chai.assert.equal(codelens.length, 0);
    });
  });

  describe("m365agents.yml CodeLensProvider", () => {
    const sandbox = sinon.createSandbox();

    afterEach(() => {
      sandbox.restore();
    });

    it("should work with correct m365agents.yml", async () => {
      const text = `
version: 1.1.0

provision:
  provision: 1 // this line shouldn't have codelens
deploy:
  publish: 2 // this line shouldn't have codelens
publish:
  ccc: 3`;
      const document = {
        fileName: "m365agents.yml",
        getText: () => {
          return text;
        },
        positionAt: () => {
          return new vscode.Position(0, 0);
        },
        lineAt: () => {
          return {
            lineNumber: 0,
            text: text,
          };
        },
      } as any as vscode.TextDocument;

      const provider = new TeamsAppYamlCodeLensProvider();
      const codelens: vscode.CodeLens[] = provider.provideCodeLenses(document) as vscode.CodeLens[];

      chai.assert.equal(codelens.length, 3);
      chai.expect(codelens[0].command?.command).eq("fx-extension.provision");
      chai.expect(codelens[0].command?.arguments).deep.eq([TelemetryTriggerFrom.CodeLens]);
      chai.expect(codelens[1].command?.command).eq("fx-extension.deploy");
      chai.expect(codelens[1].command?.arguments).deep.eq([TelemetryTriggerFrom.CodeLens]);
      chai.expect(codelens[2].command?.command).eq("fx-extension.publish");
      chai.expect(codelens[2].command?.arguments).deep.eq([TelemetryTriggerFrom.CodeLens]);
    });
  });

  describe("manifest*.xml CodeLensProvider", () => {
    const sandbox = sinon.createSandbox();

    afterEach(() => {
      sandbox.restore();
    });

    it("should work with correct manifest.xml", async () => {
      const text = `
    <OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">
       <Id>518f978a-6cf4-46f8-8f1e-10881613fe54</Id>
        <Version>1.0.0.0</Version>
        <ProviderName>Contoso</ProviderName>
        <DefaultLocale>en-US</DefaultLocale>
    </OfficeApp>
    `;

      const document = {
        fileName: "manifest-localhost.yml",
        getText: () => {
          return text;
        },
        positionAt: () => {
          return new vscode.Position(0, 0);
        },
        lineAt: () => {
          return {
            lineNumber: 0,
            text: text,
          };
        },
      } as any as vscode.TextDocument;

      const provider = new OfficeDevManifestCodeLensProvider();
      const codelens: vscode.CodeLens[] = provider.provideCodeLenses(document) as vscode.CodeLens[];
      chai.assert.equal(codelens.length, 1);
      chai.expect(codelens[0].command?.command).eq("fx-extension.generateManifestGUID");
      chai
        .expect(codelens[0].command?.arguments?.[0])
        .deep.eq("518f978a-6cf4-46f8-8f1e-10881613fe54");
    });
  });

  describe("OneDriveSharePointCodeLensProvider", () => {
    const sandbox = sinon.createSandbox();

    afterEach(() => {
      sandbox.restore();
    });

    it("should not provide codelens when feature flag is disabled", async () => {
      sandbox.stub(featureFlagManager, "getBooleanValue").returns(false);
      const document = {
        fileName: "agent.json",
        getText: () => {
          return JSON.stringify({
            site_id: "test-site-id",
            web_id: "test-web-id",
            list_id: "test-list-id",
            unique_id: "test-unique-id",
          });
        },
        positionAt: () => {
          return new vscode.Position(0, 0);
        },
        lineAt: () => {
          return {
            lineNumber: 0,
            text: "test",
          };
        },
      } as any as vscode.TextDocument;

      const provider = new OneDriveSharePointCodeLensProvider();
      const codelens = await provider.provideCodeLenses(document);

      chai.assert.equal((codelens as vscode.CodeLens[]).length, 0);
    });

    it("should not provide codelens when manifest file does not exist", async () => {
      sandbox.stub(featureFlagManager, "getBooleanValue").returns(true);
      sandbox.stub(fs, "existsSync").returns(false);
      sandbox
        .stub(globalVariables, "workspaceUri")
        .value(vscode.Uri.parse(path.resolve(__dirname, "unknown")));

      const document = {
        fileName: "agent.json",
        getText: () => {
          return JSON.stringify({
            site_id: "test-site-id",
            web_id: "test-web-id",
            list_id: "test-list-id",
            unique_id: "test-unique-id",
          });
        },
        positionAt: () => {
          return new vscode.Position(0, 0);
        },
        lineAt: () => {
          return {
            lineNumber: 0,
            text: "test",
          };
        },
      } as any as vscode.TextDocument;

      const provider = new OneDriveSharePointCodeLensProvider();
      const codelens = await provider.provideCodeLenses(document);

      chai.assert.equal((codelens as vscode.CodeLens[]).length, 0);
    });

    it("should not provide codelens when not a copilot project", async () => {
      sandbox.stub(featureFlagManager, "getBooleanValue").returns(true);
      sandbox.stub(fs, "existsSync").returns(true);
      const manifest = new TeamsAppManifest();
      manifest.copilotAgents = {};
      sandbox.stub(fs, "readFileSync").returns(JSON.stringify(manifest));
      sandbox
        .stub(globalVariables, "workspaceUri")
        .value(vscode.Uri.parse(path.resolve(__dirname, "unknown")));

      const document = {
        fileName: "agent.json",
        getText: () => {
          return JSON.stringify({
            site_id: "test-site-id",
            web_id: "test-web-id",
            list_id: "test-list-id",
            unique_id: "test-unique-id",
          });
        },
        positionAt: () => {
          return new vscode.Position(0, 0);
        },
        lineAt: () => {
          return {
            lineNumber: 0,
            text: "test",
          };
        },
      } as any as vscode.TextDocument;

      const provider = new OneDriveSharePointCodeLensProvider();
      const codelens = await provider.provideCodeLenses(document);

      chai.assert.equal((codelens as vscode.CodeLens[]).length, 0);
    });

    it("should provide codelens for SharePoint IDs", async () => {
      sandbox.stub(featureFlagManager, "getBooleanValue").returns(true);
      sandbox.stub(fs, "existsSync").returns(true);
      const manifest = {
        copilotAgents: {
          declarativeAgents: [
            {
              id: "test-agent",
              file: "agent.json",
            },
          ],
        },
        capabilities: ["copilotGpt"],
      };
      sandbox.stub(fs, "readFileSync").returns(JSON.stringify(manifest));
      sandbox
        .stub(globalVariables, "workspaceUri")
        .value(vscode.Uri.parse(path.resolve(__dirname, "unknown")));

      const document = {
        fileName: "agent.json",
        getText: () => {
          return JSON.stringify({
            site_id: "test-site-id",
            web_id: "test-web-id",
            list_id: "test-list-id",
            unique_id: "test-unique-id",
          });
        },
        positionAt: () => {
          return new vscode.Position(0, 0);
        },
        lineAt: () => {
          return {
            lineNumber: 0,
            text: "{",
          };
        },
      } as any as vscode.TextDocument;

      const provider = new OneDriveSharePointCodeLensProvider();
      const codelens = await provider.provideCodeLenses(document);

      chai.assert.equal((codelens as vscode.CodeLens[]).length, 1);
      chai.assert.isTrue((codelens as vscode.CodeLens[])[0] instanceof SharePointIdCodeLens);
    });

    it("should resolve codelens with item details", async () => {
      const range = new vscode.Range(new vscode.Position(0, 0), new vscode.Position(0, 1));
      const lens = new SharePointIdCodeLens(
        JSON.stringify({
          site_id: "test-site-id",
          unique_id: "test-unique-id",
        }),
        range
      );

      const mockCore = {
        getODSPItemDetails: sinon.stub(),
      };
      mockCore.getODSPItemDetails.resolves(
        ok({
          id: "test-id",
          name: "Test Item",
          webUrl: "https://test.sharepoint.com",
          type: "file",
          lastModifiedDateTime: new Date().toISOString(),
          createdDateTime: new Date().toISOString(),
          size: 0,
        })
      );
      sandbox.stub(globalVariables, "core").value(mockCore);

      const provider = new OneDriveSharePointCodeLensProvider();
      const resolvedLens = await provider.resolveCodeLens(
        lens,
        new vscode.CancellationTokenSource().token
      );

      chai.assert.isDefined(resolvedLens.command);
      chai.assert.equal(resolvedLens.command?.command, "fx-extension.openOneDriveSharePointUrl");
      chai.assert.isTrue(resolvedLens.command?.title.includes("Test Item"));
      chai.assert.deepEqual(resolvedLens.command?.arguments, ["https://test.sharepoint.com"]);
    });

    it("should handle error when resolving codelens", async () => {
      const range = new vscode.Range(new vscode.Position(0, 0), new vscode.Position(0, 1));
      const lens = new SharePointIdCodeLens(
        JSON.stringify({
          site_id: "test-site-id",
          unique_id: "test-unique-id",
        }),
        range
      );

      const errorOptions: SystemErrorOptions = {
        source: "test",
        name: "TestError",
        message: "Test error",
        error: new Error("Test error"),
      };
      const error = new SystemError(errorOptions);
      const mockCore = {
        getODSPItemDetails: sinon.stub(),
      };
      mockCore.getODSPItemDetails.resolves(err(error));
      sandbox.stub(globalVariables, "core").value(mockCore);

      const provider = new OneDriveSharePointCodeLensProvider();
      const resolvedLens = await provider.resolveCodeLens(
        lens,
        new vscode.CancellationTokenSource().token
      );

      chai.assert.isDefined(resolvedLens.command);
      chai.assert.equal(resolvedLens.command?.command, "");
      chai.assert.isTrue(resolvedLens.command?.title.includes("Test error"));
    });

    it("should handle missing required IDs when resolving codelens", async () => {
      const range = new vscode.Range(new vscode.Position(0, 0), new vscode.Position(0, 1));
      const lens = new SharePointIdCodeLens(
        JSON.stringify({
          // missing site_id
          unique_id: "test-unique-id",
        }),
        range
      );

      const provider = new OneDriveSharePointCodeLensProvider();
      const resolvedLens = await provider.resolveCodeLens(
        lens,
        new vscode.CancellationTokenSource().token
      );

      chai.assert.isDefined(resolvedLens.command);
      chai.assert.equal(resolvedLens.command?.command, "");
      chai.assert.isTrue(resolvedLens.command?.title.includes("Missing required SharePoint IDs"));
    });

    it("should handle missing unique ID when resolving codelens", async () => {
      const range = new vscode.Range(new vscode.Position(0, 0), new vscode.Position(0, 1));
      const lens = new SharePointIdCodeLens(
        JSON.stringify({
          site_id: "test-site-id",
        }),
        range
      );
      const mockCore = {
        getODSPItemDetails: sinon.stub(),
      };
      mockCore.getODSPItemDetails.resolves(
        ok({
          id: "test-id",
          name: "Test Item",
          webUrl: "https://test.sharepoint.com",
          type: "file",
          lastModifiedDateTime: new Date().toISOString(),
          createdDateTime: new Date().toISOString(),
          size: 0,
        })
      );
      sandbox.stub(globalVariables, "core").value(mockCore);

      const provider = new OneDriveSharePointCodeLensProvider();
      const resolvedLens = await provider.resolveCodeLens(
        lens,
        new vscode.CancellationTokenSource().token
      );

      chai.assert.isDefined(resolvedLens.command);
      chai.assert.equal(resolvedLens.command?.command, "fx-extension.openOneDriveSharePointUrl");
      chai.assert.isTrue(resolvedLens.command?.title.includes("Test Item"));
      chai.assert.deepEqual(resolvedLens.command?.arguments, ["https://test.sharepoint.com"]);
    });

    it("should return unmodified lens when not a SharePointIdCodeLens", async () => {
      const range = new vscode.Range(new vscode.Position(0, 0), new vscode.Position(0, 1));
      const lens = new vscode.CodeLens(range, {
        title: "Original command",
        command: "test.command",
      });

      const provider = new OneDriveSharePointCodeLensProvider();
      const resolvedLens = await provider.resolveCodeLens(
        lens,
        new vscode.CancellationTokenSource().token
      );

      // Explicitly verify that the original lens is returned without modification
      chai.assert.strictEqual(resolvedLens, lens);
      chai.assert.deepEqual(resolvedLens.command, {
        title: "Original command",
        command: "test.command",
      });
      chai.assert.deepEqual(resolvedLens.range, range);
    });

    it("should handle undefined SharePoint IDs when resolving codelens", async () => {
      const range = new vscode.Range(new vscode.Position(0, 0), new vscode.Position(0, 1));
      const lens = new SharePointIdCodeLens(
        JSON.stringify({
          // both site_id and unique_id are undefined
        }),
        range
      );

      const provider = new OneDriveSharePointCodeLensProvider();
      const resolvedLens = await provider.resolveCodeLens(
        lens,
        new vscode.CancellationTokenSource().token
      );

      chai.assert.isDefined(resolvedLens.command);
      chai.assert.equal(resolvedLens.command?.command, "");
      chai.assert.isTrue(resolvedLens.command?.title.includes("Missing required SharePoint IDs"));
    });
  });

  describe("DeclarativeAgentSensitivityLabelCodeLensProvider", () => {
    const sandbox = sinon.createSandbox();
    setTools(new MockTools());
    afterEach(() => {
      sandbox.restore();
    });

    it("should not provide codelens when projectPath is not available", async () => {
      // Stub getSystemInputs to return an object without projectPath
      sandbox.stub(require("../../src/utils/systemEnvUtils"), "getSystemInputs").returns({
        // No projectPath included
      });

      const document = {
        fileName: "agent.json",
        getText: () => {
          return JSON.stringify({
            type: "declarative",
            sensitivity_label: { id: "internal-only" },
          });
        },
        positionAt: () => {
          return new vscode.Position(0, 0);
        },
        lineAt: () => {
          return {
            lineNumber: 0,
            text: "test",
          };
        },
        uri: {
          fsPath: path.join("unknown", "appPackage", "agent.json"),
        },
      } as any as vscode.TextDocument;

      const provider = new DeclarativeAgentSensitivityLabelCodeLensProvider();
      const codelens = await provider.provideCodeLenses(document);

      chai.assert.equal((codelens as vscode.CodeLens[]).length, 0);
    });

    it("should not provide codelens when manifest file does not exist", async () => {
      sandbox.stub(fs, "pathExistsSync").returns(false);
      sandbox
        .stub(globalVariables, "workspaceUri")
        .value(vscode.Uri.parse(path.resolve(__dirname, "unknown")));

      const document = {
        fileName: "agent.json",
        getText: () => {
          return JSON.stringify({
            type: "declarative",
            sensitivity_label: { id: "internal-only" },
          });
        },
        positionAt: () => {
          return new vscode.Position(0, 0);
        },
        lineAt: () => {
          return {
            lineNumber: 0,
            text: "test",
          };
        },
        uri: {
          fsPath: path.join(__dirname, "unknown", "appPackage", "agent.json"),
        },
      } as any as vscode.TextDocument;

      const provider = new DeclarativeAgentSensitivityLabelCodeLensProvider();
      const codelens = await provider.provideCodeLenses(document);

      chai.assert.equal((codelens as vscode.CodeLens[]).length, 0);
    });

    it("should not provide codelens when document path doesn't match declarative agent file path", async () => {
      sandbox.stub(fs, "pathExistsSync").returns(true);
      const projectPath = path.join(__dirname, "unknown");
      const agentPath = "agent.json";

      const manifest = {
        copilotAgents: {
          declarativeAgents: [
            {
              id: "test-agent",
              file: agentPath,
            },
          ],
        },
      };
      sandbox.stub(fs, "readFileSync").returns(JSON.stringify(manifest));
      sandbox.stub(globalVariables, "workspaceUri").value(vscode.Uri.parse(projectPath));

      // Create a document with a different path than the agent file path
      const document = {
        fileName: "different-agent.json",
        getText: () => {
          return JSON.stringify({
            type: "declarative",
            sensitivity_label: { id: "internal-only" },
          });
        },
        positionAt: () => {
          return new vscode.Position(0, 0);
        },
        lineAt: () => {
          return {
            lineNumber: 0,
            text: "test",
          };
        },
        uri: {
          fsPath: path.join(projectPath, "appPackage", "different-agent.json"),
        },
      } as any as vscode.TextDocument;

      const provider = new DeclarativeAgentSensitivityLabelCodeLensProvider();
      const codelens = await provider.provideCodeLenses(document);

      chai.assert.equal((codelens as vscode.CodeLens[]).length, 0);
    });

    it("should not provide codelens when manifest does not contain declarative agent", async () => {
      sandbox.stub(fs, "pathExistsSync").returns(true);
      const projectPath = path.join(__dirname, "unknown");
      const agentPath = "agent.json";

      const manifest = {};
      sandbox.stub(fs, "readFileSync").returns(JSON.stringify(manifest));
      sandbox.stub(globalVariables, "workspaceUri").value(vscode.Uri.parse(projectPath));

      // Create a document with a different path than the agent file path
      const document = {
        fileName: agentPath,
        getText: () => {
          return JSON.stringify({
            type: "declarative",
            sensitivity_label: { id: "internal-only" },
          });
        },
        positionAt: () => {
          return new vscode.Position(0, 0);
        },
        lineAt: () => {
          return {
            lineNumber: 0,
            text: "test",
          };
        },
        uri: {
          fsPath: path.join(projectPath, "appPackage", "different-agent.json"),
        },
      } as any as vscode.TextDocument;

      const provider = new DeclarativeAgentSensitivityLabelCodeLensProvider();
      const codelens = await provider.provideCodeLenses(document);

      chai.assert.equal((codelens as vscode.CodeLens[]).length, 0);
    });

    it("should provide sensitivity label codelens when label exists", async () => {
      sandbox.stub(fs, "pathExistsSync").returns(true);
      const projectPath = path.join(__dirname, "unknown");
      const agentPath = "agent.json";
      const absoluteAgentPath = path.join(projectPath, "appPackage", agentPath);

      const manifest = {
        copilotAgents: {
          declarativeAgents: [
            {
              id: "test-agent",
              file: agentPath,
            },
          ],
        },
      };
      sandbox.stub(fs, "readFileSync").returns(JSON.stringify(manifest));
      sandbox.stub(globalVariables, "workspaceUri").value(vscode.Uri.parse(projectPath));

      // Setup login and token provider mocks
      const mockTokenProvider = {
        m365TokenProvider: {
          getStatus: () => {
            return Promise.resolve(
              ok({
                status: "SignedIn",
                token: "mock_token",
                accountInfo: {
                  unique_name: "test@test.com",
                  tid: "test-tenant-id",
                },
              })
            );
          },
        },
      };
      sandbox.stub(tools, "tokenProvider").value(mockTokenProvider);

      // Mock sensitivity labels API response
      const mockGraphClient = {
        listSensitivityLabels: () => {
          return Promise.resolve({
            isOk: () => true,
            value: [
              {
                id: "test-label-id",
                displayName: "Test Label",
              },
            ],
          });
        },
      };
      sandbox
        .stub(GraphClient.prototype, "listSensitivityLabels")
        .callsFake(mockGraphClient.listSensitivityLabels as any);

      const document = {
        fileName: "agent.json",
        getText: () => {
          return JSON.stringify({
            type: "declarative",
            sensitivity_label: {
              id: "test-label-id",
            },
          });
        },
        positionAt: () => {
          return new vscode.Position(0, 0);
        },
        lineAt: () => {
          return {
            lineNumber: 0,
            text: '"sensitivity_label": { "id": "test-label-id" }',
          };
        },
        uri: {
          fsPath: absoluteAgentPath,
        },
      } as any as vscode.TextDocument;
      sandbox.stub(copilotGptManifestUtils, "readDeclarativeAgentManifestFile").resolves(
        ok({
          type: "declarative",
          sensitivity_label: {
            id: "test-label-id",
          },
        } as any)
      );
      const provider = new DeclarativeAgentSensitivityLabelCodeLensProvider();
      const codelens = await provider.provideCodeLenses(document);

      chai.assert.equal((codelens as vscode.CodeLens[]).length, 1);
      const lens = codelens[0] as vscode.CodeLens;
      chai.assert.equal(lens.command?.command, "fx-extension.setSensitivityLabel");
      chai.assert.isTrue(lens.command?.title.includes("Test Label"));
    });

    it("should show login codelens when user not logged in", async () => {
      sandbox.stub(fs, "pathExistsSync").returns(true);
      const projectPath = path.join(__dirname, "unknown");
      const agentPath = "agent.json";
      const absoluteAgentPath = path.join(projectPath, "appPackage", agentPath);

      const manifest = {
        copilotAgents: {
          declarativeAgents: [
            {
              id: "test-agent",
              file: agentPath,
            },
          ],
        },
      };
      sandbox.stub(fs, "readFileSync").returns(JSON.stringify(manifest));
      sandbox.stub(globalVariables, "workspaceUri").value(vscode.Uri.parse(projectPath));

      // Setup login status as not signed in
      const mockTokenProvider = {
        m365TokenProvider: {
          getStatus: () => {
            return Promise.resolve(
              ok({
                isOk: () => true,
                value: {
                  status: "NotSignedIn",
                },
              })
            );
          },
        },
      };
      sandbox.stub(tools, "tokenProvider").value(mockTokenProvider);
      sandbox.stub(copilotGptManifestUtils, "readDeclarativeAgentManifestFile").resolves(
        ok({
          type: "declarative",
          sensitivity_label: {
            id: "test-label-id",
          },
        } as any)
      );
      const document = {
        fileName: "agent.json",
        getText: () => {
          return JSON.stringify({
            type: "declarative",
            sensitivity_label: {
              id: "test-label-id",
            },
          });
        },
        positionAt: () => {
          return new vscode.Position(0, 0);
        },
        lineAt: () => {
          return {
            lineNumber: 0,
            text: '"sensitivity_label": "test-label-id"',
          };
        },
        uri: {
          fsPath: absoluteAgentPath,
        },
      } as any as vscode.TextDocument;

      const provider = new DeclarativeAgentSensitivityLabelCodeLensProvider();
      const codelens = await provider.provideCodeLenses(document);

      chai.assert.equal((codelens as vscode.CodeLens[]).length, 1);
      const lens = codelens[0] as vscode.CodeLens;
      chai.assert.equal(lens.command?.command, "fx-extension.m365PreAuth");
    });

    it("should provide add sensitivity label codelens when label not exists", async () => {
      sandbox.stub(fs, "pathExistsSync").returns(true);
      const projectPath = path.join(__dirname, "unknown");
      const agentPath = "agent.json";
      const absoluteAgentPath = path.join(projectPath, "appPackage", agentPath);

      const manifest = {
        copilotAgents: {
          declarativeAgents: [
            {
              id: "test-agent",
              file: agentPath,
            },
          ],
        },
      };
      sandbox.stub(fs, "readFileSync").returns(JSON.stringify(manifest));
      sandbox.stub(globalVariables, "workspaceUri").value(vscode.Uri.parse(projectPath));

      const document = {
        fileName: "agent.json",
        getText: () => {
          return JSON.stringify({
            type: "declarative",
            // No sensitivity_label field
          });
        },
        positionAt: () => {
          return new vscode.Position(0, 0);
        },
        lineAt: () => {
          return {
            lineNumber: 0,
            text: "{",
          };
        },
        uri: {
          fsPath: absoluteAgentPath,
        },
      } as any as vscode.TextDocument;

      const provider = new DeclarativeAgentSensitivityLabelCodeLensProvider();
      const codelens = await provider.provideCodeLenses(document);

      chai.assert.equal((codelens as vscode.CodeLens[]).length, 1);
      const lens = codelens[0] as vscode.CodeLens;
      chai.assert.equal(lens.command?.command, "fx-extension.setSensitivityLabel");
    });
  });

  describe("WorkspaceMCPConfigCodeLensProvider", () => {
    const sandbox = sinon.createSandbox();

    afterEach(() => {
      sandbox.restore();
    });

    it("should provide codelens for valid MCP config with single server", async () => {
      const mcpConfigText = `{
  // MCP Server Configuration
  "servers": {
    "my-server": {
      "command": "node",
      "args": ["server.js"],
      "env": {
        "API_KEY": "test-key"
      }
    }
  }
}`;

      const document = {
        fileName: "mcp-config.json",
        getText: () => {
          return mcpConfigText;
        },
        positionAt: (offset: number) => {
          const lines = mcpConfigText.substring(0, offset).split("\n");
          return new vscode.Position(lines.length - 1, lines[lines.length - 1].length);
        },
        lineAt: () => {
          return {
            lineNumber: 0,
            text: mcpConfigText,
          };
        },
      } as unknown as vscode.TextDocument;

      const provider = new WorkspaceMCPConfigCodeLensProvider();
      const codelens: vscode.CodeLens[] = provider.provideCodeLenses(document) as vscode.CodeLens[];

      chai.assert.equal(codelens.length, 1);
      chai.expect(codelens[0].command?.title).to.equal("⚡ ATK: Fetch action from MCP");
      chai.expect(codelens[0].command?.command).to.equal("fx-extension.updateActionWithMCP");
      chai.expect(codelens[0].command?.arguments).to.deep.equal([
        {
          serverName: "my-server",
          serverConfig: {
            command: "node",
            args: ["server.js"],
            env: {
              API_KEY: "test-key",
            },
          },
        },
      ]);
    });

    it("should provide codelens for multiple servers", async () => {
      const mcpConfigText = `{
  "servers": {
    "server-one": {
      "command": "python",
      "args": ["app.py"]
    },
    "server-two": {
      "command": "node", 
      "args": ["index.js"]
    }
  }
}`;

      const document = {
        fileName: "mcp-config.json",
        getText: () => {
          return mcpConfigText;
        },
        positionAt: (offset: number) => {
          const lines = mcpConfigText.substring(0, offset).split("\n");
          return new vscode.Position(lines.length - 1, lines[lines.length - 1].length);
        },
        lineAt: () => {
          return {
            lineNumber: 0,
            text: mcpConfigText,
          };
        },
      } as unknown as vscode.TextDocument;

      const provider = new WorkspaceMCPConfigCodeLensProvider();
      const codelens: vscode.CodeLens[] = provider.provideCodeLenses(document) as vscode.CodeLens[];

      chai.assert.equal(codelens.length, 2);

      chai.expect(codelens[0].command?.title).to.equal("⚡ ATK: Fetch action from MCP");
      chai.expect(codelens[0].command?.command).to.equal("fx-extension.updateActionWithMCP");
      chai.expect(codelens[0].command?.arguments).to.deep.equal([
        {
          serverName: "server-one",
          serverConfig: {
            command: "python",
            args: ["app.py"],
          },
        },
      ]);

      chai.expect(codelens[1].command?.title).to.equal("⚡ ATK: Fetch action from MCP");
      chai.expect(codelens[1].command?.command).to.equal("fx-extension.updateActionWithMCP");
      chai.expect(codelens[1].command?.arguments).to.deep.equal([
        {
          serverName: "server-two",
          serverConfig: {
            command: "node",
            args: ["index.js"],
          },
        },
      ]);
    });

    it("should handle JSON with comments (JSONC)", async () => {
      const mcpConfigText = `{
  // This is a comment
  "servers": {
    // Another comment
    "test-server": {
      "command": "npm",
      "args": ["start"] // inline comment
    }
  }
  /* Multi-line
     comment */
}`;

      const document = {
        fileName: "mcp-config.json",
        getText: () => {
          return mcpConfigText;
        },
        positionAt: (offset: number) => {
          const lines = mcpConfigText.substring(0, offset).split("\n");
          return new vscode.Position(lines.length - 1, lines[lines.length - 1].length);
        },
        lineAt: () => {
          return {
            lineNumber: 0,
            text: mcpConfigText,
          };
        },
      } as unknown as vscode.TextDocument;

      const provider = new WorkspaceMCPConfigCodeLensProvider();
      const codelens: vscode.CodeLens[] = provider.provideCodeLenses(document) as vscode.CodeLens[];

      chai.assert.equal(codelens.length, 1);
      chai.expect(codelens[0].command?.title).to.equal("⚡ ATK: Fetch action from MCP");
      chai.expect(codelens[0].command?.command).to.equal("fx-extension.updateActionWithMCP");
      chai.expect(codelens[0].command?.arguments).to.deep.equal([
        {
          serverName: "test-server",
          serverConfig: {
            command: "npm",
            args: ["start"],
          },
        },
      ]);
    });

    it("should return empty array when no servers property exists", async () => {
      const mcpConfigText = `{
  "otherProperty": {
    "value": "test"
  }
}`;

      const document = {
        fileName: "mcp-config.json",
        getText: () => {
          return mcpConfigText;
        },
        positionAt: () => {
          return new vscode.Position(0, 0);
        },
        lineAt: () => {
          return {
            lineNumber: 0,
            text: mcpConfigText,
          };
        },
      } as unknown as vscode.TextDocument;

      const provider = new WorkspaceMCPConfigCodeLensProvider();
      const codelens: vscode.CodeLens[] = provider.provideCodeLenses(document) as vscode.CodeLens[];

      chai.assert.equal(codelens.length, 0);
    });

    it("should return empty array when servers is not an object", async () => {
      const mcpConfigText = `{
  "servers": "not-an-object"
}`;

      const document = {
        fileName: "mcp-config.json",
        getText: () => {
          return mcpConfigText;
        },
        positionAt: () => {
          return new vscode.Position(0, 0);
        },
        lineAt: () => {
          return {
            lineNumber: 0,
            text: mcpConfigText,
          };
        },
      } as unknown as vscode.TextDocument;

      const provider = new WorkspaceMCPConfigCodeLensProvider();
      const codelens: vscode.CodeLens[] = provider.provideCodeLenses(document) as vscode.CodeLens[];

      chai.assert.equal(codelens.length, 0);
    });

    it("should return empty array when servers object is empty", async () => {
      const mcpConfigText = `{
  "servers": {}
}`;

      const document = {
        fileName: "mcp-config.json",
        getText: () => {
          return mcpConfigText;
        },
        positionAt: () => {
          return new vscode.Position(0, 0);
        },
        lineAt: () => {
          return {
            lineNumber: 0,
            text: mcpConfigText,
          };
        },
      } as unknown as vscode.TextDocument;

      const provider = new WorkspaceMCPConfigCodeLensProvider();
      const codelens: vscode.CodeLens[] = provider.provideCodeLenses(document) as vscode.CodeLens[];

      chai.assert.equal(codelens.length, 0);
    });

    it("should handle empty document", async () => {
      const mcpConfigText = "";

      const document = {
        fileName: "mcp-config.json",
        getText: () => {
          return mcpConfigText;
        },
        positionAt: () => {
          return new vscode.Position(0, 0);
        },
        lineAt: () => {
          return {
            lineNumber: 0,
            text: mcpConfigText,
          };
        },
      } as unknown as vscode.TextDocument;

      const provider = new WorkspaceMCPConfigCodeLensProvider();
      const codelens: vscode.CodeLens[] = provider.provideCodeLenses(document) as vscode.CodeLens[];

      chai.assert.equal(codelens.length, 0);
    });

    it("should handle document with only whitespace", async () => {
      const mcpConfigText = "   \n   \t   \n   ";

      const document = {
        fileName: "mcp-config.json",
        getText: () => {
          return mcpConfigText;
        },
        positionAt: () => {
          return new vscode.Position(0, 0);
        },
        lineAt: () => {
          return {
            lineNumber: 0,
            text: mcpConfigText,
          };
        },
      } as unknown as vscode.TextDocument;

      const provider = new WorkspaceMCPConfigCodeLensProvider();
      const codelens: vscode.CodeLens[] = provider.provideCodeLenses(document) as vscode.CodeLens[];

      chai.assert.equal(codelens.length, 0);
    });

    it("should calculate correct range for server name", async () => {
      const mcpConfigText = `{
  "servers": {
    "my-test-server": {
      "command": "node"
    }
  }
}`;

      const document = {
        fileName: "mcp-config.json",
        getText: () => {
          return mcpConfigText;
        },
        positionAt: (offset: number) => {
          // Find the line and character position
          const beforeOffset = mcpConfigText.substring(0, offset);
          const lines = beforeOffset.split("\n");
          const line = lines.length - 1;
          const character = lines[line].length;
          return new vscode.Position(line, character);
        },
        lineAt: () => {
          return {
            lineNumber: 0,
            text: mcpConfigText,
          };
        },
      } as unknown as vscode.TextDocument;

      const provider = new WorkspaceMCPConfigCodeLensProvider();
      const codelens: vscode.CodeLens[] = provider.provideCodeLenses(document) as vscode.CodeLens[];

      chai.assert.equal(codelens.length, 1);

      // The range should cover the server name with quotes
      const expectedStartPos = mcpConfigText.indexOf('"my-test-server"');
      chai.assert.isTrue(expectedStartPos !== -1, "Server name should be found in the text");

      // Check that the range properties are set correctly
      chai.assert.equal(codelens[0].range.start.line, codelens[0].range.end.line);
      chai.assert.equal(
        codelens[0].range.end.character - codelens[0].range.start.character,
        "my-test-server".length + 2
      );
    });

    it("should handle server names with special characters", async () => {
      const mcpConfigText = `{
  "servers": {
    "server-with-dashes_and_underscores.123": {
      "command": "python",
      "args": ["-m", "server"]
    },
    "@scoped/server-name": {
      "command": "npm",
      "args": ["start"]
    }
  }
}`;

      const document = {
        fileName: "mcp-config.json",
        getText: () => {
          return mcpConfigText;
        },
        positionAt: (offset: number) => {
          const lines = mcpConfigText.substring(0, offset).split("\n");
          return new vscode.Position(lines.length - 1, lines[lines.length - 1].length);
        },
        lineAt: () => {
          return {
            lineNumber: 0,
            text: mcpConfigText,
          };
        },
      } as unknown as vscode.TextDocument;

      const provider = new WorkspaceMCPConfigCodeLensProvider();
      const codelens: vscode.CodeLens[] = provider.provideCodeLenses(document) as vscode.CodeLens[];

      chai.assert.equal(codelens.length, 2);

      chai.expect(codelens[0].command?.arguments).to.deep.equal([
        {
          serverName: "server-with-dashes_and_underscores.123",
          serverConfig: {
            command: "python",
            args: ["-m", "server"],
          },
        },
      ]);

      chai.expect(codelens[1].command?.arguments).to.deep.equal([
        {
          serverName: "@scoped/server-name",
          serverConfig: {
            command: "npm",
            args: ["start"],
          },
        },
      ]);
    });

    it("should handle null or undefined server configurations", async () => {
      const mcpConfigText = `{
  "servers": {
    "null-server": null,
    "undefined-server": {
      "command": "node"
    }
  }
}`;

      const document = {
        fileName: "mcp-config.json",
        getText: () => {
          return mcpConfigText;
        },
        positionAt: (offset: number) => {
          const lines = mcpConfigText.substring(0, offset).split("\n");
          return new vscode.Position(lines.length - 1, lines[lines.length - 1].length);
        },
        lineAt: () => {
          return {
            lineNumber: 0,
            text: mcpConfigText,
          };
        },
      } as unknown as vscode.TextDocument;

      const provider = new WorkspaceMCPConfigCodeLensProvider();
      const codelens: vscode.CodeLens[] = provider.provideCodeLenses(document) as vscode.CodeLens[];

      chai.assert.equal(codelens.length, 2);

      chai.expect(codelens[0].command?.arguments).to.deep.equal([
        {
          serverName: "null-server",
          serverConfig: null,
        },
      ]);

      chai.expect(codelens[1].command?.arguments).to.deep.equal([
        {
          serverName: "undefined-server",
          serverConfig: {
            command: "node",
          },
        },
      ]);
    });
  });
});
