import { TeamsAppManifest, ok } from "@microsoft/teamsfx-api";
import { envUtil } from "@microsoft/teamsfx-core";
import * as chai from "chai";
import fs from "fs-extra";
import * as sinon from "sinon";
import * as vscode from "vscode";
import {
  AadAppTemplateCodeLensProvider,
  ApiPluginCodeLensProvider,
  CopilotPluginCodeLensProvider,
  CryptoCodeLensProvider,
  ManifestTemplateCodeLensProvider,
  OfficeDevManifestCodeLensProvider,
  PermissionsJsonFileCodeLensProvider,
  PlaceholderCodeLens,
  TeamsAppYamlCodeLensProvider,
} from "../../src/codeLensProvider";
import * as globalVariables from "../../src/globalVariables";
import { TelemetryTriggerFrom } from "../../src/telemetry/extTelemetryEvents";
import path = require("path");

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
          return "{name: 'test'}";
        },
      };

      const aadProvider = new AadAppTemplateCodeLensProvider();
      const res = await aadProvider.provideCodeLenses(document);
      chai.assert.isTrue(
        res != null && res[0].command!.command === "fx-extension.openPreviewAadFile"
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

  describe("teamsapp.yml CodeLensProvider", () => {
    const sandbox = sinon.createSandbox();

    afterEach(() => {
      sandbox.restore();
    });

    it("should work with correct teamsapp.yml", async () => {
      const text = `
version: 1.1.0

provision:
  provision: 1 // this line shouldn't have codelens
deploy:
  publish: 2 // this line shouldn't have codelens
publish:
  ccc: 3`;
      const document = {
        fileName: "teamsapp.yml",
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
});
