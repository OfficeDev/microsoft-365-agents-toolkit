// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { err, ok, TeamsAppManifest } from "@microsoft/teamsfx-api";
import AdmZip from "adm-zip";
import chai from "chai";
import fs from "fs-extra";
import "mocha";
import * as sinon from "sinon";
import { v4 as uuid } from "uuid";
import { CopilotAgentPublishDriver } from "../../../../src/component/driver/copilotAgent/publish";
import { CopilotAgentPublishArgs } from "../../../../src/component/driver/copilotAgent/interfaces/PublishArgs";
import { AppStudioError } from "../../../../src/component/driver/teamsApp/errors";
import { Constants } from "../../../../src/component/driver/teamsApp/constants";
import { UserCancelError } from "../../../../src/error/common";
import { MockedLogProvider, MockedUserInteraction } from "../../../plugins/solution/util";
import { MockedM365Provider } from "../../../core/utils";
import { PackageService } from "../../../../src/component/m365/packageService";
import * as McpCertVerification from "../../../../src/component/driver/teamsApp/utils/McpCertVerification";

describe("copilotAgent/publish", async () => {
  const createManifestWithDA = (declarativeAgentFile: string) => {
    const manifest = new TeamsAppManifest();
    manifest.copilotAgents = {
      declarativeAgents: [{ id: "da1", file: declarativeAgentFile }],
    };
    return manifest;
  };
  const driver = new CopilotAgentPublishDriver();
  const mockedDriverContext: any = {
    m365TokenProvider: new MockedM365Provider(),
    logProvider: new MockedLogProvider(),
    ui: new MockedUserInteraction(),
    projectPath: "./",
  };

  afterEach(() => {
    sinon.restore();
  });

  it("should throw error if file not exists", async () => {
    const args: CopilotAgentPublishArgs = {
      appPackagePath: "fakepath",
    };
    const outputEnvVarNames = new Map<string, string>([
      ["titleId", "M365_PUBLISHED_TITLE_ID"],
      ["appId", "M365_PUBLISHED_APP_ID"],
    ]);

    const result = (await driver.execute(args, mockedDriverContext, outputEnvVarNames)).result;
    chai.assert(result.isErr());
    if (result.isErr()) {
      chai.assert.equal(AppStudioError.FileNotFoundError.name, result.error.name);
    }
  });

  it("invalid param error - empty appPackagePath", async () => {
    const args: CopilotAgentPublishArgs = {
      appPackagePath: "",
    };
    const outputEnvVarNames = new Map<string, string>([
      ["titleId", "M365_PUBLISHED_TITLE_ID"],
      ["appId", "M365_PUBLISHED_APP_ID"],
    ]);

    const result = (await driver.execute(args, mockedDriverContext, outputEnvVarNames)).result;
    chai.assert(result.isErr());
    if (result.isErr()) {
      chai.assert.equal("InvalidActionInputError", result.error.name);
    }
  });

  it("invalid param error - invalid scope", async () => {
    const args: CopilotAgentPublishArgs = {
      appPackagePath: "fakepath",
      scope: "invalid",
    };
    const outputEnvVarNames = new Map<string, string>([
      ["titleId", "M365_PUBLISHED_TITLE_ID"],
      ["appId", "M365_PUBLISHED_APP_ID"],
    ]);

    sinon.stub(fs, "pathExists").resolves(true);

    const result = (await driver.execute(args, mockedDriverContext, outputEnvVarNames)).result;
    chai.assert(result.isErr());
    if (result.isErr()) {
      chai.assert.equal("InvalidActionInputError", result.error.name);
    }
  });

  it("invalid param error - missing writeToEnvironmentFile", async () => {
    const args: CopilotAgentPublishArgs = {
      appPackagePath: "fakepath",
    };

    sinon.stub(fs, "pathExists").resolves(true);

    const result = (await driver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isErr());
    if (result.isErr()) {
      chai.assert.equal("InvalidActionInputError", result.error.name);
    }
  });

  it("happy path - default scope (Personal)", async () => {
    const args: CopilotAgentPublishArgs = {
      appPackagePath: "fakepath",
    };
    const outputEnvVarNames = new Map<string, string>([
      ["titleId", "M365_TITLE_ID"],
      ["appId", "M365_APP_ID"],
    ]);

    const titleId = uuid();
    const appId = uuid();

    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(fs, "readFile").callsFake(async () => {
      const zip = new AdmZip();
      zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify(new TeamsAppManifest())));
      zip.addFile("color.png", Buffer.from(""));
      zip.addFile("outline.png", Buffer.from(""));
      return zip.toBuffer();
    });
    sinon.stub(PackageService.prototype, "getTitleServiceUrl").resolves("https://fake.url");
    sinon.stub(PackageService.prototype, "publishAgent").resolves([titleId, appId, ""]);

    const result = await driver.execute(args, mockedDriverContext, outputEnvVarNames);
    chai.assert.isTrue(result.result.isOk());
    if (result.result.isOk()) {
      chai.assert.equal(result.result.value.get("M365_TITLE_ID"), titleId);
      chai.assert.equal(result.result.value.get("M365_APP_ID"), appId);
    }
  });

  it("happy path - tenant scope", async () => {
    const args: CopilotAgentPublishArgs = {
      appPackagePath: "fakepath",
      scope: "tenant",
    };
    const outputEnvVarNames = new Map<string, string>([
      ["titleId", "M365_PUBLISHED_TITLE_ID"],
      ["appId", "M365_PUBLISHED_APP_ID"],
    ]);

    const titleId = uuid();
    const appId = uuid();

    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(fs, "readFile").callsFake(async () => {
      const zip = new AdmZip();
      zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify(new TeamsAppManifest())));
      zip.addFile("color.png", Buffer.from(""));
      zip.addFile("outline.png", Buffer.from(""));
      return zip.toBuffer();
    });
    sinon.stub(PackageService.prototype, "getTitleServiceUrl").resolves("https://fake.url");
    sinon.stub(PackageService.prototype, "publishAgent").resolves([titleId, appId, ""]);

    const result = await driver.execute(args, mockedDriverContext, outputEnvVarNames);
    chai.assert.isTrue(result.result.isOk());
    if (result.result.isOk()) {
      chai.assert.equal(result.result.value.get("M365_PUBLISHED_TITLE_ID"), titleId);
      chai.assert.equal(result.result.value.get("M365_PUBLISHED_APP_ID"), appId);
    }
  });

  it("happy path - shared scope", async () => {
    const args: CopilotAgentPublishArgs = {
      appPackagePath: "fakepath",
      scope: "Shared",
    };
    const outputEnvVarNames = new Map<string, string>([
      ["titleId", "M365_TITLE_ID"],
      ["appId", "M365_APP_ID"],
    ]);

    const titleId = uuid();
    const appId = uuid();

    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(fs, "readFile").callsFake(async () => {
      const zip = new AdmZip();
      zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify(new TeamsAppManifest())));
      zip.addFile("color.png", Buffer.from(""));
      zip.addFile("outline.png", Buffer.from(""));
      return zip.toBuffer();
    });
    sinon.stub(PackageService.prototype, "getTitleServiceUrl").resolves("https://fake.url");
    sinon.stub(PackageService.prototype, "publishAgent").resolves([titleId, appId, ""]);

    const result = await driver.execute(args, mockedDriverContext, outputEnvVarNames);
    chai.assert.isTrue(result.result.isOk());
    if (result.result.isOk()) {
      chai.assert.equal(result.result.value.get("M365_TITLE_ID"), titleId);
      chai.assert.equal(result.result.value.get("M365_APP_ID"), appId);
    }
  });

  it("should return token error when getAccessToken fails", async () => {
    const args: CopilotAgentPublishArgs = {
      appPackagePath: "fakepath",
    };
    const outputEnvVarNames = new Map<string, string>([
      ["titleId", "M365_TITLE_ID"],
      ["appId", "M365_APP_ID"],
    ]);

    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(fs, "readFile").callsFake(async () => {
      const zip = new AdmZip();
      const manifest = new TeamsAppManifest();
      manifest.id = uuid();
      zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify(manifest)));
      return zip.toBuffer();
    });

    const tokenError = new UserCancelError();
    sinon.stub(mockedDriverContext.m365TokenProvider, "getAccessToken").resolves(err(tokenError));

    const result = (await driver.execute(args, mockedDriverContext, outputEnvVarNames)).result;
    chai.assert.isTrue(result.isErr());
    if (result.isErr()) {
      chai.assert.equal(result.error, tokenError);
    }
  });

  it("should return error when publishAgent throws", async () => {
    const args: CopilotAgentPublishArgs = {
      appPackagePath: "fakepath",
    };
    const outputEnvVarNames = new Map<string, string>([
      ["titleId", "M365_TITLE_ID"],
      ["appId", "M365_APP_ID"],
    ]);

    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(fs, "readFile").callsFake(async () => {
      const zip = new AdmZip();
      const manifest = new TeamsAppManifest();
      manifest.id = uuid();
      zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify(manifest)));
      return zip.toBuffer();
    });
    sinon.stub(PackageService.prototype, "getTitleServiceUrl").resolves("https://fake.url");
    sinon.stub(PackageService.prototype, "publishAgent").rejects(new Error("publish failed"));

    const result = (await driver.execute(args, mockedDriverContext, outputEnvVarNames)).result;
    chai.assert.isTrue(result.isErr());
    if (result.isErr()) {
      chai.assert.include(result.error.message, "publish failed");
    }
  });

  it("run method should work", async () => {
    const args: CopilotAgentPublishArgs = {
      appPackagePath: "fakepath",
    };

    const titleId = uuid();
    const appId = uuid();

    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(fs, "readFile").callsFake(async () => {
      const zip = new AdmZip();
      zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify(new TeamsAppManifest())));
      return zip.toBuffer();
    });
    sinon.stub(PackageService.prototype, "getTitleServiceUrl").resolves("https://fake.url");
    sinon.stub(PackageService.prototype, "publishAgent").resolves([titleId, appId, ""]);

    const result = await driver.run(args, mockedDriverContext);
    chai.assert.isTrue(result.isErr());
  });

  it("should throw error if manifest.json not found in package", async () => {
    const args: CopilotAgentPublishArgs = {
      appPackagePath: "fakepath",
    };
    const outputEnvVarNames = new Map<string, string>([
      ["titleId", "M365_TITLE_ID"],
      ["appId", "M365_APP_ID"],
    ]);

    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(fs, "readFile").callsFake(async () => {
      const zip = new AdmZip();
      zip.addFile("other.json", Buffer.from("{}"));
      return zip.toBuffer();
    });

    const result = (await driver.execute(args, mockedDriverContext, outputEnvVarNames)).result;
    chai.assert.isTrue(result.isErr());
    if (result.isErr()) {
      chai.assert.equal(AppStudioError.FileNotFoundError.name, result.error.name);
    }
  });

  it("should return shareLink when published with shared scope", async () => {
    const args: CopilotAgentPublishArgs = {
      appPackagePath: "fakepath",
      scope: "Shared",
    };
    const outputEnvVarNames = new Map<string, string>([
      ["titleId", "M365_TITLE_ID"],
      ["appId", "M365_APP_ID"],
      ["shareLink", "SHARE_LINK"],
    ]);

    const titleId = uuid();
    const appId = uuid();
    const shareLink = "https://fake.sharelink.com";

    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(fs, "readFile").callsFake(async () => {
      const zip = new AdmZip();
      zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify(new TeamsAppManifest())));
      return zip.toBuffer();
    });
    sinon.stub(PackageService.prototype, "getTitleServiceUrl").resolves("https://fake.url");
    sinon.stub(PackageService.prototype, "publishAgent").resolves([titleId, appId, shareLink]);

    const result = await driver.execute(args, mockedDriverContext, outputEnvVarNames);
    chai.assert.isTrue(result.result.isOk());
    if (result.result.isOk()) {
      chai.assert.equal(result.result.value.get("M365_TITLE_ID"), titleId);
      chai.assert.equal(result.result.value.get("M365_APP_ID"), appId);
      chai.assert.equal(result.result.value.get("SHARE_LINK"), shareLink);
    }
  });

  it("should not set shareLink when shareLinkKey is not provided", async () => {
    const args: CopilotAgentPublishArgs = {
      appPackagePath: "fakepath",
      scope: "Shared",
    };
    const outputEnvVarNames = new Map<string, string>([
      ["titleId", "M365_TITLE_ID"],
      ["appId", "M365_APP_ID"],
    ]);

    const titleId = uuid();
    const appId = uuid();
    const shareLink = "https://fake.sharelink.com";

    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(fs, "readFile").callsFake(async () => {
      const zip = new AdmZip();
      zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify(new TeamsAppManifest())));
      return zip.toBuffer();
    });
    sinon.stub(PackageService.prototype, "getTitleServiceUrl").resolves("https://fake.url");
    sinon.stub(PackageService.prototype, "publishAgent").resolves([titleId, appId, shareLink]);

    const result = await driver.execute(args, mockedDriverContext, outputEnvVarNames);
    chai.assert.isTrue(result.result.isOk());
    if (result.result.isOk()) {
      chai.assert.equal(result.result.value.get("M365_TITLE_ID"), titleId);
      chai.assert.equal(result.result.value.get("M365_APP_ID"), appId);
      chai.assert.isUndefined(result.result.value.get("SHARE_LINK"));
    }
  });

  it("should verify MCP certs for declarative agents with actions", async () => {
    const args: CopilotAgentPublishArgs = {
      appPackagePath: "fakepath",
    };
    const outputEnvVarNames = new Map<string, string>([
      ["titleId", "M365_TITLE_ID"],
      ["appId", "M365_APP_ID"],
    ]);

    const titleId = uuid();
    const appId = uuid();

    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(fs, "readFile").callsFake(async () => {
      const zip = new AdmZip();
      const manifest = createManifestWithDA("declarativeAgent.json");
      zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify(manifest)));
      const daManifest = {
        actions: [{ id: "action1", file: "plugin.json" }],
      };
      zip.addFile("declarativeAgent.json", Buffer.from(JSON.stringify(daManifest)));
      zip.addFile("plugin.json", Buffer.from(JSON.stringify({ runtimes: [] })));
      return zip.toBuffer();
    });
    sinon.stub(McpCertVerification, "verifyLocalMCPPluginCerts").resolves(true);
    sinon.stub(PackageService.prototype, "getTitleServiceUrl").resolves("https://fake.url");
    sinon.stub(PackageService.prototype, "publishAgent").resolves([titleId, appId, ""]);

    const result = await driver.execute(args, mockedDriverContext, outputEnvVarNames);
    chai.assert.isTrue(result.result.isOk());
  });

  it("should fail when MCP cert verification fails", async () => {
    const args: CopilotAgentPublishArgs = {
      appPackagePath: "fakepath",
    };
    const outputEnvVarNames = new Map<string, string>([
      ["titleId", "M365_TITLE_ID"],
      ["appId", "M365_APP_ID"],
    ]);

    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(fs, "readFile").callsFake(async () => {
      const zip = new AdmZip();
      const manifest = createManifestWithDA("declarativeAgent.json");
      zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify(manifest)));
      const daManifest = {
        actions: [{ id: "action1", file: "plugin.json" }],
      };
      zip.addFile("declarativeAgent.json", Buffer.from(JSON.stringify(daManifest)));
      zip.addFile("plugin.json", Buffer.from(JSON.stringify({ runtimes: [] })));
      return zip.toBuffer();
    });
    sinon.stub(McpCertVerification, "verifyLocalMCPPluginCerts").resolves(false);

    const result = (await driver.execute(args, mockedDriverContext, outputEnvVarNames)).result;
    chai.assert.isTrue(result.isErr());
    if (result.isErr()) {
      chai.assert.equal(AppStudioError.ValidationFailedError.name, result.error.name);
    }
  });

  it("should skip MCP verification when declarative agent file not found", async () => {
    const args: CopilotAgentPublishArgs = {
      appPackagePath: "fakepath",
    };
    const outputEnvVarNames = new Map<string, string>([
      ["titleId", "M365_TITLE_ID"],
      ["appId", "M365_APP_ID"],
    ]);

    const titleId = uuid();
    const appId = uuid();

    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(fs, "readFile").callsFake(async () => {
      const zip = new AdmZip();
      const manifest = createManifestWithDA("nonexistent.json");
      zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify(manifest)));
      return zip.toBuffer();
    });
    sinon.stub(PackageService.prototype, "getTitleServiceUrl").resolves("https://fake.url");
    sinon.stub(PackageService.prototype, "publishAgent").resolves([titleId, appId, ""]);

    const result = await driver.execute(args, mockedDriverContext, outputEnvVarNames);
    chai.assert.isTrue(result.result.isOk());
  });

  it("should skip action file verification when action file not found", async () => {
    const args: CopilotAgentPublishArgs = {
      appPackagePath: "fakepath",
    };
    const outputEnvVarNames = new Map<string, string>([
      ["titleId", "M365_TITLE_ID"],
      ["appId", "M365_APP_ID"],
    ]);

    const titleId = uuid();
    const appId = uuid();

    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(fs, "readFile").callsFake(async () => {
      const zip = new AdmZip();
      const manifest = createManifestWithDA("declarativeAgent.json");
      zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify(manifest)));
      const daManifest = {
        actions: [{ id: "action1", file: "nonexistent-plugin.json" }],
      };
      zip.addFile("declarativeAgent.json", Buffer.from(JSON.stringify(daManifest)));
      return zip.toBuffer();
    });
    sinon.stub(PackageService.prototype, "getTitleServiceUrl").resolves("https://fake.url");
    sinon.stub(PackageService.prototype, "publishAgent").resolves([titleId, appId, ""]);

    const result = await driver.execute(args, mockedDriverContext, outputEnvVarNames);
    chai.assert.isTrue(result.result.isOk());
  });

  it("should handle copilotExtensions.declarativeCopilots format", async () => {
    const args: CopilotAgentPublishArgs = {
      appPackagePath: "fakepath",
    };
    const outputEnvVarNames = new Map<string, string>([
      ["titleId", "M365_TITLE_ID"],
      ["appId", "M365_APP_ID"],
    ]);

    const titleId = uuid();
    const appId = uuid();

    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(fs, "readFile").callsFake(async () => {
      const zip = new AdmZip();
      const manifest = new TeamsAppManifest();
      manifest.copilotExtensions = {
        declarativeCopilots: [{ id: "dc1", file: "declarativeAgent.json" }],
      };
      zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify(manifest)));
      const daManifest = {
        actions: [{ id: "action1", file: "plugin.json" }],
      };
      zip.addFile("declarativeAgent.json", Buffer.from(JSON.stringify(daManifest)));
      zip.addFile("plugin.json", Buffer.from(JSON.stringify({ runtimes: [] })));
      return zip.toBuffer();
    });
    sinon.stub(McpCertVerification, "verifyLocalMCPPluginCerts").resolves(true);
    sinon.stub(PackageService.prototype, "getTitleServiceUrl").resolves("https://fake.url");
    sinon.stub(PackageService.prototype, "publishAgent").resolves([titleId, appId, ""]);

    const result = await driver.execute(args, mockedDriverContext, outputEnvVarNames);
    chai.assert.isTrue(result.result.isOk());
  });
});
