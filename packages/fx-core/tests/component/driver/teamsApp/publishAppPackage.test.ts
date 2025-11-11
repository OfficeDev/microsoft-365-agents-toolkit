// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { ok, Platform, TeamsAppManifest } from "@microsoft/teamsfx-api";
import AdmZip from "adm-zip";
import chai from "chai";
import fs from "fs-extra";
import "mocha";
import * as sinon from "sinon";
import { v4 as uuid } from "uuid";
import { teamsDevPortalClient } from "../../../../src/client/teamsDevPortalClient";
import { AppStudioError } from "../../../../src/component/driver/teamsApp/errors";
import { PublishingState } from "../../../../src/component/driver/teamsApp/interfaces/appdefinitions/IPublishingAppDefinition";
import { PublishAppPackageArgs } from "../../../../src/component/driver/teamsApp/interfaces/PublishAppPackageArgs";
import { PublishAppPackageDriver } from "../../../../src/component/driver/teamsApp/publishAppPackage";
import { UserCancelError } from "../../../../src/error/common";
import { MockedLogProvider, MockedUserInteraction } from "../../../plugins/solution/util";
import { Constants } from "./../../../../src/component/driver/teamsApp/constants";
import { MockedM365Provider } from "../../../core/utils";
import { ODRProvider } from "../../../../src/component/utils/odrProvider";

describe("teamsApp/publishAppPackage", async () => {
  const teamsAppDriver = new PublishAppPackageDriver();
  const mockedDriverContext: any = {
    m365TokenProvider: new MockedM365Provider(),
    logProvider: new MockedLogProvider(),
    ui: new MockedUserInteraction(),
    projectPath: "./",
  };

  const state = {
    lastModifiedDateTime: new Date(),
    teamsAppId: "",
    displayName: "fakeName",
    publishingState: PublishingState.submitted,
  };

  afterEach(() => {
    sinon.restore();
  });

  it("should throw error if file not exists", async () => {
    const args: PublishAppPackageArgs = {
      appPackagePath: "fakepath",
    };

    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isErr());
    if (result.isErr()) {
      chai.assert.equal(AppStudioError.FileNotFoundError.name, result.error.name);
    }
  });

  it("invalid param error", async () => {
    const args: PublishAppPackageArgs = {
      appPackagePath: "",
    };

    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isErr());
    if (result.isErr()) {
      chai.assert.equal("InvalidActionInputError", result.error.name);
    }
  });

  it("happy path", async () => {
    const args: PublishAppPackageArgs = {
      appPackagePath: "fakepath",
    };

    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(fs, "readFile").callsFake(async () => {
      const zip = new AdmZip();
      zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify(new TeamsAppManifest())));
      zip.addFile("color.png", new Buffer(""));
      zip.addFile("outlie.png", new Buffer(""));

      const archivedFile = zip.toBuffer();
      return archivedFile;
    });
    sinon.stub(teamsDevPortalClient, "getStaggedApp").resolves(undefined);
    sinon.stub(teamsDevPortalClient, "publishTeamsApp").resolves(uuid());

    const result = await teamsAppDriver.execute(args, mockedDriverContext);
    console.log(JSON.stringify(result));
    chai.assert.isTrue(result.result.isOk());
  });

  it("happy path - user cancel", async () => {
    const args: PublishAppPackageArgs = {
      appPackagePath: "fakepath",
    };

    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(fs, "readFile").callsFake(async () => {
      const zip = new AdmZip();
      zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify(new TeamsAppManifest())));
      zip.addFile("color.png", new Buffer(""));
      zip.addFile("outlie.png", new Buffer(""));

      const archivedFile = zip.toBuffer();
      return archivedFile;
    });
    sinon.stub(teamsDevPortalClient, "getStaggedApp").resolves(state);
    sinon.stub(mockedDriverContext.ui, "showMessage").resolves(ok("Cancel"));

    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert.isTrue(result.isErr());
    if (result.isErr()) {
      chai.assert.isTrue(result.error instanceof UserCancelError);
    }
  });

  it("happy path - update published app", async () => {
    const args: PublishAppPackageArgs = {
      appPackagePath: "fakepath",
    };

    mockedDriverContext.platform = Platform.CLI;

    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(fs, "readFile").callsFake(async () => {
      const zip = new AdmZip();
      zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify(new TeamsAppManifest())));
      zip.addFile("color.png", new Buffer(""));
      zip.addFile("outlie.png", new Buffer(""));

      const archivedFile = zip.toBuffer();
      return archivedFile;
    });
    sinon.stub(teamsDevPortalClient, "getStaggedApp").resolves(state);
    sinon.stub(teamsDevPortalClient, "publishTeamsAppUpdate").resolves(uuid());
    sinon.stub(mockedDriverContext.ui, "showMessage").resolves(ok("Confirm"));

    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert.isTrue(result.isOk());
  });

  describe("MCP plugin certificate verification", () => {
    it("should pass when no declarative agents exist", async () => {
      const args: PublishAppPackageArgs = {
        appPackagePath: "fakepath",
      };

      sinon.stub(fs, "pathExists").resolves(true);
      sinon.stub(fs, "readFile").callsFake(async () => {
        const zip = new AdmZip();
        const manifest = new TeamsAppManifest();
        zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify(manifest)));
        return zip.toBuffer();
      });
      sinon.stub(teamsDevPortalClient, "getStaggedApp").resolves(undefined);
      sinon.stub(teamsDevPortalClient, "publishTeamsApp").resolves(uuid());

      const result = await teamsAppDriver.execute(args, mockedDriverContext);
      chai.assert.isTrue(result.result.isOk());
    });

    it("should pass when declarativeAgent has no actions", async () => {
      const args: PublishAppPackageArgs = {
        appPackagePath: "fakepath",
      };

      sinon.stub(fs, "pathExists").resolves(true);
      sinon.stub(fs, "readFile").callsFake(async () => {
        const zip = new AdmZip();
        const manifest: any = new TeamsAppManifest();
        manifest.copilotAgents = {
          declarativeAgents: [{ id: "agent1", file: "declarativeAgent.json" }],
        };
        zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify(manifest)));
        zip.addFile(
          "declarativeAgent.json",
          Buffer.from(
            JSON.stringify({
              name: "Test Agent",
              description: "Test",
            })
          )
        );
        return zip.toBuffer();
      });
      sinon.stub(teamsDevPortalClient, "getStaggedApp").resolves(undefined);
      sinon.stub(teamsDevPortalClient, "publishTeamsApp").resolves(uuid());

      const result = await teamsAppDriver.execute(args, mockedDriverContext);
      chai.assert.isTrue(result.result.isOk());
    });

    it("should pass when action has no runtimes", async () => {
      const args: PublishAppPackageArgs = {
        appPackagePath: "fakepath",
      };

      sinon.stub(fs, "pathExists").resolves(true);
      sinon.stub(fs, "readFile").callsFake(async () => {
        const zip = new AdmZip();
        const manifest: any = new TeamsAppManifest();
        manifest.copilotAgents = {
          declarativeAgents: [{ id: "agent1", file: "declarativeAgent.json" }],
        };
        zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify(manifest)));
        zip.addFile(
          "declarativeAgent.json",
          Buffer.from(
            JSON.stringify({
              name: "Test Agent",
              description: "Test",
              actions: [{ id: "action1", file: "ai-plugin.json" }],
            })
          )
        );
        zip.addFile(
          "ai-plugin.json",
          Buffer.from(
            JSON.stringify({
              schema_version: "v2.1",
              name_for_human: "Test Plugin",
            })
          )
        );
        return zip.toBuffer();
      });
      sinon.stub(teamsDevPortalClient, "getStaggedApp").resolves(undefined);
      sinon.stub(teamsDevPortalClient, "publishTeamsApp").resolves(uuid());

      const result = await teamsAppDriver.execute(args, mockedDriverContext);
      chai.assert.isTrue(result.result.isOk());
    });

    it("should pass when action has only non-LocalPlugin runtimes", async () => {
      const args: PublishAppPackageArgs = {
        appPackagePath: "fakepath",
      };

      sinon.stub(fs, "pathExists").resolves(true);
      sinon.stub(fs, "readFile").callsFake(async () => {
        const zip = new AdmZip();
        const manifest: any = new TeamsAppManifest();
        manifest.copilotAgents = {
          declarativeAgents: [{ id: "agent1", file: "declarativeAgent.json" }],
        };
        zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify(manifest)));
        zip.addFile(
          "declarativeAgent.json",
          Buffer.from(
            JSON.stringify({
              name: "Test Agent",
              description: "Test",
              actions: [{ id: "action1", file: "ai-plugin.json" }],
            })
          )
        );
        zip.addFile(
          "ai-plugin.json",
          Buffer.from(
            JSON.stringify({
              schema_version: "v2.1",
              name_for_human: "Test Plugin",
              runtimes: [
                {
                  type: "OpenApi",
                  spec: { url: "https://api.example.com/openapi.json" },
                },
              ],
            })
          )
        );
        return zip.toBuffer();
      });
      sinon.stub(teamsDevPortalClient, "getStaggedApp").resolves(undefined);
      sinon.stub(teamsDevPortalClient, "publishTeamsApp").resolves(uuid());

      const result = await teamsAppDriver.execute(args, mockedDriverContext);
      chai.assert.isTrue(result.result.isOk());
    });

    it("should pass when LocalPlugin runtime is not MCP (no mcp:// prefix)", async () => {
      const args: PublishAppPackageArgs = {
        appPackagePath: "fakepath",
      };

      sinon.stub(fs, "pathExists").resolves(true);
      sinon.stub(fs, "readFile").callsFake(async () => {
        const zip = new AdmZip();
        const manifest: any = new TeamsAppManifest();
        manifest.copilotAgents = {
          declarativeAgents: [{ id: "agent1", file: "declarativeAgent.json" }],
        };
        zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify(manifest)));
        zip.addFile(
          "declarativeAgent.json",
          Buffer.from(
            JSON.stringify({
              name: "Test Agent",
              description: "Test",
              actions: [{ id: "action1", file: "ai-plugin.json" }],
            })
          )
        );
        zip.addFile(
          "ai-plugin.json",
          Buffer.from(
            JSON.stringify({
              schema_version: "v2.1",
              name_for_human: "Test Plugin",
              runtimes: [
                {
                  type: "LocalPlugin",
                  spec: { local_endpoint: "http://localhost:3000" },
                },
              ],
            })
          )
        );
        return zip.toBuffer();
      });
      sinon.stub(teamsDevPortalClient, "getStaggedApp").resolves(undefined);
      sinon.stub(teamsDevPortalClient, "publishTeamsApp").resolves(uuid());

      const result = await teamsAppDriver.execute(args, mockedDriverContext);
      chai.assert.isTrue(result.result.isOk());
    });

    it("should pass when MCP server not found in ODR list (non-MCP local plugin)", async () => {
      const args: PublishAppPackageArgs = {
        appPackagePath: "fakepath",
      };

      sinon.stub(fs, "pathExists").resolves(true);
      sinon.stub(fs, "readFile").callsFake(async () => {
        const zip = new AdmZip();
        const manifest: any = new TeamsAppManifest();
        manifest.copilotAgents = {
          declarativeAgents: [{ id: "agent1", file: "declarativeAgent.json" }],
        };
        zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify(manifest)));
        zip.addFile(
          "declarativeAgent.json",
          Buffer.from(
            JSON.stringify({
              name: "Test Agent",
              description: "Test",
              actions: [{ id: "action1", file: "ai-plugin.json" }],
            })
          )
        );
        zip.addFile(
          "ai-plugin.json",
          Buffer.from(
            JSON.stringify({
              schema_version: "v2.1",
              name_for_human: "Test Plugin",
              runtimes: [
                {
                  type: "LocalPlugin",
                  spec: { local_endpoint: "mcp://unknown-server" },
                },
              ],
            })
          )
        );
        return zip.toBuffer();
      });
      sinon.stub(ODRProvider, "listServers").resolves([]);
      sinon.stub(teamsDevPortalClient, "getStaggedApp").resolves(undefined);
      sinon.stub(teamsDevPortalClient, "publishTeamsApp").resolves(uuid());

      const result = await teamsAppDriver.execute(args, mockedDriverContext);
      chai.assert.isTrue(result.result.isOk());
    });

    it("should pass when MCP plugin has valid certificate", async () => {
      const args: PublishAppPackageArgs = {
        appPackagePath: "fakepath",
      };

      sinon.stub(fs, "pathExists").resolves(true);
      sinon.stub(fs, "readFile").callsFake(async () => {
        const zip = new AdmZip();
        const manifest: any = new TeamsAppManifest();
        manifest.copilotAgents = {
          declarativeAgents: [{ id: "agent1", file: "declarativeAgent.json" }],
        };
        zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify(manifest)));
        zip.addFile(
          "declarativeAgent.json",
          Buffer.from(
            JSON.stringify({
              name: "Test Agent",
              description: "Test",
              actions: [{ id: "action1", file: "ai-plugin.json" }],
            })
          )
        );
        zip.addFile(
          "ai-plugin.json",
          Buffer.from(
            JSON.stringify({
              schema_version: "v2.1",
              name_for_human: "Test Plugin",
              runtimes: [
                {
                  type: "LocalPlugin",
                  spec: { local_endpoint: "mcp://test-server" },
                },
              ],
            })
          )
        );
        return zip.toBuffer();
      });
      sinon.stub(ODRProvider, "listServers").resolves([
        {
          name: "test-server",
          display_name: "Test Server",
          description: "Test",
          version: "1.0.0",
          identifier: "test-server",
          packageFamily: "TestPackage_12345",
          command: "test.exe",
          args: [],
          tools: [],
        },
      ]);
      sinon
        .stub(PublishAppPackageDriver.prototype as any, "verifyPackageFamilyCertIsValid")
        .resolves(true);
      sinon.stub(teamsDevPortalClient, "getStaggedApp").resolves(undefined);
      sinon.stub(teamsDevPortalClient, "publishTeamsApp").resolves(uuid());

      const result = await teamsAppDriver.execute(args, mockedDriverContext);
      chai.assert.isTrue(result.result.isOk());
    });

    it("should fail when MCP plugin has self-signed certificate", async () => {
      const args: PublishAppPackageArgs = {
        appPackagePath: "fakepath",
      };

      sinon.stub(fs, "pathExists").resolves(true);
      sinon.stub(fs, "readFile").callsFake(async () => {
        const zip = new AdmZip();
        const manifest: any = new TeamsAppManifest();
        manifest.copilotAgents = {
          declarativeAgents: [{ id: "agent1", file: "declarativeAgent.json" }],
        };
        zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify(manifest)));
        zip.addFile(
          "declarativeAgent.json",
          Buffer.from(
            JSON.stringify({
              name: "Test Agent",
              description: "Test",
              actions: [{ id: "action1", file: "ai-plugin.json" }],
            })
          )
        );
        zip.addFile(
          "ai-plugin.json",
          Buffer.from(
            JSON.stringify({
              schema_version: "v2.1",
              name_for_human: "Test Plugin",
              runtimes: [
                {
                  type: "LocalPlugin",
                  spec: { local_endpoint: "mcp://test-server" },
                },
              ],
            })
          )
        );
        return zip.toBuffer();
      });
      sinon.stub(ODRProvider, "listServers").resolves([
        {
          name: "test-server",
          display_name: "Test Server",
          description: "Test",
          version: "1.0.0",
          identifier: "test-server",
          packageFamily: "TestPackage_12345",
          command: "test.exe",
          args: [],
          tools: [],
        },
      ]);
      sinon
        .stub(PublishAppPackageDriver.prototype as any, "verifyPackageFamilyCertIsValid")
        .resolves(false);

      const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
      chai.assert.isTrue(result.isErr());
      if (result.isErr()) {
        chai.assert.include(result.error.message, "certificate verification failed");
      }
    });

    it("should fail when MCP plugin has no certificate", async () => {
      const args: PublishAppPackageArgs = {
        appPackagePath: "fakepath",
      };

      sinon.stub(fs, "pathExists").resolves(true);
      sinon.stub(fs, "readFile").callsFake(async () => {
        const zip = new AdmZip();
        const manifest: any = new TeamsAppManifest();
        manifest.copilotAgents = {
          declarativeAgents: [{ id: "agent1", file: "declarativeAgent.json" }],
        };
        zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify(manifest)));
        zip.addFile(
          "declarativeAgent.json",
          Buffer.from(
            JSON.stringify({
              name: "Test Agent",
              description: "Test",
              actions: [{ id: "action1", file: "ai-plugin.json" }],
            })
          )
        );
        zip.addFile(
          "ai-plugin.json",
          Buffer.from(
            JSON.stringify({
              schema_version: "v2.1",
              name_for_human: "Test Plugin",
              runtimes: [
                {
                  type: "LocalPlugin",
                  spec: { local_endpoint: "mcp://test-server" },
                },
              ],
            })
          )
        );
        return zip.toBuffer();
      });
      sinon.stub(ODRProvider, "listServers").resolves([
        {
          name: "test-server",
          display_name: "Test Server",
          description: "Test",
          version: "1.0.0",
          identifier: "test-server",
          packageFamily: "TestPackage_12345",
          command: "test.exe",
          args: [],
          tools: [],
        },
      ]);
      sinon
        .stub(PublishAppPackageDriver.prototype as any, "verifyPackageFamilyCertIsValid")
        .resolves(false);

      const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
      chai.assert.isTrue(result.isErr());
      if (result.isErr()) {
        chai.assert.include(result.error.message, "certificate verification failed");
      }
    });

    it("should verify multiple actions with mixed runtimes", async () => {
      const args: PublishAppPackageArgs = {
        appPackagePath: "fakepath",
      };

      sinon.stub(fs, "pathExists").resolves(true);
      sinon.stub(fs, "readFile").callsFake(async () => {
        const zip = new AdmZip();
        const manifest: any = new TeamsAppManifest();
        manifest.copilotExtensions = {
          declarativeCopilots: [{ id: "agent1", file: "declarativeAgent.json" }],
        };
        zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify(manifest)));
        zip.addFile(
          "declarativeAgent.json",
          Buffer.from(
            JSON.stringify({
              name: "Test Agent",
              description: "Test",
              actions: [
                { id: "action1", file: "plugin1.json" },
                { id: "action2", file: "plugin2.json" },
              ],
            })
          )
        );
        zip.addFile(
          "plugin1.json",
          Buffer.from(
            JSON.stringify({
              schema_version: "v2.1",
              name_for_human: "Plugin 1",
              runtimes: [
                {
                  type: "OpenApi",
                  spec: { url: "https://api.example.com/openapi.json" },
                },
              ],
            })
          )
        );
        zip.addFile(
          "plugin2.json",
          Buffer.from(
            JSON.stringify({
              schema_version: "v2.1",
              name_for_human: "Plugin 2",
              runtimes: [
                {
                  type: "LocalPlugin",
                  spec: { local_endpoint: "mcp://test-server" },
                },
              ],
            })
          )
        );
        return zip.toBuffer();
      });
      sinon.stub(ODRProvider, "listServers").resolves([
        {
          name: "test-server",
          display_name: "Test Server",
          description: "Test",
          version: "1.0.0",
          identifier: "test-server",
          packageFamily: "TestPackage_12345",
          command: "test.exe",
          args: [],
          tools: [],
        },
      ]);
      sinon
        .stub(PublishAppPackageDriver.prototype as any, "verifyPackageFamilyCertIsValid")
        .resolves(true);
      sinon.stub(teamsDevPortalClient, "getStaggedApp").resolves(undefined);
      sinon.stub(teamsDevPortalClient, "publishTeamsApp").resolves(uuid());

      const result = await teamsAppDriver.execute(args, mockedDriverContext);
      chai.assert.isTrue(result.result.isOk());
    });
  });
});
