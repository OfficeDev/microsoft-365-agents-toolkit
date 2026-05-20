// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { err, ok, Platform, TeamsAppManifest } from "@microsoft/teamsfx-api";
import AdmZip from "adm-zip";
import chai from "chai";
import fs from "fs-extra";
import "mocha";
import mockedEnv from "mocked-env";
import * as sinon from "sinon";
import { v4 as uuid } from "uuid";
import { GraphClient } from "../../../../src/client/graphClient";
import { teamsDevPortalClient } from "../../../../src/client/teamsDevPortalClient";
import { SovereignCloudEnvironment } from "../../../../src/common/accountUtils";
import { FeatureFlagName } from "../../../../src/common/featureFlags";
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
  let restoreEnv: (() => void) | undefined;
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
    restoreEnv?.();
    restoreEnv = undefined;
  });

  it("skip publish in GCCH", async () => {
    restoreEnv = mockedEnv({
      [FeatureFlagName.SovereignCloudEnvironment]: SovereignCloudEnvironment.GCCH,
    });
    const publishTeamsAppSpy = sinon.spy(teamsDevPortalClient, "publishTeamsApp");
    const pathExistsStub = sinon.stub(fs, "pathExists");

    const args: PublishAppPackageArgs = {
      appPackagePath: "fakepath",
    };

    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isOk());
    sinon.assert.notCalled(publishTeamsAppSpy);
    sinon.assert.notCalled(pathExistsStub);
  });

  it("skip publish in DoD", async () => {
    restoreEnv = mockedEnv({
      [FeatureFlagName.SovereignCloudEnvironment]: SovereignCloudEnvironment.DOD,
    });
    const publishTeamsAppSpy = sinon.spy(teamsDevPortalClient, "publishTeamsApp");
    const pathExistsStub = sinon.stub(fs, "pathExists");

    const args: PublishAppPackageArgs = {
      appPackagePath: "fakepath",
    };

    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isOk());
    sinon.assert.notCalled(publishTeamsAppSpy);
    sinon.assert.notCalled(pathExistsStub);
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
    sinon.stub(GraphClient.prototype, "getStagedApp").resolves(undefined);
    sinon.stub(GraphClient.prototype, "publishTeamsApp").resolves(uuid());

    const result = await teamsAppDriver.execute(args, mockedDriverContext);
    chai.assert.isTrue(result.result.isOk());
  });

  it("should return token error when getAccessToken fails", async () => {
    const args: PublishAppPackageArgs = {
      appPackagePath: "fakepath",
    };

    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(fs, "readFile").callsFake(async () => {
      const zip = new AdmZip();
      const manifest = new TeamsAppManifest();
      manifest.id = uuid();
      manifest.name = { short: "test-app", full: "test-app" } as any;
      zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify(manifest)));
      return zip.toBuffer();
    });

    const tokenError = new UserCancelError();
    sinon.stub(mockedDriverContext.m365TokenProvider, "getAccessToken").resolves(err(tokenError));

    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert.isTrue(result.isErr());
    if (result.isErr()) {
      chai.assert.equal(result.error, tokenError);
    }
  });

  it("should return error when publishTeamsApp throws", async () => {
    const args: PublishAppPackageArgs = {
      appPackagePath: "fakepath",
    };

    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(fs, "readFile").callsFake(async () => {
      const zip = new AdmZip();
      const manifest = new TeamsAppManifest();
      manifest.id = uuid();
      manifest.name = { short: "test-app", full: "test-app" } as any;
      zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify(manifest)));
      return zip.toBuffer();
    });

    sinon.stub(GraphClient.prototype, "getStagedApp").resolves(undefined);
    sinon.stub(GraphClient.prototype, "publishTeamsApp").rejects(new Error("publish failed"));

    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert.isTrue(result.isErr());
    if (result.isErr()) {
      chai.assert.include(result.error.message, "publish failed");
    }
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
    sinon.stub(GraphClient.prototype, "getStagedApp").resolves(state);
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
    sinon.stub(GraphClient.prototype, "getStagedApp").resolves(state);
    sinon.stub(GraphClient.prototype, "publishTeamsAppUpdate").resolves(uuid());
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
      sinon.stub(GraphClient.prototype, "getStagedApp").resolves(undefined);
      sinon.stub(GraphClient.prototype, "publishTeamsApp").resolves(uuid());

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
      sinon.stub(GraphClient.prototype, "getStagedApp").resolves(undefined);
      sinon.stub(GraphClient.prototype, "publishTeamsApp").resolves(uuid());

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
      sinon.stub(GraphClient.prototype, "getStagedApp").resolves(undefined);
      sinon.stub(GraphClient.prototype, "publishTeamsApp").resolves(uuid());

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
      sinon.stub(GraphClient.prototype, "getStagedApp").resolves(undefined);
      sinon.stub(GraphClient.prototype, "publishTeamsApp").resolves(uuid());

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
      sinon.stub(GraphClient.prototype, "getStagedApp").resolves(undefined);
      sinon.stub(GraphClient.prototype, "publishTeamsApp").resolves(uuid());

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
      sinon.stub(GraphClient.prototype, "getStagedApp").resolves(undefined);
      sinon.stub(GraphClient.prototype, "publishTeamsApp").resolves(uuid());

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
      sinon.stub(GraphClient.prototype, "getStagedApp").resolves(undefined);
      sinon.stub(GraphClient.prototype, "publishTeamsApp").resolves(uuid());

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
      sinon.stub(GraphClient.prototype, "getStagedApp").resolves(undefined);
      sinon.stub(GraphClient.prototype, "publishTeamsApp").resolves(uuid());

      const result = await teamsAppDriver.execute(args, mockedDriverContext);
      chai.assert.isTrue(result.result.isOk());
    });
  });

  // eslint-disable-next-line no-secrets/no-secrets
  describe("verifyPackageFamilyCertIsValid", () => {
    const teamsAppDriver = new PublishAppPackageDriver();
    let execStub: sinon.SinonStub;

    beforeEach(() => {
      const childProcess = require("child_process");
      execStub = sinon.stub(childProcess, "exec");
    });

    afterEach(() => {
      sinon.restore();
    });

    it("should return true when package has valid certificate (Store signature)", async () => {
      execStub.callsFake(
        (
          command: string,
          callback: (error: Error | null, result: { stdout: string; stderr: string }) => void
        ) => {
          callback(null, {
            stdout: "$_.SignatureKind\n-----------------\nStore\n",
            stderr: "",
          });
        }
      );

      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const result = await (teamsAppDriver as any).verifyPackageFamilyCertIsValid(
        "TestPackage_12345"
      );
      chai.assert.isTrue(result);
    });

    it("should return true when package has System signature", async () => {
      execStub.callsFake(
        (
          command: string,
          callback: (error: Error | null, result: { stdout: string; stderr: string }) => void
        ) => {
          callback(null, {
            stdout: "$_.SignatureKind\n-----------------\nSystem\n",
            stderr: "",
          });
        }
      );

      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const result = await (teamsAppDriver as any).verifyPackageFamilyCertIsValid(
        "TestPackage_12345"
      );
      chai.assert.isTrue(result);
    });

    it("should return false when package has Developer signature (self-signed)", async () => {
      execStub.callsFake(
        (
          command: string,
          callback: (error: Error | null, result: { stdout: string; stderr: string }) => void
        ) => {
          callback(null, {
            stdout: "$_.SignatureKind\n-----------------\nDeveloper\n",
            stderr: "",
          });
        }
      );

      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const result = await (teamsAppDriver as any).verifyPackageFamilyCertIsValid(
        "TestPackage_12345"
      );
      chai.assert.isFalse(result);
    });

    it("should return false when package has developer in mixed case", async () => {
      execStub.callsFake(
        (
          command: string,
          callback: (error: Error | null, result: { stdout: string; stderr: string }) => void
        ) => {
          callback(null, {
            stdout: "$_.SignatureKind\n-----------------\nDEVELOPER\n",
            stderr: "",
          });
        }
      );

      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const result = await (teamsAppDriver as any).verifyPackageFamilyCertIsValid(
        "TestPackage_12345"
      );
      chai.assert.isFalse(result);
    });

    it("should return false when stdout is empty", async () => {
      execStub.callsFake(
        (
          command: string,
          callback: (error: Error | null, result: { stdout: string; stderr: string }) => void
        ) => {
          callback(null, {
            stdout: "",
            stderr: "",
          });
        }
      );

      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const result = await (teamsAppDriver as any).verifyPackageFamilyCertIsValid(
        "TestPackage_12345"
      );
      chai.assert.isFalse(result);
    });

    it("should return false when stdout is null", async () => {
      execStub.callsFake(
        (
          command: string,
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          callback: (error: Error | null, result: { stdout: any; stderr: string }) => void
        ) => {
          callback(null, {
            stdout: null,
            stderr: "",
          });
        }
      );

      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const result = await (teamsAppDriver as any).verifyPackageFamilyCertIsValid(
        "TestPackage_12345"
      );
      chai.assert.isFalse(result);
    });

    it("should return false when stdout is undefined", async () => {
      execStub.callsFake(
        (
          command: string,
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          callback: (error: Error | null, result: { stdout: any; stderr: string }) => void
        ) => {
          callback(null, {
            stdout: undefined,
            stderr: "",
          });
        }
      );

      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const result = await (teamsAppDriver as any).verifyPackageFamilyCertIsValid(
        "TestPackage_12345"
      );
      chai.assert.isFalse(result);
    });

    it("should return false when PowerShell command fails", async () => {
      execStub.callsFake(
        (
          command: string,
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          callback: (error: Error | null, result: any) => void
        ) => {
          callback(new Error("PowerShell execution failed"), null);
        }
      );

      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const result = await (teamsAppDriver as any).verifyPackageFamilyCertIsValid(
        "TestPackage_12345"
      );
      chai.assert.isFalse(result);
    });

    it("should return false when package not found", async () => {
      execStub.callsFake(
        (
          command: string,
          callback: (error: Error | null, result: { stdout: string; stderr: string }) => void
        ) => {
          callback(null, {
            stdout: "",
            stderr: "Get-AppxPackage : No packages were found",
          });
        }
      );

      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const result = await (teamsAppDriver as any).verifyPackageFamilyCertIsValid(
        "NonExistentPackage_99999"
      );
      chai.assert.isFalse(result);
    });

    it("should return true when package has Enterprise signature", async () => {
      execStub.callsFake(
        (
          command: string,
          callback: (error: Error | null, result: { stdout: string; stderr: string }) => void
        ) => {
          callback(null, {
            stdout: "$_.SignatureKind\n-----------------\nEnterprise\n",
            stderr: "",
          });
        }
      );

      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const result = await (teamsAppDriver as any).verifyPackageFamilyCertIsValid(
        "TestPackage_12345"
      );
      chai.assert.isTrue(result);
    });

    it("should handle stdout with extra whitespace", async () => {
      execStub.callsFake(
        (
          command: string,
          callback: (error: Error | null, result: { stdout: string; stderr: string }) => void
        ) => {
          callback(null, {
            stdout: "  \n\n  Store  \n\n  ",
            stderr: "",
          });
        }
      );

      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const result = await (teamsAppDriver as any).verifyPackageFamilyCertIsValid(
        "TestPackage_12345"
      );
      chai.assert.isTrue(result);
    });

    it("should correctly use the package name in PowerShell command", async () => {
      const packageName = "MyApp_abc123xyz";
      execStub.callsFake(
        (
          command: string,
          callback: (error: Error | null, result: { stdout: string; stderr: string }) => void
        ) => {
          chai.assert.include(command, packageName);
          callback(null, {
            stdout: "Store",
            stderr: "",
          });
        }
      );

      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      await (teamsAppDriver as any).verifyPackageFamilyCertIsValid(packageName);
      chai.assert.isTrue(execStub.calledOnce);
    });

    it("should handle timeout errors gracefully", async () => {
      execStub.callsFake(
        (
          command: string,
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          callback: (error: NodeJS.ErrnoException | null, result: any) => void
        ) => {
          const timeoutError = new Error("Command timed out") as NodeJS.ErrnoException;
          timeoutError.code = "ETIMEDOUT";
          callback(timeoutError, null);
        }
      );

      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const result = await (teamsAppDriver as any).verifyPackageFamilyCertIsValid(
        "TestPackage_12345"
      );
      chai.assert.isFalse(result);
    });

    it("should return false when stdout contains developer anywhere in text", async () => {
      execStub.callsFake(
        (
          command: string,
          callback: (error: Error | null, result: { stdout: string; stderr: string }) => void
        ) => {
          callback(null, {
            stdout:
              "SignatureKind: Developer (Self-signed certificate)\nPackageFullName: TestPackage",
            stderr: "",
          });
        }
      );

      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const result = await (teamsAppDriver as any).verifyPackageFamilyCertIsValid(
        "TestPackage_12345"
      );
      chai.assert.isFalse(result);
    });

    it("should handle PowerShell access denied errors", async () => {
      execStub.callsFake(
        (
          command: string,
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          callback: (error: NodeJS.ErrnoException | null, result: any) => void
        ) => {
          const accessError = new Error("Access denied") as NodeJS.ErrnoException;
          accessError.code = "EACCES";
          callback(accessError, null);
        }
      );

      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const result = await (teamsAppDriver as any).verifyPackageFamilyCertIsValid(
        "TestPackage_12345"
      );
      chai.assert.isFalse(result);
    });
  });
});
