// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { err, ok, Platform, TeamsAppManifest } from "@microsoft/teamsfx-api";
import AdmZip from "adm-zip";
import fs from "fs-extra";
import mockedEnv from "mocked-env";
import { v4 as uuid } from "uuid";
import { chai, expect, vi } from "vitest";
import { GraphClient } from "../../../../src/client/graphClient";
import { teamsDevPortalClient } from "../../../../src/client/teamsDevPortalClientProvider";
import { SovereignCloudEnvironment } from "../../../../src/common/accountUtils";
import { FeatureFlagName } from "../../../../src/common/featureFlags";
import { AppStudioError } from "../../../../src/component/driver/teamsApp/errors";
import { PublishingState } from "../../../../src/component/driver/teamsApp/interfaces/appdefinitions/IPublishingAppDefinition";
import { PublishAppPackageArgs } from "../../../../src/component/driver/teamsApp/interfaces/PublishAppPackageArgs";
import { PublishAppPackageDriver } from "../../../../src/component/driver/teamsApp/publishAppPackage";
import * as McpCertVerification from "../../../../src/component/driver/teamsApp/utils/McpCertVerification";
import { ODRProvider } from "../../../../src/component/utils/odrProvider";
import { UserCancelError } from "../../../../src/error/common";
import { MockedM365Provider } from "../../../core/utils";
import { MockedLogProvider, MockedUserInteraction } from "../../../plugins/solution/util";
import { Constants } from "./../../../../src/component/driver/teamsApp/constants";

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
    vi.restoreAllMocks();
    vi.restoreAllMocks();
    restoreEnv?.();
    restoreEnv = undefined;
  });

  it("skip publish in GCCH", async () => {
    restoreEnv = mockedEnv({
      [FeatureFlagName.SovereignCloudEnvironment]: SovereignCloudEnvironment.GCCH,
    });
    const publishTeamsAppSpy = vi.spyOn(teamsDevPortalClient, "publishTeamsApp");
    const pathExistsStub = vi.spyOn(fs, "pathExists");

    const args: PublishAppPackageArgs = {
      appPackagePath: "fakepath",
    };

    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isOk());
    expect(publishTeamsAppSpy).not.toHaveBeenCalled();
    expect(pathExistsStub).not.toHaveBeenCalled();
  });

  it("skip publish in DoD", async () => {
    restoreEnv = mockedEnv({
      [FeatureFlagName.SovereignCloudEnvironment]: SovereignCloudEnvironment.DOD,
    });
    const publishTeamsAppSpy = vi.spyOn(teamsDevPortalClient, "publishTeamsApp");
    const pathExistsStub = vi.spyOn(fs, "pathExists");

    const args: PublishAppPackageArgs = {
      appPackagePath: "fakepath",
    };

    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isOk());
    expect(publishTeamsAppSpy).not.toHaveBeenCalled();
    expect(pathExistsStub).not.toHaveBeenCalled();
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

    vi.spyOn(fs, "pathExists").mockResolvedValue(true);
    vi.spyOn(fs, "readFile").mockImplementation(async () => {
      const zip = new AdmZip();
      zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify(new TeamsAppManifest())));
      zip.addFile("color.png", Buffer.from(""));
      zip.addFile("outlie.png", Buffer.from(""));

      const archivedFile = zip.toBuffer();
      return archivedFile;
    });
    vi.spyOn(GraphClient.prototype, "getStagedApp").mockResolvedValue(undefined);
    vi.spyOn(GraphClient.prototype, "publishTeamsApp").mockResolvedValue(uuid());

    const result = await teamsAppDriver.execute(args, mockedDriverContext);
    chai.assert.isTrue(result.result.isOk());
  });

  it("should return token error when getAccessToken fails", async () => {
    const args: PublishAppPackageArgs = {
      appPackagePath: "fakepath",
    };

    vi.spyOn(fs, "pathExists").mockResolvedValue(true);
    vi.spyOn(fs, "readFile").mockImplementation(async () => {
      const zip = new AdmZip();
      const manifest = new TeamsAppManifest();
      manifest.id = uuid();
      manifest.name = { short: "test-app", full: "test-app" } as any;
      zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify(manifest)));
      return zip.toBuffer();
    });

    const tokenError = new UserCancelError();
    vi.spyOn(mockedDriverContext.m365TokenProvider, "getAccessToken").mockResolvedValue(
      err(tokenError)
    );

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

    vi.spyOn(fs, "pathExists").mockResolvedValue(true);
    vi.spyOn(fs, "readFile").mockImplementation(async () => {
      const zip = new AdmZip();
      const manifest = new TeamsAppManifest();
      manifest.id = uuid();
      manifest.name = { short: "test-app", full: "test-app" } as any;
      zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify(manifest)));
      return zip.toBuffer();
    });

    vi.spyOn(GraphClient.prototype, "getStagedApp").mockResolvedValue(undefined);
    vi.spyOn(GraphClient.prototype, "publishTeamsApp").mockRejectedValue(
      new Error("publish failed")
    );

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

    vi.spyOn(fs, "pathExists").mockResolvedValue(true);
    vi.spyOn(fs, "readFile").mockImplementation(async () => {
      const zip = new AdmZip();
      zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify(new TeamsAppManifest())));
      zip.addFile("color.png", Buffer.from(""));
      zip.addFile("outlie.png", Buffer.from(""));

      const archivedFile = zip.toBuffer();
      return archivedFile;
    });
    vi.spyOn(GraphClient.prototype, "getStagedApp").mockResolvedValue(state);
    vi.spyOn(mockedDriverContext.ui, "showMessage").mockResolvedValue(ok("Cancel"));

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

    vi.spyOn(fs, "pathExists").mockResolvedValue(true);
    vi.spyOn(fs, "readFile").mockImplementation(async () => {
      const zip = new AdmZip();
      zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify(new TeamsAppManifest())));
      zip.addFile("color.png", Buffer.from(""));
      zip.addFile("outlie.png", Buffer.from(""));

      const archivedFile = zip.toBuffer();
      return archivedFile;
    });
    vi.spyOn(GraphClient.prototype, "getStagedApp").mockResolvedValue(state);
    vi.spyOn(GraphClient.prototype, "publishTeamsAppUpdate").mockResolvedValue(uuid());
    vi.spyOn(mockedDriverContext.ui, "showMessage").mockResolvedValue(ok("Confirm"));

    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert.isTrue(result.isOk());
  });

  describe("MCP plugin certificate verification", () => {
    it("should pass when no declarative agents exist", async () => {
      const args: PublishAppPackageArgs = {
        appPackagePath: "fakepath",
      };

      vi.spyOn(fs, "pathExists").mockResolvedValue(true);
      vi.spyOn(fs, "readFile").mockImplementation(async () => {
        const zip = new AdmZip();
        const manifest = new TeamsAppManifest();
        zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify(manifest)));
        return zip.toBuffer();
      });
      vi.spyOn(GraphClient.prototype, "getStagedApp").mockResolvedValue(undefined);
      vi.spyOn(GraphClient.prototype, "publishTeamsApp").mockResolvedValue(uuid());

      const result = await teamsAppDriver.execute(args, mockedDriverContext);
      chai.assert.isTrue(result.result.isOk());
    });

    it("should pass when declarativeAgent has no actions", async () => {
      const args: PublishAppPackageArgs = {
        appPackagePath: "fakepath",
      };

      vi.spyOn(fs, "pathExists").mockResolvedValue(true);
      vi.spyOn(fs, "readFile").mockImplementation(async () => {
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
      vi.spyOn(GraphClient.prototype, "getStagedApp").mockResolvedValue(undefined);
      vi.spyOn(GraphClient.prototype, "publishTeamsApp").mockResolvedValue(uuid());

      const result = await teamsAppDriver.execute(args, mockedDriverContext);
      chai.assert.isTrue(result.result.isOk());
    });

    it("should pass when action has no runtimes", async () => {
      const args: PublishAppPackageArgs = {
        appPackagePath: "fakepath",
      };

      vi.spyOn(fs, "pathExists").mockResolvedValue(true);
      vi.spyOn(fs, "readFile").mockImplementation(async () => {
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
      vi.spyOn(GraphClient.prototype, "getStagedApp").mockResolvedValue(undefined);
      vi.spyOn(GraphClient.prototype, "publishTeamsApp").mockResolvedValue(uuid());

      const result = await teamsAppDriver.execute(args, mockedDriverContext);
      chai.assert.isTrue(result.result.isOk());
    });

    it("should pass when action has only non-LocalPlugin runtimes", async () => {
      const args: PublishAppPackageArgs = {
        appPackagePath: "fakepath",
      };

      vi.spyOn(fs, "pathExists").mockResolvedValue(true);
      vi.spyOn(fs, "readFile").mockImplementation(async () => {
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
      vi.spyOn(GraphClient.prototype, "getStagedApp").mockResolvedValue(undefined);
      vi.spyOn(GraphClient.prototype, "publishTeamsApp").mockResolvedValue(uuid());

      const result = await teamsAppDriver.execute(args, mockedDriverContext);
      chai.assert.isTrue(result.result.isOk());
    });

    it("should pass when LocalPlugin runtime is not MCP (no mcp:// prefix)", async () => {
      const args: PublishAppPackageArgs = {
        appPackagePath: "fakepath",
      };

      vi.spyOn(fs, "pathExists").mockResolvedValue(true);
      vi.spyOn(fs, "readFile").mockImplementation(async () => {
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
      vi.spyOn(GraphClient.prototype, "getStagedApp").mockResolvedValue(undefined);
      vi.spyOn(GraphClient.prototype, "publishTeamsApp").mockResolvedValue(uuid());

      const result = await teamsAppDriver.execute(args, mockedDriverContext);
      chai.assert.isTrue(result.result.isOk());
    });

    it("should pass when MCP server not found in ODR list (non-MCP local plugin)", async () => {
      const args: PublishAppPackageArgs = {
        appPackagePath: "fakepath",
      };

      vi.spyOn(fs, "pathExists").mockResolvedValue(true);
      vi.spyOn(fs, "readFile").mockImplementation(async () => {
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
      vi.spyOn(ODRProvider, "listServers").mockResolvedValue([]);
      vi.spyOn(GraphClient.prototype, "getStagedApp").mockResolvedValue(undefined);
      vi.spyOn(GraphClient.prototype, "publishTeamsApp").mockResolvedValue(uuid());

      const result = await teamsAppDriver.execute(args, mockedDriverContext);
      chai.assert.isTrue(result.result.isOk());
    });

    it("should pass when MCP plugin has valid certificate", async () => {
      const args: PublishAppPackageArgs = {
        appPackagePath: "fakepath",
      };

      vi.spyOn(fs, "pathExists").mockResolvedValue(true);
      vi.spyOn(fs, "readFile").mockImplementation(async () => {
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
      vi.spyOn(ODRProvider, "listServers").mockResolvedValue([
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
      vi.spyOn(McpCertVerification, "verifyLocalMCPPluginCerts").mockResolvedValue(true);
      vi.spyOn(GraphClient.prototype, "getStagedApp").mockResolvedValue(undefined);
      vi.spyOn(GraphClient.prototype, "publishTeamsApp").mockResolvedValue(uuid());

      const result = await teamsAppDriver.execute(args, mockedDriverContext);
      chai.assert.isTrue(result.result.isOk());
    });

    it("should fail when MCP plugin has self-signed certificate", async () => {
      const args: PublishAppPackageArgs = {
        appPackagePath: "fakepath",
      };

      vi.spyOn(fs, "pathExists").mockResolvedValue(true);
      vi.spyOn(fs, "readFile").mockImplementation(async () => {
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
      vi.spyOn(ODRProvider, "listServers").mockResolvedValue([
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
      vi.spyOn(McpCertVerification, "verifyLocalMCPPluginCerts").mockResolvedValue(false);

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

      vi.spyOn(fs, "pathExists").mockResolvedValue(true);
      vi.spyOn(fs, "readFile").mockImplementation(async () => {
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
      vi.spyOn(ODRProvider, "listServers").mockResolvedValue([
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
      vi.spyOn(McpCertVerification, "verifyLocalMCPPluginCerts").mockResolvedValue(false);

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

      vi.spyOn(fs, "pathExists").mockResolvedValue(true);
      vi.spyOn(fs, "readFile").mockImplementation(async () => {
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
      vi.spyOn(ODRProvider, "listServers").mockResolvedValue([
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
      vi.spyOn(McpCertVerification, "verifyLocalMCPPluginCerts").mockResolvedValue(true);
      vi.spyOn(GraphClient.prototype, "getStagedApp").mockResolvedValue(undefined);
      vi.spyOn(GraphClient.prototype, "publishTeamsApp").mockResolvedValue(uuid());

      const result = await teamsAppDriver.execute(args, mockedDriverContext);
      chai.assert.isTrue(result.result.isOk());
    });
  });

  // eslint-disable-next-line no-secrets/no-secrets
  describe("verifyPackageFamilyCertIsValid", () => {
    afterEach(() => {
      vi.restoreAllMocks();
    });

    for (const signature of ["Store", "System", "Enterprise"]) {
      it(`accepts the ${signature} signature`, async () => {
        vi.spyOn(McpCertVerification.mcpCertDeps, "execFile").mockResolvedValue({
          stdout: ` \r\n${signature}\r\n `,
          stderr: "",
        });

        expect(await McpCertVerification.verifyPackageFamilyCertIsValid("TestPackage_12345")).toBe(
          true
        );
      });
    }

    for (const stdout of [
      "",
      " \r\n ",
      "Developer",
      "DEVELOPER",
      "None",
      "Unknown",
      "NotStore",
      "Store\nDeveloper",
      "Store\nStore",
      "Get-AppxPackage : Access denied",
    ]) {
      it(`rejects missing or untrusted signature output ${JSON.stringify(stdout)}`, async () => {
        vi.spyOn(McpCertVerification.mcpCertDeps, "execFile").mockResolvedValue({
          stdout,
          stderr: "",
        });

        expect(await McpCertVerification.verifyPackageFamilyCertIsValid("TestPackage_12345")).toBe(
          false
        );
      });
    }

    it("passes registry metadata only as environment data, never as command text", async () => {
      const execFileStub = vi
        .spyOn(McpCertVerification.mcpCertDeps, "execFile")
        .mockResolvedValue({ stdout: "", stderr: "" });
      const packageNames = ["MyApp_abc123xyz", "Registry_'\";|&$()`\r\n_value"];

      for (const packageName of packageNames) {
        await McpCertVerification.verifyPackageFamilyCertIsValid(packageName);
      }

      expect(execFileStub).toHaveBeenCalledTimes(2);
      const [file, args, options] = execFileStub.mock.calls[0];
      expect(file).toBe("powershell.exe");
      expect(args).toEqual([
        "-NoProfile",
        "-NonInteractive",
        "-Command",
        expect.stringContaining(
          "Where-Object { $_.PackageFamilyName -eq $env:ATK_MCP_PACKAGE_FAMILY }"
        ),
      ]);
      expect(args?.join(" ")).not.toContain(packageNames[0]);
      expect(execFileStub.mock.calls[1][1]).toEqual(args);
      expect(options).toMatchObject({
        shell: false,
        encoding: "utf8",
        env: { ATK_MCP_PACKAGE_FAMILY: packageNames[0] },
      });
      expect(execFileStub.mock.calls[1][2]).toMatchObject({
        env: { ATK_MCP_PACKAGE_FAMILY: packageNames[1] },
      });
      expect(args?.join(" ")).toContain("Get-AppxPackage -ErrorAction Stop");
      expect(args?.join(" ")).toContain("Select-Object -ExpandProperty SignatureKind");
    });

    it.skipIf(process.platform !== "win32")(
      "treats registry metacharacters literally in the real PowerShell query",
      async () => {
        const packageFamily = "Registry_'\";|&$()`_value";
        const servers = ODRProvider.parseODRListOutput({
          servers: [
            {
              name: "test-server",
              packages: [{ identifier: "test-server" }],
              _meta: {
                "io.modelcontextprotocol.registry/publisher-provided": {
                  "com.microsoft.windows": {
                    manifest: {
                      _meta: { "com.microsoft.windows": { package_family_name: packageFamily } },
                      server: { mcp_config: { command: "odr", args: [] } },
                    },
                  },
                },
              },
            },
          ],
        });
        vi.spyOn(ODRProvider, "listServers").mockResolvedValue(servers);
        const zip = new AdmZip();
        zip.addFile(
          "plugin.json",
          Buffer.from(
            JSON.stringify({
              runtimes: [{ type: "LocalPlugin", spec: { local_endpoint: "mcp://test-server" } }],
            })
          )
        );
        const plugin = zip.getEntry("plugin.json");
        if (!plugin) throw new Error("Missing test plugin");

        const execFile = McpCertVerification.mcpCertDeps.execFile;
        vi.spyOn(McpCertVerification.mcpCertDeps, "execFile").mockImplementation(
          (file, args, options) => {
            const script = args?.at(-1);
            if (typeof script !== "string") throw new Error("Missing PowerShell script");
            // Replace only the OS package source, retaining the production query and transport.
            const packageSource =
              "function Get-AppxPackage { " +
              "[pscustomobject]@{ PackageFamilyName = $env:ATK_MCP_PACKAGE_FAMILY; SignatureKind = 'Store' }; " +
              "[pscustomobject]@{ PackageFamilyName = 'UnrelatedPackage'; SignatureKind = 'Developer' } }; ";
            return execFile(
              file,
              ["-NoProfile", "-NonInteractive", "-Command", packageSource + script],
              options
            );
          }
        );

        expect(await McpCertVerification.verifyLocalMCPPluginCerts(plugin)).toBe(true);
      }
    );

    for (const message of ["PowerShell execution failed", "Command timed out", "Access denied"]) {
      it(`fails closed and logs when querying fails: ${message}`, async () => {
        const error = new Error(message);
        vi.spyOn(McpCertVerification.mcpCertDeps, "execFile").mockRejectedValue(error);
        const logStub = vi.spyOn(console, "error").mockImplementation(() => {});

        expect(await McpCertVerification.verifyPackageFamilyCertIsValid("TestPackage_12345")).toBe(
          false
        );
        expect(logStub).toHaveBeenCalledWith("Unable to get cert info for package name", error);
      });
    }
  });
});
