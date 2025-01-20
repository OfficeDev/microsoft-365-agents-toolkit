// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  err,
  FxError,
  Inputs,
  ok,
  Platform,
  Result,
  Stage,
  UserError,
  Void,
} from "@microsoft/teamsfx-api";
import {
  DependencyStatus,
  DepsManager,
  NodejsNotFoundError,
  QuestionNames,
  SyncManifestInputs,
  teamsDevPortalClient,
  TestToolInstallOptions,
} from "@microsoft/teamsfx-core";
import { assert } from "chai";
import "mocha";
import sinon from "sinon";
import { Duplex } from "stream";
import { CancellationToken, createMessageConnection } from "vscode-jsonrpc";
import { setFunc } from "../src/customizedFuncAdapter";
import ServerConnection from "../src/serverConnection";
import { SyncManifestInputsForVS } from "@microsoft/teamsfx-core/build/component/driver/teamsApp/interfaces/SyncManifest";

class TestStream extends Duplex {
  _write(chunk: string, _encoding: string, done: () => void) {
    this.emit("data", chunk);
    done();
  }

  _read(_size: number) {}
}

describe("serverConnections", () => {
  const sandbox = sinon.createSandbox();
  const up = new TestStream();
  const down = new TestStream();
  const msgConn = createMessageConnection(up as any, down as any);

  afterEach(() => {
    sandbox.restore();
  });

  it("connection", () => {
    const connection = new ServerConnection(msgConn);
    assert.equal(connection["connection"], msgConn);
  });

  it("listen", () => {
    const stub = sandbox.stub(msgConn, "listen");
    const connection = new ServerConnection(msgConn);
    connection.listen();
    assert.isTrue(stub.calledOnce);
  });

  it("getQuestionsRequest", () => {
    const connection = new ServerConnection(msgConn);
    const fake = sandbox.fake.returns(ok(undefined));
    sandbox.replace(connection["core"], "getQuestions", fake);
    const stage = Stage.create;
    const inputs = { platform: Platform.VS };
    const token = {};
    const res = connection.getQuestionsRequest(stage, inputs as Inputs, token as CancellationToken);
    assert.isTrue(res.isOk());
    if (res.isOk()) {
      assert.isUndefined(res.value);
    }
  });

  it("createProjectRequest", () => {
    const connection = new ServerConnection(msgConn);
    const fake = sandbox.fake.returns({ projectPath: "test" });
    sandbox.replace(connection["core"], "createProject", fake);
    const inputs = {
      platform: "vs",
    };
    const token = {};
    const res = connection.createProjectRequest(inputs as Inputs, token as CancellationToken);
    res.then((data) => {
      assert.equal(data, ok({ projectPath: "test" }));
    });
  });

  it("localDebugRequest", () => {
    const connection = new ServerConnection(msgConn);
    const fake = sandbox.fake.returns("test");
    sandbox.replace(connection["core"], "localDebug", fake);
    const inputs = {
      platform: "vs",
    };
    const token = {};
    const res = connection.localDebugRequest(inputs as Inputs, token as CancellationToken);
    res.then((data) => {
      assert.equal(data, ok(undefined));
    });
  });

  it("preProvisionResourcesRequest", () => {
    const connection = new ServerConnection(msgConn);
    const fake = sandbox.fake.returns({
      needAzureLogin: true,
      needM365Login: true,
      resolvedAzureSubscriptionId: undefined,
      resolvedAzureResourceGroupName: undefined,
    });
    sandbox.replace(connection["core"], "preProvisionForVS", fake);
    const inputs = {
      platform: "vs",
    };
    const token = {};
    const res = connection.preProvisionResourcesRequest(
      inputs as Inputs,
      token as CancellationToken
    );
    res.then((data) => {
      assert.equal(
        data,
        ok({
          needAzureLogin: true,
          needM365Login: true,
          resolvedAzureSubscriptionId: undefined,
          resolvedAzureResourceGroupName: undefined,
        })
      );
    });
  });

  it("provisionResourcesRequest", () => {
    const connection = new ServerConnection(msgConn);
    const fake = sandbox.fake.returns("test");
    sandbox.replace(connection["core"], "provisionResources", fake);
    const inputs = {
      platform: "vs",
    };
    const token = {};
    const res = connection.provisionResourcesRequest(inputs as Inputs, token as CancellationToken);
    res.then((data) => {
      assert.equal(data, ok(undefined));
    });
  });

  it("preCheckYmlAndEnvForVSRequest", () => {
    const connection = new ServerConnection(msgConn);
    const fake = sandbox.fake.returns("test");
    sandbox.replace(connection["core"], "preCheckYmlAndEnvForVS", fake);
    const inputs = {
      platform: "vs",
    };
    const token = {};
    const res = connection.preCheckYmlAndEnvForVSRequest(
      inputs as Inputs,
      token as CancellationToken
    );
    res.then((data) => {
      assert.equal(data, ok(undefined));
    });
  });

  it("validateManifestForVSRequest", () => {
    const connection = new ServerConnection(msgConn);
    const fake = sandbox.fake.returns("test");
    sandbox.replace(connection["core"], "validateManifest", fake);
    const inputs = {
      platform: "vs",
    };
    const token = {};
    const res = connection.validateManifestForVSRequest(
      inputs as Inputs,
      token as CancellationToken
    );
    res.then((data) => {
      assert.equal(data, ok(undefined));
    });
  });

  it("deployArtifactsRequest", () => {
    const connection = new ServerConnection(msgConn);
    const fake = sandbox.fake.returns("test");
    sandbox.replace(connection["core"], "deployArtifacts", fake);
    const inputs = {
      platform: "vs",
    };
    const token = {};
    const res = connection.deployArtifactsRequest(inputs as Inputs, token as CancellationToken);
    res.then((data) => {
      assert.equal(data, ok(undefined));
    });
  });

  it("deployTeamsAppManifestRequest - v3", async () => {
    const connection = new ServerConnection(msgConn);
    const fake = sandbox.fake.resolves(ok("test"));
    sandbox.replace(connection["core"], "deployTeamsManifest", fake);
    const inputs = {
      platform: "vs",
    };
    const token = {};
    const res = await connection.deployTeamsAppManifestRequest(
      inputs as Inputs,
      token as CancellationToken
    );
    assert.isTrue(res.isOk());
    if (res.isOk()) {
      assert.deepEqual(res.value, {});
    }
    sandbox.restore();
  });

  it("buildArtifactsRequest - V3", () => {
    const connection = new ServerConnection(msgConn);
    const fake = sandbox.fake.resolves(ok("test"));
    sandbox.replace(connection["core"], "createAppPackage", fake);
    const inputs = {
      platform: "vs",
      projectPath: ".",
    };
    const token = {};
    const res = connection.buildArtifactsRequest(inputs as Inputs, token as CancellationToken);
    res.then((data) => {
      assert.equal(data, ok("test"));
    });
  });

  it("publishApplicationRequest", () => {
    const connection = new ServerConnection(msgConn);
    const fake = sandbox.fake.returns("test");
    sandbox.replace(connection["core"], "publishApplication", fake);
    const inputs = {
      platform: "vs",
    };
    const token = {};
    const res = connection.publishApplicationRequest(inputs as Inputs, token as CancellationToken);
    res.then((data) => {
      assert.equal(data, ok(undefined));
    });
  });

  it("getLaunchUrlRequest", () => {
    const connection = new ServerConnection(msgConn);
    const fake = sandbox.fake.returns("test");
    sandbox.replace(connection["core"], "previewWithManifest", fake);
    const inputs = {
      platform: "vs",
    };
    const token = {};
    const res = connection.getLaunchUrlRequest(inputs as Inputs, token as CancellationToken);
    res.then((data) => {
      assert.equal(data, ok("test"));
    });
  });

  it("customizeLocalFuncRequest", () => {
    const connection = new ServerConnection(msgConn);
    const fake = sandbox.fake.returns("test");
    const id = setFunc(fake);
    const inputs = {
      platform: "vs",
    };
    const token = {};
    const res = connection.customizeLocalFuncRequest(
      id,
      inputs as Inputs,
      token as CancellationToken
    );
    res.then((data) => {
      assert.equal(data, ok(undefined));
    });
  });

  it("customizeValidateFuncRequest", () => {
    const connection = new ServerConnection(msgConn);
    const fake = sandbox.fake.returns("test");
    const id = setFunc(fake);
    const inputs = {
      platform: "vs",
    };
    const token = {};
    const res = connection.customizeValidateFuncRequest(
      id,
      inputs,
      inputs as Inputs,
      token as CancellationToken
    );
    res.then((data) => {
      assert.equal(data, ok("test"));
    });
  });

  it("customizeOnSelectionChangeFuncRequest", () => {
    const connection = new ServerConnection(msgConn);
    const fake = sandbox.fake.returns("test");
    const id = setFunc(fake);
    const inputs = new Set<string>("test");
    const token = {};
    const res = connection.customizeOnSelectionChangeFuncRequest(
      id,
      inputs,
      inputs,
      token as CancellationToken
    );
    res.then((data) => {
      assert.equal(data, ok("test"));
    });
  });

  it("getSideloadingStatusRequest", () => {
    const connection = new ServerConnection(msgConn);
    const accountToken = {
      token: "test token",
    };
    const cancelToken = {};
    const res = connection.getSideloadingStatusRequest(
      accountToken,
      cancelToken as CancellationToken
    );
    res.then((data) => {
      assert.equal(data, ok("undefined"));
    });
  });

  it("getCopilotStatusRequest", () => {
    const connection = new ServerConnection(msgConn);
    const accountToken = {
      token: "test token",
    };
    const cancelToken = {};
    const res = connection.getCopilotStatusRequest(accountToken, cancelToken as CancellationToken);
    res.then((data) => {
      assert.equal(data, ok("undefined"));
    });
  });

  it("addSsoRequest", () => {
    const connection = new ServerConnection(msgConn);
    const fake = sandbox.fake.returns("test");
    sandbox.replace(connection["core"], "createProject", fake);

    const inputs = {
      platform: "vs",
    };
    const token = {};
    const res = connection.addSsoRequest(inputs as Inputs, token as CancellationToken);
    res.then((data) => {
      assert.equal(data, ok("test"));
    });
  });

  it("getProjectMigrationStatusRequest", () => {
    const connection = new ServerConnection(msgConn);
    const fake = sandbox.fake.returns({
      currentVersion: "3.0.0",
      isSupport: 0,
      trackingId: "1234-3213-4325-1231",
    });
    sandbox.replace(connection["core"], "projectVersionCheck", fake);

    const inputs = {
      platform: "vs",
    };
    const token = {};
    const res = connection.getProjectMigrationStatusRequest(
      inputs as Inputs,
      token as CancellationToken
    );
    res.then((data) => {
      assert.equal(data.isOk(), true);
    });
  });

  it("migrateProjectRequest - ok(true)", async () => {
    const connection = new ServerConnection(msgConn);
    sandbox.replace(connection["core"], "phantomMigrationV3", sandbox.fake.returns(Void));
    connection
      .migrateProjectRequest(
        {
          platform: "vs",
        } as Inputs,
        {} as CancellationToken
      )
      .then((data) => {
        assert.equal(data, ok(true));
      });
  });

  it("migrateProjectRequest - ok(false)", async () => {
    const connection = new ServerConnection(msgConn);
    sandbox.replace(connection["core"], "phantomMigrationV3", sandbox.fake.returns("test"));
    connection
      .migrateProjectRequest(
        {
          platform: "vs",
        } as Inputs,
        {} as CancellationToken
      )
      .then((data) => {
        assert.equal(data, ok(false));
      });
  });

  it("publishInDeveloperPortalRequest", () => {
    const connection = new ServerConnection(msgConn);
    const fake = sandbox.fake.resolves(ok(Void));
    sandbox.replace(connection["core"], "publishInDeveloperPortal", fake);
    const inputs = {
      platform: "vs",
      projectPath: "test",
      appPackage: "appPackage",
    };
    const token = {};
    const res = connection.publishInDeveloperPortalRequest(
      inputs as Inputs,
      token as CancellationToken
    );
    res.then((data) => {
      assert.equal(data.isOk(), true);
    });
  });

  it("setRegionRequest", () => {
    const connection = new ServerConnection(msgConn);
    const accountToken = {
      token: "fakeToken",
    };
    sinon.stub(teamsDevPortalClient, "setRegionEndpointByToken").callsFake(async () => {});

    const res = connection.setRegionRequest(accountToken, {} as CancellationToken);
    res.then((data) => {
      assert.equal(data.isOk(), true);
    });
  });

  it("listDevTunnelsRequest fail with wrong token", async () => {
    const connection = new ServerConnection(msgConn);
    const fake = sandbox.fake.returns("test");
    const inputs = {
      platform: "vs",
      devTunnelToken: "token",
    };
    const token = {};
    const res = await connection.listDevTunnelsRequest(
      inputs as Inputs,
      token as CancellationToken
    );
    assert.isTrue(res.isErr());
  });

  it("copilotPluginListOperations fail", async () => {
    const connection = new ServerConnection(msgConn);
    const fake = sandbox.fake.resolves(err(new UserError("source", "name", "", "")));
    sandbox.replace(connection["core"], "copilotPluginListOperations", fake);
    const res = await connection.listOpenAPISpecOperationsRequest(
      {} as Inputs,
      {} as CancellationToken
    );
    assert.isTrue(res.isErr());
    if (res.isErr()) {
      assert.equal(res.error.source, "source");
    }
  });

  it("copilotPluginAddAPIRequest", async () => {
    const connection = new ServerConnection(msgConn);
    const fake = sandbox.fake.resolves(ok(undefined));
    sandbox.replace(connection["core"], "copilotPluginAddAPI", fake);
    const res = await connection.copilotPluginAddAPIRequest({} as Inputs, {} as CancellationToken);
    assert.isTrue(res.isOk());
  });

  it("checkAndInstallTestTool", async () => {
    const connection = new ServerConnection(msgConn);
    sandbox.stub(DepsManager.prototype, "ensureDependency").resolves({
      isInstalled: true,
      command: "mock command",
      details: { installVersion: "mock version", binFolders: ["mock bin folders"] },
    } as DependencyStatus);
    const res = await connection.checkAndInstallTestTool(
      {} as TestToolInstallOptions & { correlationId: string },
      {} as CancellationToken
    );
    assert.isTrue(res.isOk());
    assert.deepEqual(res._unsafeUnwrap(), {
      isInstalled: true,
      command: "mock command",
      details: { installVersion: "mock version", binFolders: ["mock bin folders"] },
    });
  });
  it("checkAndInstallTestTool DepenendencyStatus error", async () => {
    const connection = new ServerConnection(msgConn);
    const err = new NodejsNotFoundError();
    sandbox.stub(DepsManager.prototype, "ensureDependency").resolves({
      isInstalled: true,
      command: "mock command",
      details: {},
      error: err,
    } as DependencyStatus);
    const res = await connection.checkAndInstallTestTool(
      {} as TestToolInstallOptions & { correlationId: string },
      {} as CancellationToken
    );
    assert.isTrue(res.isOk());
    assert.deepEqual(res._unsafeUnwrap(), {
      isInstalled: true,
      command: "mock command",
      details: {},
      error: {
        message: err.message,
        helpLink: err.helpLink,
      },
    });
  });
  it("checkAndInstallTestTool error", async () => {
    const connection = new ServerConnection(msgConn);
    sandbox.stub(DepsManager.prototype, "ensureDependency").rejects("MockError");
    const res = await connection.checkAndInstallTestTool(
      {} as TestToolInstallOptions & { correlationId: string },
      {} as CancellationToken
    );
    assert.isFalse(res.isOk());
    assert.match(res._unsafeUnwrapErr().message, /MockError/);
  });

  it("listPluginApiSpecs fail", async () => {
    const connection = new ServerConnection(msgConn);
    const fake = sandbox.fake.resolves(err("error"));
    sandbox.replace(connection["core"], "listPluginApiSpecs", fake);
    const res = await connection.listPluginApiSpecs({} as Inputs, {} as CancellationToken);
    assert.isTrue(res.isErr());
  });

  it("listPluginApiSpecsRequest", async () => {
    const connection = new ServerConnection(msgConn);
    const fake = sandbox.fake.resolves(ok(undefined));
    sandbox.replace(connection["core"], "listPluginApiSpecs", fake);
    const res = await connection.listPluginApiSpecs({} as Inputs, {} as CancellationToken);
    assert.isTrue(res.isOk());
  });

  it("syncTeamsAppManifestRequest", async () => {
    const connection = new ServerConnection(msgConn);
    sandbox
      .stub(connection["core"], "syncManifest")
      .callsFake((inputs: SyncManifestInputs): Promise<Result<any, FxError>> => {
        if (!inputs[QuestionNames.TeamsAppId]) {
          return Promise.resolve(ok(undefined));
        } else {
          return Promise.resolve(err(new UserError("source", "name", "", "")));
        }
      });
    const res = await connection.syncTeamsAppManifestRequest(
      {
        correlationId: "123",
      } as SyncManifestInputsForVS,
      {} as CancellationToken
    );
    assert.isTrue(res.isOk());
  });

  it("syncTeamsAppManifestRequest with teamsApp Id", async () => {
    const connection = new ServerConnection(msgConn);
    sandbox
      .stub(connection["core"], "syncManifest")
      .callsFake((inputs: SyncManifestInputs): Promise<Result<any, FxError>> => {
        if (inputs[QuestionNames.TeamsAppId] === "123") {
          return Promise.resolve(ok(undefined));
        } else {
          return Promise.resolve(err(new UserError("source", "name", "", "")));
        }
      });
    const res = await connection.syncTeamsAppManifestRequest(
      {
        teamsAppFromTdp: {
          teamsAppId: "123",
        },
      } as SyncManifestInputsForVS,
      {} as CancellationToken
    );
    assert.isTrue(res.isOk());
  });
  it("syncTeamsAppManifestRequest with teamsAppFromTdp but no teamsAppId", async () => {
    const connection = new ServerConnection(msgConn);
    sandbox
      .stub(connection["core"], "syncManifest")
      .callsFake((inputs: SyncManifestInputs): Promise<Result<any, FxError>> => {
        if (!inputs[QuestionNames.TeamsAppId]) {
          return Promise.resolve(ok(undefined));
        } else {
          return Promise.resolve(err(new UserError("source", "name", "", "")));
        }
      });
    const res = await connection.syncTeamsAppManifestRequest(
      {
        teamsAppFromTdp: {},
      } as SyncManifestInputsForVS,
      {} as CancellationToken
    );
    assert.isTrue(res.isOk());
  });
  it("isDeclarativeAgentRequest - success", async () => {
    const connection = new ServerConnection(msgConn);
    sandbox
      .stub(connection["core"], "isDelcarativeAgentApp")
      .callsFake((inputs: Inputs): Promise<Result<any, FxError>> => {
        return Promise.resolve(ok(true));
      });
    const res = await connection.isDeclarativeAgentRequest(
      {
        correlationId: "123",
      } as Inputs,
      {} as CancellationToken
    );
    assert.isTrue(res.isOk());
    if (res.isOk()) {
      assert.equal(res.value, true);
    }
  });

  it("isDeclarativeAgentRequest - failed", async () => {
    const connection = new ServerConnection(msgConn);
    sandbox
      .stub(connection["core"], "isDelcarativeAgentApp")
      .callsFake((inputs: Inputs): Promise<Result<any, FxError>> => {
        return Promise.resolve(err(new UserError("source", "name", "", "")));
      });
    const res = await connection.isDeclarativeAgentRequest(
      {
        correlationId: "123",
      } as Inputs,
      {} as CancellationToken
    );
    assert.isTrue(res.isErr());
  });
});
