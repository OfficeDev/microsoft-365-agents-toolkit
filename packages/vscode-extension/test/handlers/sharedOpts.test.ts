import * as sinon from "sinon";
import * as chai from "chai";
import * as uuid from "uuid";
import * as globalVariables from "../../src/globalVariables";
import * as systemEnvUtils from "../../src/utils/systemEnvUtils";
import * as vscode from "vscode";
import * as telemetryUtils from "../../src/utils/telemetryUtils";
import {
  Platform,
  Stage,
  err,
  UserError,
  Inputs,
  ok,
  Result,
  FxError,
} from "@microsoft/teamsfx-api";
import { processResult, runCommand } from "../../src/handlers/sharedOpts";
import { ExtTelemetry } from "../../src/telemetry/extTelemetry";
import { MockCore } from "../mocks/mockCore";
import { RecommendedOperations } from "../../src/debug/common/debugConstants";
import { UserCancelError } from "@microsoft/teamsfx-core";
import { TelemetryEvent } from "../../src/telemetry/extTelemetryEvents";

describe("SharedOpts", () => {
  describe("runCommand()", function () {
    const sandbox = sinon.createSandbox();

    beforeEach(() => {
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      sandbox.stub(ExtTelemetry, "sendTelemetryErrorEvent");
    });

    afterEach(() => {
      sandbox.restore();
    });

    it("create sample with projectid", async () => {
      sandbox.restore();
      sandbox.stub(globalVariables, "core").value(new MockCore());
      const sendTelemetryEvent = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      sandbox.stub(ExtTelemetry, "sendTelemetryErrorEvent");
      const createProject = sandbox.spy(globalVariables.core, "createProject");
      sandbox.stub(vscode.commands, "executeCommand");
      const inputs = { projectId: uuid.v4(), platform: Platform.VSCode };

      await runCommand(Stage.create, inputs);

      sinon.assert.calledOnce(createProject);
      chai.assert.isTrue(createProject.args[0][0].projectId != undefined);
      chai.assert.isTrue(sendTelemetryEvent.args[0][1]!["new-project-id"] != undefined);
    });

    it("create from scratch without projectid", async () => {
      sandbox.restore();
      sandbox.stub(globalVariables, "core").value(new MockCore());
      const sendTelemetryEvent = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      sandbox.stub(ExtTelemetry, "sendTelemetryErrorEvent");
      const createProject = sandbox.spy(globalVariables.core, "createProject");
      sandbox.stub(vscode.commands, "executeCommand");

      await runCommand(Stage.create);
      sinon.assert.calledOnce(createProject);
      chai.assert.isTrue(createProject.args[0][0].projectId != undefined);
      chai.assert.isTrue(sendTelemetryEvent.args[0][1]!["new-project-id"] != undefined);
    });

    it("provisionResources", async () => {
      sandbox.stub(globalVariables, "core").value(new MockCore());
      const provisionResources = sandbox.spy(globalVariables.core, "provisionResources");

      await runCommand(Stage.provision);
      sinon.assert.calledOnce(provisionResources);
    });
    it("deployTeamsManifest", async () => {
      sandbox.stub(globalVariables, "core").value(new MockCore());
      const deployTeamsManifest = sandbox.spy(globalVariables.core, "deployTeamsManifest");

      await runCommand(Stage.deployTeams);
      sinon.assert.calledOnce(deployTeamsManifest);
    });
    it("addWebpart", async () => {
      sandbox.stub(globalVariables, "core").value(new MockCore());
      const addWebpart = sandbox.spy(globalVariables.core, "addWebpart");

      await runCommand(Stage.addWebpart);
      sinon.assert.calledOnce(addWebpart);
    });
    it("createAppPackage", async () => {
      sandbox.stub(globalVariables, "core").value(new MockCore());
      const createAppPackage = sandbox.spy(globalVariables.core, "createAppPackage");

      await runCommand(Stage.createAppPackage);
      sinon.assert.calledOnce(createAppPackage);
    });
    it("error", async () => {
      sandbox.stub(globalVariables, "core").value(new MockCore());
      try {
        await runCommand("none" as any);
        sinon.assert.fail("should not reach here");
      } catch (e) {}
    });
    it("provisionResources - local", async () => {
      const mockCore = new MockCore();
      const mockCoreStub = sandbox
        .stub(mockCore, "provisionResources")
        .resolves(err(new UserError("test", "test", "test")));
      sandbox.stub(globalVariables, "core").value(mockCore);

      const res = await runCommand(Stage.provision, {
        platform: Platform.VSCode,
        env: "local",
      } as Inputs);
      chai.assert.isTrue(res.isErr());
      if (res.isErr()) {
        chai.assert.equal(res.error.recommendedOperation, RecommendedOperations.DebugInTestTool);
      }
      sinon.assert.calledOnce(mockCoreStub);
    });

    it("deployArtifacts", async () => {
      sandbox.stub(globalVariables, "core").value(new MockCore());
      const deployArtifacts = sandbox.spy(globalVariables.core, "deployArtifacts");

      await runCommand(Stage.deploy);
      sinon.assert.calledOnce(deployArtifacts);
    });

    it("deployArtifacts - local", async () => {
      const mockCore = new MockCore();
      const mockCoreStub = sandbox
        .stub(mockCore, "deployArtifacts")
        .resolves(err(new UserError("test", "test", "test")));
      sandbox.stub(globalVariables, "core").value(mockCore);

      await runCommand(Stage.deploy, {
        platform: Platform.VSCode,
        env: "local",
      } as Inputs);
      sinon.assert.calledOnce(mockCoreStub);
    });

    it("deployAadManifest", async () => {
      sandbox.stub(globalVariables, "core").value(new MockCore());
      const deployAadManifest = sandbox.spy(globalVariables.core, "deployAadManifest");
      const input: Inputs = systemEnvUtils.getSystemInputs();
      await runCommand(Stage.deployAad, input);

      sandbox.assert.calledOnce(deployAadManifest);
    });

    it("deployAadManifest happy path", async () => {
      sandbox.stub(globalVariables, "core").value(new MockCore());
      sandbox.stub(globalVariables.core, "deployAadManifest").resolves(ok(undefined));
      const input: Inputs = systemEnvUtils.getSystemInputs();
      const res = await runCommand(Stage.deployAad, input);
      chai.assert.isTrue(res.isOk());
      if (res.isOk()) {
        chai.assert.strictEqual(res.value, undefined);
      }
    });

    it("localDebug", async () => {
      sandbox.stub(globalVariables, "core").value(new MockCore());

      let ignoreEnvInfo: boolean | undefined = undefined;
      let localDebugCalled = 0;
      sandbox
        .stub(globalVariables.core, "localDebug")
        .callsFake(async (inputs: Inputs): Promise<Result<undefined, FxError>> => {
          ignoreEnvInfo = inputs.ignoreEnvInfo;
          localDebugCalled += 1;
          return ok(undefined);
        });

      await runCommand(Stage.debug);
      chai.expect(ignoreEnvInfo).to.equal(false);
      chai.expect(localDebugCalled).equals(1);
    });

    it("publishApplication", async () => {
      sandbox.stub(globalVariables, "core").value(new MockCore());
      const publishApplication = sandbox.spy(globalVariables.core, "publishApplication");

      await runCommand(Stage.publish);
      sinon.assert.calledOnce(publishApplication);
    });

    it("createEnv", async () => {
      sandbox.stub(globalVariables, "core").value(new MockCore());
      const createEnv = sandbox.spy(globalVariables.core, "createEnv");
      sandbox.stub(vscode.commands, "executeCommand");

      await runCommand(Stage.createEnv);
      sinon.assert.calledOnce(createEnv);
    });
    it("syncManifest", async () => {
      sandbox.stub(globalVariables, "core").value(new MockCore());
      const syncManifest = sandbox.spy(globalVariables.core, "syncManifest");
      sandbox.stub(vscode.commands, "executeCommand");

      await runCommand(Stage.syncManifest);
      sinon.assert.calledOnce(syncManifest);
    });
  });

  describe("processResult", () => {
    const sandbox = sinon.createSandbox();

    beforeEach(() => {
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      sandbox.stub(ExtTelemetry, "sendTelemetryErrorEvent");
    });

    afterEach(() => {
      sandbox.restore();
    });

    it("UserCancelError", async () => {
      sandbox.stub(telemetryUtils, "getTeamsAppTelemetryInfoByEnv").resolves({
        appId: "mockId",
        tenantId: "mockTenantId",
      });
      await processResult("", err(new UserCancelError()), {
        platform: Platform.VSCode,
        env: "dev",
      });
    });
    it("CreateNewEnvironment", async () => {
      await processResult(TelemetryEvent.CreateNewEnvironment, ok(null), {
        platform: Platform.VSCode,
        sourceEnvName: "dev",
        targetEnvName: "dev1",
      });
    });
  });
});
