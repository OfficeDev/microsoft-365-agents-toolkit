import { LogLevel } from "@microsoft/teamsfx-api";
import * as chai from "chai";
import * as sinon from "sinon";
import * as vscode from "vscode";
import VsCodeLogInstance from "../../src/commonlib/log";
import { configMgr } from "../../src/config";
import { ExtTelemetry } from "../../src/telemetry/extTelemetry";

describe("configMgr", () => {
  const sanbox = sinon.createSandbox();
  describe("loadLogLevel", () => {
    afterEach(async () => {
      sanbox.restore();
    });
    it("Debug", () => {
      sanbox.stub(vscode.workspace, "getConfiguration").returns({
        get: () => {
          return "Debug";
        },
      } as any);
      configMgr.loadLogLevel();
      chai.assert.equal(VsCodeLogInstance.logLevel, LogLevel.Debug);
    });

    it("Verbose", () => {
      sanbox.stub(vscode.workspace, "getConfiguration").returns({
        get: () => {
          return "Verbose";
        },
      } as any);
      configMgr.loadLogLevel();
      chai.assert.equal(VsCodeLogInstance.logLevel, LogLevel.Verbose);
    });

    it("Info", () => {
      sanbox.stub(vscode.workspace, "getConfiguration").returns({
        get: () => {
          return "Info";
        },
      } as any);
      configMgr.loadLogLevel();
      chai.assert.equal(VsCodeLogInstance.logLevel, LogLevel.Info);
    });
  });

  describe("changeConfigCallback", () => {
    afterEach(() => {
      sanbox.restore();
    });
    it("happy", () => {
      const stub = sanbox.stub(configMgr, "loadConfigs").returns();
      configMgr.changeConfigCallback({ affectsConfiguration: () => true });
      chai.assert.isTrue(stub.called);
    });
  });
  describe("loadConfigs", () => {
    beforeEach(async () => {
      sanbox.stub(ExtTelemetry, "sendTelemetryEvent");
      sanbox.stub(vscode.workspace, "getConfiguration").returns({
        get: () => {
          return "test";
        },
      } as any);
    });
    afterEach(() => {
      sanbox.restore();
    });
    it("happy", () => {
      const stub = sanbox.stub(configMgr, "loadLogLevel").returns();
      const stub2 = sanbox.stub(configMgr, "loadFeatureFlags").returns();
      configMgr.loadConfigs();
      chai.assert.isTrue(stub.called);
      chai.assert.isTrue(stub2.called);
    });
  });

  describe("loadFeatureFlags", () => {
    afterEach(() => {
      sanbox.restore();
    });
    it("happy", () => {
      const stub = sanbox.stub(configMgr, "getConfiguration").returns(false);
      configMgr.loadFeatureFlags();
      chai.assert.isTrue(stub.called);
    });
  });

  describe("registerConfigChangeCallback", () => {
    afterEach(() => {
      sanbox.restore();
    });
    it("happy", () => {
      const stub = sanbox.stub(configMgr, "loadConfigs").returns();
      configMgr.registerConfigChangeCallback();
      chai.assert.isTrue(stub.called);
    });
  });
});
