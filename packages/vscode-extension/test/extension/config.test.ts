import { err, LogLevel, ok, UserError } from "@microsoft/teamsfx-api";
import * as chai from "chai";
import * as sinon from "sinon";
import * as vscode from "vscode";
import VsCodeLogInstance from "../../src/commonlib/log";
import { configMgr } from "../../src/config";
import { ExtTelemetry } from "../../src/telemetry/extTelemetry";
import * as vsc_ui from "../../src/qm/vsc_ui";
import * as lifecycleHandlers from "../../src/handlers/lifecycleHandlers";

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

  describe("checkKiotaInstallation", async () => {
    afterEach(() => {
      sanbox.restore();
    });
    beforeEach(() => {
      sanbox.stub(vsc_ui, "VS_CODE_UI").value(new vsc_ui.VsCodeUI(<vscode.ExtensionContext>{}));
    });
    it("should skip if current value is enabled", async () => {
      const configStub = sanbox.stub(vscode.workspace, "getConfiguration").returns({
        get: () => {
          return "Enabled";
        },
        update: () => {
          return;
        },
      } as any);
      await configMgr.checkKiotaInstallation();
      chai.assert.isTrue(configStub.calledOnce);
    });
    it("should skip if kiota not installed", async () => {
      sanbox.stub(lifecycleHandlers, "validateKiotaInstallation").returns(false);
      const configStub = sanbox.stub(vscode.workspace, "getConfiguration").returns({
        get: (key: string) => {
          if (key === "enableMicrosoftKiota") {
            return true;
          } else {
            return "Undefined";
          }
        },
        update: (key: string, value: string) => {
          chai.assert.equal(key, "enableMicrosoftKiotaString");
          chai.assert.equal(value, "Enabled");
          return;
        },
      } as any);
      await configMgr.checkKiotaInstallation();
      chai.assert.isTrue(configStub.calledOnce);
    });
    it("should set enabled if previous value is true", async () => {
      sanbox.stub(vscode.workspace, "getConfiguration").returns({
        get: (key: string) => {
          if (key === "enableMicrosoftKiota") {
            return true;
          } else {
            return "Undefined";
          }
        },
        update: (key: string, value: string) => {
          chai.assert.equal(key, "enableMicrosoftKiotaString");
          chai.assert.equal(value, "Enabled");
          return;
        },
      } as any);
      sanbox.stub(lifecycleHandlers, "validateKiotaInstallation").returns(true);
      await configMgr.checkKiotaInstallation();
    });
    it("should ask user and set enabled", async () => {
      sanbox.stub(vscode.workspace, "getConfiguration").returns({
        get: (key: string) => {
          if (key === "enableMicrosoftKiota") {
            return undefined;
          } else {
            return "Undefined";
          }
        },
        update: (key: string, value: string) => {
          chai.assert.equal(key, "enableMicrosoftKiotaString");
          chai.assert.equal(value, "Enabled");
          return;
        },
      } as any);
      sanbox.stub(lifecycleHandlers, "validateKiotaInstallation").returns(true);
      sanbox.stub(vsc_ui.VS_CODE_UI, "showMessage").resolves(ok("Yes"));
      await configMgr.checkKiotaInstallation();
    });
    it("should ask user and set disabled", async () => {
      sanbox.stub(vscode.workspace, "getConfiguration").returns({
        get: (key: string) => {
          if (key === "enableMicrosoftKiota") {
            return undefined;
          } else {
            return "Undefined";
          }
        },
        update: (key: string, value: string) => {
          chai.assert.equal(key, "enableMicrosoftKiotaString");
          chai.assert.equal(value, "Disabled");
          return;
        },
      } as any);
      sanbox.stub(lifecycleHandlers, "validateKiotaInstallation").returns(true);
      sanbox.stub(vsc_ui.VS_CODE_UI, "showMessage").resolves(ok("No"));
      await configMgr.checkKiotaInstallation();
    });
    it("should ask user and set disabled if user cancel", async () => {
      sanbox.stub(vscode.workspace, "getConfiguration").returns({
        get: (key: string) => {
          if (key === "enableMicrosoftKiota") {
            return undefined;
          } else {
            return "Undefined";
          }
        },
        update: (key: string, value: string) => {
          chai.assert.equal(key, "enableMicrosoftKiotaString");
          chai.assert.equal(value, "Disabled");
          return;
        },
      } as any);
      sanbox.stub(lifecycleHandlers, "validateKiotaInstallation").returns(true);
      sanbox
        .stub(vsc_ui.VS_CODE_UI, "showMessage")
        .resolves(err(new UserError("source", "errorcode", "errormessage")));
      await configMgr.checkKiotaInstallation();
    });
  });
});
